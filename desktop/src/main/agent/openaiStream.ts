// OpenAI-compatible chat-completions streaming, the dialect DeepSeek speaks.
// Kept free of Electron imports so the SSE parsing and delta accumulation are
// unit-testable with a fake fetch.

export interface ChatToolFunction {
  name: string
  description: string
  parameters: Record<string, unknown>
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool'
  content: string
  tool_calls?: ChatToolCall[]
  tool_call_id?: string
}

export interface ChatToolCall {
  id: string
  type: 'function'
  function: { name: string; arguments: string }
}

export interface ChatTurn {
  content: string
  toolCalls: ChatToolCall[]
  finishReason: string | null
}

export interface StreamCallbacks {
  onTextDelta?: (delta: string) => void
}

export interface ChatClientOptions {
  baseUrl: string
  apiKey: string
  model: string
  fetchImpl?: (input: string | URL, init?: RequestInit) => Promise<Response>
}

interface StreamedChoice {
  delta?: {
    content?: string | null
    tool_calls?: Array<{
      index: number
      id?: string
      function?: { name?: string; arguments?: string }
    }>
  }
  finish_reason?: string | null
}

/** Split an SSE byte stream into `data:` payloads; exported for tests. */
export async function* sseData(body: ReadableStream<Uint8Array>): AsyncGenerator<string> {
  const reader = body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      let boundary: number
      while ((boundary = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, boundary).replace(/\r$/, '')
        buffer = buffer.slice(boundary + 1)
        if (!line.startsWith('data:')) continue
        const payload = line.slice(5).trim()
        if (payload === '[DONE]') return
        if (payload) yield payload
      }
    }
  } finally {
    reader.releaseLock()
  }
}

/**
 * One streamed model turn. Tool-call fragments arrive as index-keyed deltas
 * (`function.arguments` accumulates across chunks) and are reassembled here.
 */
export async function streamChatTurn(
  options: ChatClientOptions,
  messages: ChatMessage[],
  tools: ChatToolFunction[],
  callbacks: StreamCallbacks,
  signal: AbortSignal
): Promise<ChatTurn> {
  const doFetch = options.fetchImpl ?? fetch
  const response = await doFetch(`${options.baseUrl}/chat/completions`, {
    method: 'POST',
    signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${options.apiKey}`,
    },
    body: JSON.stringify({
      model: options.model,
      messages,
      stream: true,
      ...(tools.length > 0
        ? { tools: tools.map((fn) => ({ type: 'function', function: fn })) }
        : {}),
    }),
  })
  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(`model request failed (${response.status}): ${detail.slice(0, 300)}`)
  }
  if (!response.body) throw new Error('model response has no body')

  let content = ''
  let finishReason: string | null = null
  const toolCalls = new Map<number, ChatToolCall>()

  for await (const payload of sseData(response.body)) {
    let parsed: { choices?: StreamedChoice[] }
    try {
      parsed = JSON.parse(payload) as { choices?: StreamedChoice[] }
    } catch {
      continue
    }
    const choice = parsed.choices?.[0]
    if (!choice) continue
    if (choice.delta?.content) {
      content += choice.delta.content
      callbacks.onTextDelta?.(choice.delta.content)
    }
    for (const fragment of choice.delta?.tool_calls ?? []) {
      const existing = toolCalls.get(fragment.index) ?? {
        id: fragment.id ?? `call_${fragment.index}`,
        type: 'function' as const,
        function: { name: '', arguments: '' },
      }
      if (fragment.id) existing.id = fragment.id
      if (fragment.function?.name) existing.function.name += fragment.function.name
      if (fragment.function?.arguments) existing.function.arguments += fragment.function.arguments
      toolCalls.set(fragment.index, existing)
    }
    if (choice.finish_reason) finishReason = choice.finish_reason
  }

  return {
    content,
    toolCalls: [...toolCalls.entries()].sort(([a], [b]) => a - b).map(([, call]) => call),
    finishReason,
  }
}
