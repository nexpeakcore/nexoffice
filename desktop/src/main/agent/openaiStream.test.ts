import { describe, expect, it } from 'bun:test'
import { sseData, streamChatTurn } from './openaiStream.js'

function bodyFromChunks(chunks: string[]): ReadableStream<Uint8Array> {
  const encoder = new TextEncoder()
  return new ReadableStream({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(encoder.encode(chunk))
      controller.close()
    },
  })
}

describe('sseData', () => {
  it('splits data lines across chunk boundaries and stops at [DONE]', async () => {
    const stream = bodyFromChunks(['data: {"a"', ':1}\n\ndata: {"b":2}\n', 'data: [DONE]\ndata: {"c":3}\n'])
    const seen: string[] = []
    for await (const payload of sseData(stream)) seen.push(payload)
    expect(seen).toEqual(['{"a":1}', '{"b":2}'])
  })
})

describe('streamChatTurn', () => {
  it('accumulates text and index-keyed tool-call fragments', async () => {
    const sse = [
      'data: {"choices":[{"delta":{"content":"Th"}}]}\n',
      'data: {"choices":[{"delta":{"content":"e sum"}}]}\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_9","function":{"name":"read_","arguments":"{\\"ran"}}]}}]}\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"range","arguments":"ge\\":\\"A1\\"}"}}]}}]}\n',
      'data: {"choices":[{"finish_reason":"tool_calls"}]}\n',
      'data: [DONE]\n',
    ]
    const deltas: string[] = []
    const result = await streamChatTurn(
      {
        baseUrl: 'https://api.test',
        apiKey: 'k',
        model: 'm',
        fetchImpl: async () =>
          new Response(bodyFromChunks(sse), { status: 200 }),
      },
      [{ role: 'user', content: 'hi' }],
      [],
      { onTextDelta: (delta) => deltas.push(delta) },
      new AbortController().signal
    )
    expect(deltas.join('')).toBe('The sum')
    expect(result.content).toBe('The sum')
    expect(result.finishReason).toBe('tool_calls')
    expect(result.toolCalls).toEqual([
      {
        id: 'call_9',
        type: 'function',
        function: { name: 'read_range', arguments: '{"range":"A1"}' },
      },
    ])
  })

  it('throws a readable error on a non-2xx response', async () => {
    await expect(
      streamChatTurn(
        {
          baseUrl: 'https://api.test',
          apiKey: 'k',
          model: 'm',
          fetchImpl: async () => new Response('invalid key', { status: 401 }),
        },
        [],
        [],
        {},
        new AbortController().signal
      )
    ).rejects.toThrow(/401.*invalid key/s)
  })
})
