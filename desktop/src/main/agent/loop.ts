import type { AgentEvent } from '../../shared/ipc.js'
import type { ChatMessage, ChatToolCall, ChatToolFunction, ChatTurn } from './openaiStream.js'

export const MAX_TOOL_ROUNDS = 8

export interface AgentLoopDeps {
  /** One streamed model turn over the accumulated transcript. */
  callModel: (
    messages: ChatMessage[],
    tools: ChatToolFunction[],
    onTextDelta: (delta: string) => void
  ) => Promise<ChatTurn>
  /** Execute one tool call; resolves to a JSON-serializable result. */
  executeTool: (name: string, args: Record<string, unknown>) => Promise<unknown>
  emit: (event: AgentEvent) => void
  tools: ChatToolFunction[]
}

function parseArgs(call: ChatToolCall): Record<string, unknown> {
  try {
    const parsed: unknown = JSON.parse(call.function.arguments || '{}')
    return typeof parsed === 'object' && parsed !== null ? (parsed as Record<string, unknown>) : {}
  } catch {
    return {}
  }
}

/**
 * Text streams straight through to `emit`; each requested tool runs and its
 * result goes back as a `tool` message until the model stops asking (or the
 * round cap trips, which is surfaced to the model rather than thrown so it
 * can wrap up with what it has).
 */
export async function runAgentLoop(messages: ChatMessage[], deps: AgentLoopDeps): Promise<void> {
  const transcript = [...messages]
  for (let round = 0; round <= MAX_TOOL_ROUNDS + 1; round++) {
    const turn = await deps.callModel(transcript, deps.tools, (delta) =>
      deps.emit({ type: 'text', delta })
    )
    if (turn.toolCalls.length === 0) {
      deps.emit({ type: 'done' })
      return
    }
    transcript.push({
      role: 'assistant',
      content: turn.content,
      tool_calls: turn.toolCalls,
    })
    for (const call of turn.toolCalls) {
      const args = parseArgs(call)
      deps.emit({
        type: 'tool',
        name: call.function.name,
        summary: call.function.arguments.slice(0, 200),
      })
      let result: unknown
      if (round >= MAX_TOOL_ROUNDS) {
        result = { error: 'tool budget exhausted — answer with what you already know' }
      } else {
        try {
          result = await deps.executeTool(call.function.name, args)
        } catch (error) {
          result = { error: error instanceof Error ? error.message : String(error) }
        }
      }
      transcript.push({
        role: 'tool',
        tool_call_id: call.id,
        content: JSON.stringify(result ?? null),
      })
    }
  }
  deps.emit({ type: 'error', message: 'the model kept requesting tools past the budget' })
}
