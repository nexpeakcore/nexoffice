import { describe, expect, it } from 'bun:test'
import type { AgentEvent } from '../../shared/ipc.js'
import { MAX_TOOL_ROUNDS, runAgentLoop } from './loop.js'
import type { ChatMessage, ChatTurn } from './openaiStream.js'

function turn(partial: Partial<ChatTurn>): ChatTurn {
  return { content: '', toolCalls: [], finishReason: 'stop', ...partial }
}

describe('runAgentLoop', () => {
  it('streams text and finishes without tools', async () => {
    const events: AgentEvent[] = []
    await runAgentLoop([{ role: 'user', content: 'hi' }], {
      tools: [],
      callModel: async (_messages, _tools, onTextDelta) => {
        onTextDelta('hel')
        onTextDelta('lo')
        return turn({ content: 'hello' })
      },
      executeTool: async () => ({}),
      emit: (event) => events.push(event),
    })
    expect(events).toEqual([
      { type: 'text', delta: 'hel' },
      { type: 'text', delta: 'lo' },
      { type: 'done' },
    ])
  })

  it('round-trips tool calls and feeds results back as tool messages', async () => {
    const seen: ChatMessage[][] = []
    const events: AgentEvent[] = []
    let calls = 0
    await runAgentLoop([{ role: 'user', content: 'sum column B' }], {
      tools: [],
      callModel: async (messages) => {
        seen.push([...messages])
        calls += 1
        if (calls === 1) {
          return turn({
            finishReason: 'tool_calls',
            toolCalls: [
              {
                id: 'call_1',
                type: 'function',
                function: { name: 'read_range', arguments: '{"range":"B1:B3"}' },
              },
            ],
          })
        }
        return turn({ content: 'the sum is 6' })
      },
      executeTool: async (name, args) => {
        expect(name).toBe('read_range')
        expect(args).toEqual({ range: 'B1:B3' })
        return { rows: [[{ input: '1' }], [{ input: '2' }], [{ input: '3' }]] }
      },
      emit: (event) => events.push(event),
    })
    const finalTranscript = seen.at(-1)!
    expect(finalTranscript.at(-1)?.role).toBe('tool')
    expect(finalTranscript.at(-1)?.tool_call_id).toBe('call_1')
    expect(events.some((event) => event.type === 'tool' && event.name === 'read_range')).toBe(true)
    expect(events.at(-1)).toEqual({ type: 'done' })
  })

  it('surfaces tool failures to the model instead of throwing', async () => {
    let calls = 0
    let toolPayload: string | undefined
    await runAgentLoop([{ role: 'user', content: 'x' }], {
      tools: [],
      callModel: async (messages) => {
        calls += 1
        toolPayload = messages.at(-1)?.role === 'tool' ? messages.at(-1)?.content : toolPayload
        if (calls === 1) {
          return turn({
            toolCalls: [
              { id: 'c', type: 'function', function: { name: 'boom', arguments: '{}' } },
            ],
          })
        }
        return turn({ content: 'ok' })
      },
      executeTool: async () => {
        throw new Error('exploded')
      },
      emit: () => {},
    })
    expect(toolPayload).toContain('exploded')
  })

  it('stops a model that never stops asking for tools', async () => {
    const events: AgentEvent[] = []
    let toolExecutions = 0
    await runAgentLoop([{ role: 'user', content: 'x' }], {
      tools: [],
      callModel: async () =>
        turn({
          toolCalls: [{ id: 'c', type: 'function', function: { name: 'spin', arguments: '{}' } }],
        }),
      executeTool: async () => {
        toolExecutions += 1
        return {}
      },
      emit: (event) => events.push(event),
    })
    expect(events.at(-1)?.type).toBe('error')
    expect(toolExecutions).toBeLessThanOrEqual(MAX_TOOL_ROUNDS)
  })
})
