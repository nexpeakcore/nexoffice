import type { DocumentKind } from '../../shared/ipc.js'
import type { ChatToolFunction } from './openaiStream.js'

/** Hard cap per read_range call; the model is told to window larger reads. */
export const READ_RANGE_CELL_CAP = 500

const XLSX_TOOLS: ChatToolFunction[] = [
  {
    name: 'list_sheets',
    description:
      'List the sheets of the open workbook: name, index, whether it is active, and its used range in A1 notation (null when empty).',
    parameters: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'read_range',
    description:
      `Read a rectangular range of cells from one sheet. Returns a row-major grid of { input, isFormula } where input is the raw cell content (formula text when isFormula). At most ${READ_RANGE_CELL_CAP} cells per call — window larger reads into consecutive calls.`,
    parameters: {
      type: 'object',
      properties: {
        sheet: {
          type: 'integer',
          description: 'Sheet index from list_sheets; defaults to the active sheet.',
        },
        range: {
          type: 'string',
          description: 'A1-notation rectangle, e.g. "A1:F25".',
        },
      },
      required: ['range'],
      additionalProperties: false,
    },
  },
]

export function toolsForDocument(kind: DocumentKind): ChatToolFunction[] {
  return kind === 'xlsx' ? XLSX_TOOLS : []
}

export function agentSystemPrompt(kind: DocumentKind, name: string, locale: string): string {
  const document =
    kind === 'xlsx'
      ? `an Excel workbook named "${name}". Use the tools to inspect it before answering questions about its contents; never guess cell values.`
      : `a document named "${name}".`
  return (
    'You are the NexOffice assistant, embedded in a desktop office suite. ' +
    `The user has ${document} ` +
    'You currently have read-only access: you cannot modify the document, only inspect it and answer. ' +
    'Be concise. Answer in the language the user writes in ' +
    `(their interface locale is ${locale}).`
  )
}
