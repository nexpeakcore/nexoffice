import type { DocumentKind } from '../../shared/ipc.js'
import type { ChatToolFunction } from './openaiStream.js'

/** Hard cap per read_range call; the model is told to window larger reads. */
export const READ_RANGE_CELL_CAP = 500
/** Hard cap per write_cells call; keeps proposals reviewable at a glance. */
export const WRITE_CELLS_CAP = 50

/** Tools that wait on user approval get a human-scale timeout. */
export const APPROVAL_TOOLS = new Set(['write_cells', 'create_chart'])

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
      `Read a rectangular range of cells from one sheet. Returns a row-major grid of { input, isFormula, value } — input is the raw content (formula text when isFormula), value is the current computed result (formula errors appear here as #NAME?, #VALUE!, …). At most ${READ_RANGE_CELL_CAP} cells per call — window larger reads into consecutive calls.`,
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
  {
    name: 'write_cells',
    description:
      `Propose edits to cells. The user sees the full list and must approve before anything is applied — the result tells you whether they approved, and includes each written cell's computed value so you can catch formula errors and fix them in a follow-up proposal. Each edit sets one cell's raw input (text, number, or a formula starting with "="); an empty input clears the cell. At most ${WRITE_CELLS_CAP} edits per call. Read the affected cells first so you do not overwrite data the user wants kept.`,
    parameters: {
      type: 'object',
      properties: {
        sheet: {
          type: 'integer',
          description: 'Sheet index from list_sheets; defaults to the active sheet.',
        },
        edits: {
          type: 'array',
          description: 'Cells to set.',
          items: {
            type: 'object',
            properties: {
              a1: { type: 'string', description: 'Single cell in A1 notation, e.g. "B7".' },
              input: {
                type: 'string',
                description: 'Raw cell input: text, number, or "=FORMULA". Empty clears.',
              },
            },
            required: ['a1', 'input'],
            additionalProperties: false,
          },
        },
      },
      required: ['edits'],
      additionalProperties: false,
    },
  },
  {
    name: 'create_chart',
    description:
      'Create a chart floating over the grid from cell ranges. The user sees the proposal and must approve it. The chart tracks its source ranges live — when the cells change, the chart re-renders. Use read_range first to find the data. Prefer a categories range (labels) plus one or more numeric series ranges of the same length.',
    parameters: {
      type: 'object',
      properties: {
        sheet: {
          type: 'integer',
          description:
            '0-based sheet index from list_sheets the chart is placed on; defaults to the active sheet.',
        },
        chart_type: {
          type: 'string',
          enum: ['column', 'bar', 'pie', 'line', 'doughnut'],
          description: 'Chart family. pie/doughnut take exactly one series.',
        },
        title: { type: 'string', description: 'Optional chart title.' },
        anchor: {
          type: 'string',
          description:
            'A1 rectangle of cells the chart floats over (position and size), e.g. "E2:L18". Pick an empty area beside the data.',
        },
        categories: {
          type: 'string',
          description:
            'A1 range of the category labels, e.g. "A2:A10". May be sheet-qualified ("Data!A2:A10") to read another sheet.',
        },
        series: {
          type: 'array',
          description: 'Data series to plot.',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Legend label for the series.' },
              values: {
                type: 'string',
                description:
                  'A1 range of numeric values, e.g. "C2:C10". May be sheet-qualified ("Data!C2:C10").',
              },
            },
            required: ['values'],
            additionalProperties: false,
          },
        },
      },
      required: ['chart_type', 'anchor', 'series'],
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
      ? `an Excel workbook named "${name}". Use the tools to inspect it before answering questions about its contents; never guess cell values. ` +
        'You may change cells with write_cells: every proposal is shown to the user for approval, so group related edits into one call and explain what you are about to change before calling it. A rejected proposal is not an error — ask what to do differently. ' +
        'When the user wants a result that should stay current as the data changes (totals, rankings, top-N blocks, lookups), write FORMULAS rather than computed constants — e.g. a live top-10 of column C is =LARGE(C$2:C$100,1)…=LARGE(C$2:C$100,10) with =INDEX(A:A,MATCH(LARGE(…),C:C,0)) for the labels. ' +
        'The formula engine supports the common function set (SUM, AVERAGE, COUNT/COUNTIF(S), SUMIF(S), IF/IFS/IFERROR, MIN/MAX, MEDIAN, ROUND, LARGE, SMALL, RANK, INDEX, MATCH, VLOOKUP/HLOOKUP/XLOOKUP, CHOOSE, text and date functions) plus the dynamic arrays SORT, SORTBY, FILTER, UNIQUE and SEQUENCE, which spill from the cell holding the formula into the cells below and to the right of it. ' +
        'Two rules about them: the cells a result needs must be empty or the whole result is #SPILL!, so anchor one where there is room; and an omitted argument between two commas does not parse yet, so write =SORT(A1:B9,1,1,TRUE), never =SORT(A1:B9,,,TRUE). ' +
        'You may create charts with create_chart (column, bar, pie, line, doughnut): the chart reads its ranges live, so it stays current as the data changes. Place the anchor over empty cells so it does not cover the data.'
      : `a document named "${name}".`
  return (
    'You are the NexOffice assistant, embedded in a desktop office suite. ' +
    `The user has ${document} ` +
    'Be concise. Answer in the language the user writes in ' +
    `(their interface locale is ${locale}).`
  )
}
