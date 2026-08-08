// Spreadsheet tools the AI assistant may call. Pure over a minimal workbook
// interface so the A1 handling, caps, and write validation are unit-testable.
// Reads execute immediately; writes are validated here but only applied after
// the user approves the proposal in the panel.

export const READ_RANGE_CELL_CAP = 500
export const WRITE_CELLS_CAP = 50

export interface AgentWorkbookAccess {
  sheetInfo(): { sheetNames: string[]; activeSheet: number }
  usedRange(sheet: number): string | null
  rangeCells(
    sheet: number,
    range: string
  ): Array<Array<{ input: string; isFormula: boolean; filterText?: string }>>
  editCells(sheet: number, edits: Array<{ row: number; col: number; input: string }>): unknown
}

export interface WriteCellsProposal {
  sheet: number
  sheetName: string
  edits: Array<{ a1: string; row: number; col: number; input: string }>
}

const A1_RANGE = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/

function columnNumber(letters: string): number {
  let value = 0
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
}

const A1_CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/

/** 0-based row/col of a single A1 cell, or null when it does not parse. */
export function a1ToRowCol(a1: string): { row: number; col: number } | null {
  const match = A1_CELL.exec(a1.trim())
  if (!match) return null
  return { row: Number(match[2]) - 1, col: columnNumber(match[1]!) - 1 }
}

/**
 * Validate a write_cells request into an applicable proposal, or an error the
 * model can act on. Never mutates — application happens after user approval.
 */
export function validateWriteCells(
  access: Pick<AgentWorkbookAccess, 'sheetInfo'>,
  args: Record<string, unknown>
): { proposal: WriteCellsProposal } | { error: string } {
  const info = access.sheetInfo()
  const sheet =
    typeof args['sheet'] === 'number' && Number.isInteger(args['sheet'])
      ? args['sheet']
      : info.activeSheet
  if (sheet < 0 || sheet >= info.sheetNames.length) {
    return { error: `sheet ${sheet} out of range (0..${info.sheetNames.length - 1})` }
  }
  const rawEdits = Array.isArray(args['edits']) ? args['edits'] : null
  if (!rawEdits || rawEdits.length === 0) {
    return { error: 'edits must be a non-empty array of { a1, input }' }
  }
  if (rawEdits.length > WRITE_CELLS_CAP) {
    return { error: `${rawEdits.length} edits exceed the cap of ${WRITE_CELLS_CAP} per call — split the change` }
  }
  const edits: WriteCellsProposal['edits'] = []
  for (const entry of rawEdits) {
    const record = typeof entry === 'object' && entry !== null ? (entry as Record<string, unknown>) : null
    const a1 = typeof record?.['a1'] === 'string' ? record['a1'].trim() : ''
    const input = typeof record?.['input'] === 'string' ? record['input'] : null
    const cell = a1ToRowCol(a1)
    if (!cell || input === null) {
      return { error: `each edit needs an A1 cell and a string input; got ${JSON.stringify(entry)}` }
    }
    edits.push({ a1: a1.toUpperCase(), row: cell.row, col: cell.col, input })
  }
  return { proposal: { sheet, sheetName: info.sheetNames[sheet]!, edits } }
}

/**
 * Apply an approved proposal as one batch (one undo step), then read the
 * written cells back so the model sees computed results — a formula that
 * evaluated to #NAME? is visible immediately instead of on the next read.
 */
export function applyWriteCells(access: AgentWorkbookAccess, proposal: WriteCellsProposal): unknown {
  access.editCells(
    proposal.sheet,
    proposal.edits.map(({ row, col, input }) => ({ row, col, input }))
  )
  const cells = proposal.edits.map((edit) => {
    const readBack = access.rangeCells(proposal.sheet, edit.a1)[0]?.[0]
    return { a1: edit.a1, value: readBack ? (readBack.filterText ?? readBack.input) : '' }
  })
  return { applied: true, sheet: proposal.sheet, cells }
}

/** Cell count of an A1 rectangle, or null when it does not parse. */
export function rangeCellCount(range: string): number | null {
  const match = A1_RANGE.exec(range.trim())
  if (!match) return null
  const [, startCol, startRow, endCol, endRow] = match
  const c0 = columnNumber(startCol!)
  const r0 = Number(startRow)
  const c1 = endCol ? columnNumber(endCol) : c0
  const r1 = endRow ? Number(endRow) : r0
  return (Math.abs(r1 - r0) + 1) * (Math.abs(c1 - c0) + 1)
}

export function executeXlsxAgentTool(
  access: AgentWorkbookAccess,
  name: string,
  args: Record<string, unknown>
): unknown {
  switch (name) {
    case 'list_sheets': {
      const info = access.sheetInfo()
      return info.sheetNames.map((sheetName, index) => ({
        index,
        name: sheetName,
        active: index === info.activeSheet,
        usedRange: access.usedRange(index),
      }))
    }
    case 'read_range': {
      const info = access.sheetInfo()
      const sheet =
        typeof args['sheet'] === 'number' && Number.isInteger(args['sheet'])
          ? args['sheet']
          : info.activeSheet
      if (sheet < 0 || sheet >= info.sheetNames.length) {
        return { error: `sheet ${sheet} out of range (0..${info.sheetNames.length - 1})` }
      }
      const range = typeof args['range'] === 'string' ? args['range'].trim() : ''
      const cells = rangeCellCount(range)
      if (cells === null) {
        return { error: `"${range}" is not an A1 range like "A1:F25"` }
      }
      if (cells > READ_RANGE_CELL_CAP) {
        return {
          error: `range has ${cells} cells; the cap is ${READ_RANGE_CELL_CAP} per call — read it in windows`,
        }
      }
      const grid = access.rangeCells(sheet, range)
      return {
        sheet,
        range,
        rows: grid.map((row) =>
          row.map(({ input, isFormula, filterText }) => ({
            input,
            isFormula,
            value: filterText ?? input,
          }))
        ),
      }
    }
    default:
      return { error: `unknown tool "${name}"` }
  }
}
