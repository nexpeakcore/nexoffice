// Read-only spreadsheet tools the AI assistant may call. Pure over a minimal
// workbook interface so the A1 handling and caps are unit-testable.

export const READ_RANGE_CELL_CAP = 500

export interface AgentWorkbookAccess {
  sheetInfo(): { sheetNames: string[]; activeSheet: number }
  usedRange(sheet: number): string | null
  rangeCells(sheet: number, range: string): Array<Array<{ input: string; isFormula: boolean }>>
}

const A1_RANGE = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/

function columnNumber(letters: string): number {
  let value = 0
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
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
        rows: grid.map((row) => row.map(({ input, isFormula }) => ({ input, isFormula }))),
      }
    }
    default:
      return { error: `unknown tool "${name}"` }
  }
}
