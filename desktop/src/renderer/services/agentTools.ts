// Spreadsheet tools the AI assistant may call. Pure over a minimal workbook
// interface so the A1 handling, caps, and write validation are unit-testable.
// Reads execute immediately; writes are validated here but only applied after
// the user approves the proposal in the panel.

export const READ_RANGE_CELL_CAP = 500
export const WRITE_CELLS_CAP = 50
export const CHART_SERIES_CAP = 8

export const CHART_TYPES = ['column', 'bar', 'pie', 'line', 'doughnut'] as const
export type ChartType = (typeof CHART_TYPES)[number]

export interface AgentWorkbookAccess {
  sheetInfo(): { sheetNames: string[]; activeSheet: number }
  usedRange(sheet: number): string | null
  rangeCells(
    sheet: number,
    range: string
  ): Array<Array<{ input: string; isFormula: boolean; filterText?: string }>>
  editCells(sheet: number, edits: Array<{ row: number; col: number; input: string }>): unknown
  addChart(args: {
    sheet: number
    chartType: ChartType
    title?: string
    anchor: string
    categories?: string
    series: Array<{ name?: string; values: string }>
  }): unknown
}

export interface WriteCellsProposal {
  sheet: number
  sheetName: string
  edits: Array<{ a1: string; row: number; col: number; input: string }>
}

export interface CreateChartProposal {
  sheet: number
  sheetName: string
  chartType: ChartType
  title?: string
  anchor: string
  categories?: string
  series: Array<{ name?: string; values: string }>
}

const A1_RANGE = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/

function columnNumber(letters: string): number {
  let value = 0
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64)
  return value
}

const A1_CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/

/** A sheet-qualified reference rebuilt from its parts, requoting when needed. */
function joinSheetQualifier(parts: { sheet?: string; range: string }): string {
  if (parts.sheet === undefined) return parts.range
  const plain = /^[A-Za-z0-9_]+$/.test(parts.sheet)
  const sheet = plain ? parts.sheet : `'${parts.sheet.replace(/'/g, "''")}'`
  return `${sheet}!${parts.range}`
}

/** Split an optional sheet qualifier off a range reference; null on malformed quoting. */
export function splitSheetQualifier(
  reference: string
): { sheet?: string; range: string } | null {
  const trimmed = reference.trim()
  const bang = trimmed.lastIndexOf('!')
  if (bang === -1) return { range: trimmed }
  const rawSheet = trimmed.slice(0, bang)
  const range = trimmed.slice(bang + 1).trim()
  if (rawSheet === '') return null
  if (rawSheet.startsWith("'")) {
    if (!rawSheet.endsWith("'") || rawSheet.length < 2) return null
    return { sheet: rawSheet.slice(1, -1).replace(/''/g, "'"), range }
  }
  return { sheet: rawSheet, range }
}

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

/**
 * Validate a create_chart request into a reviewable proposal, or an error the
 * model can act on. Never mutates — application happens after user approval.
 */
export function validateCreateChart(
  access: Pick<AgentWorkbookAccess, 'sheetInfo'>,
  args: Record<string, unknown>
): { proposal: CreateChartProposal } | { error: string } {
  const info = access.sheetInfo()
  const requestedSheet =
    typeof args['sheet'] === 'number' && Number.isInteger(args['sheet'])
      ? args['sheet']
      : info.activeSheet
  if (requestedSheet < 0 || requestedSheet >= info.sheetNames.length) {
    return { error: `sheet ${requestedSheet} out of range (0..${info.sheetNames.length - 1})` }
  }
  const chartType = typeof args['chart_type'] === 'string' ? args['chart_type'] : ''
  if (!(CHART_TYPES as readonly string[]).includes(chartType)) {
    return { error: `chart_type must be one of ${CHART_TYPES.join(', ')}; got "${chartType}"` }
  }
  const sheetIndexByName = (name: string): number =>
    info.sheetNames.findIndex((candidate) => candidate.toLowerCase() === name.toLowerCase())
  const rawAnchor = typeof args['anchor'] === 'string' ? args['anchor'] : ''
  const anchorParts = splitSheetQualifier(rawAnchor)
  const anchor = anchorParts ? anchorParts.range.toUpperCase() : ''
  if (rangeCellCount(anchor) === null || !anchor.includes(':')) {
    return { error: `anchor must be an A1 rectangle like "D2:K16"; got "${rawAnchor}"` }
  }
  let sheet = requestedSheet
  if (anchorParts?.sheet !== undefined) {
    const anchorSheet = sheetIndexByName(anchorParts.sheet)
    if (anchorSheet === -1) {
      return { error: `anchor names unknown sheet "${anchorParts.sheet}"` }
    }
    if (typeof args['sheet'] === 'number' && anchorSheet !== sheet) {
      return {
        error: `anchor sheet "${anchorParts.sheet}" conflicts with sheet index ${sheet}`,
      }
    }
    sheet = anchorSheet
  }
  const rawSeries = Array.isArray(args['series']) ? args['series'] : null
  if (!rawSeries || rawSeries.length === 0) {
    return { error: 'series must be a non-empty array of { name?, values }' }
  }
  if (rawSeries.length > CHART_SERIES_CAP) {
    return { error: `${rawSeries.length} series exceed the cap of ${CHART_SERIES_CAP}` }
  }
  if ((chartType === 'pie' || chartType === 'doughnut') && rawSeries.length > 1) {
    return { error: `a ${chartType} chart takes exactly one series` }
  }
  const series: CreateChartProposal['series'] = []
  for (const entry of rawSeries) {
    const record = typeof entry === 'object' && entry !== null ? (entry as Record<string, unknown>) : null
    const values = typeof record?.['values'] === 'string' ? record['values'].trim() : ''
    const valuesParts = splitSheetQualifier(values)
    if (valuesParts === null || rangeCellCount(valuesParts.range) === null) {
      return { error: `each series needs an A1 values range; got ${JSON.stringify(entry)}` }
    }
    if (valuesParts.sheet !== undefined && sheetIndexByName(valuesParts.sheet) === -1) {
      return { error: `series range names unknown sheet "${valuesParts.sheet}"` }
    }
    const normalized = joinSheetQualifier(valuesParts)
    const name = typeof record?.['name'] === 'string' ? record['name'] : undefined
    series.push(name === undefined ? { values: normalized } : { name, values: normalized })
  }
  let categories = typeof args['categories'] === 'string' ? args['categories'].trim() : undefined
  if (categories !== undefined) {
    const parts = splitSheetQualifier(categories)
    if (parts === null || rangeCellCount(parts.range) === null) {
      return { error: `categories must be an A1 range; got "${categories}"` }
    }
    if (parts.sheet !== undefined && sheetIndexByName(parts.sheet) === -1) {
      return { error: `categories range names unknown sheet "${parts.sheet}"` }
    }
    categories = joinSheetQualifier(parts)
  }
  const title = typeof args['title'] === 'string' && args['title'].trim() !== '' ? args['title'].trim() : undefined
  const proposal: CreateChartProposal = {
    sheet,
    sheetName: info.sheetNames[sheet]!,
    chartType: chartType as ChartType,
    anchor,
    series,
  }
  if (title !== undefined) proposal.title = title
  if (categories !== undefined) proposal.categories = categories
  return { proposal }
}

/** Apply a user-approved chart proposal; the engine re-validates ranges. */
export function applyCreateChart(
  access: AgentWorkbookAccess,
  proposal: CreateChartProposal
): unknown {
  access.addChart({
    sheet: proposal.sheet,
    chartType: proposal.chartType,
    ...(proposal.title !== undefined ? { title: proposal.title } : {}),
    anchor: proposal.anchor,
    ...(proposal.categories !== undefined ? { categories: proposal.categories } : {}),
    series: proposal.series,
  })
  return {
    applied: true,
    sheet: proposal.sheet,
    chartType: proposal.chartType,
    anchor: proposal.anchor,
  }
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
