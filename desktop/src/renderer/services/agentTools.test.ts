import { describe, expect, it } from 'bun:test'
import {
  a1ToRowCol,
  applyWriteCells,
  executeXlsxAgentTool,
  rangeCellCount,
  READ_RANGE_CELL_CAP,
  validateWriteCells,
  WRITE_CELLS_CAP,
  type AgentWorkbookAccess,
  applyCreateChart,
  validateCreateChart,
} from './agentTools'

function fakeAccess(): AgentWorkbookAccess & {
  edited: Array<{ sheet: number; edits: unknown[] }>
  charts: unknown[]
} {
  const edited: Array<{ sheet: number; edits: unknown[] }> = []
  const charts: unknown[] = []
  return {
    edited,
    charts,
    editCells: (sheet, edits) => {
      edited.push({ sheet, edits })
      return {}
    },
    addChart: (args) => {
      charts.push(args)
      return {}
    },
    sheetInfo: () => ({ sheetNames: ['Budget', 'Summary'], activeSheet: 1 }),
    usedRange: (sheet) => (sheet === 0 ? 'A1:C9' : null),
    rangeCells: (_sheet, range) => {
      if (range === 'A1:B2') {
        return [
          [
            { input: '1', isFormula: false },
            { input: '=A1*2', isFormula: true, filterText: '2' },
          ],
          [
            { input: 'x', isFormula: false },
            { input: '', isFormula: false },
          ],
        ]
      }
      return [[{ input: '', isFormula: false }]]
    },
  }
}

describe('rangeCellCount', () => {
  it('measures rectangles and single cells', () => {
    expect(rangeCellCount('A1:F25')).toBe(150)
    expect(rangeCellCount('B2')).toBe(1)
    expect(rangeCellCount('$A$1:$B$2')).toBe(4)
  })

  it('rejects non-ranges', () => {
    expect(rangeCellCount('Sheet1!A1')).toBeNull()
    expect(rangeCellCount('1:5')).toBeNull()
    expect(rangeCellCount('')).toBeNull()
  })
})

describe('executeXlsxAgentTool', () => {
  it('lists sheets with used ranges', () => {
    expect(executeXlsxAgentTool(fakeAccess(), 'list_sheets', {})).toEqual([
      { index: 0, name: 'Budget', active: false, usedRange: 'A1:C9' },
      { index: 1, name: 'Summary', active: true, usedRange: null },
    ])
  })

  it('reads a range from an explicit sheet', () => {
    const result = executeXlsxAgentTool(fakeAccess(), 'read_range', {
      sheet: 0,
      range: 'A1:B2',
    }) as { rows: unknown[][] }
    expect(result.rows).toHaveLength(2)
    expect(result.rows[0]![1]).toEqual({ input: '=A1*2', isFormula: true, value: '2' })
  })

  it('defaults to the active sheet', () => {
    const result = executeXlsxAgentTool(fakeAccess(), 'read_range', { range: 'A1' }) as {
      sheet: number
    }
    expect(result.sheet).toBe(1)
  })

  it('refuses oversized ranges instead of reading them', () => {
    const result = executeXlsxAgentTool(fakeAccess(), 'read_range', { range: 'A1:Z100' }) as {
      error: string
    }
    expect(result.error).toContain(String(READ_RANGE_CELL_CAP))
  })

  it('parses A1 cells', () => {
    expect(a1ToRowCol('B7')).toEqual({ row: 6, col: 1 })
    expect(a1ToRowCol('$AA$10')).toEqual({ row: 9, col: 26 })
    expect(a1ToRowCol('B7:C9')).toBeNull()
    expect(a1ToRowCol('7B')).toBeNull()
  })

  it('validates write proposals without mutating', () => {
    const access = fakeAccess()
    const validated = validateWriteCells(access, {
      sheet: 0,
      edits: [
        { a1: 'b7', input: '=SUM(B1:B6)' },
        { a1: 'C1', input: '' },
      ],
    })
    expect(validated).toEqual({
      proposal: {
        sheet: 0,
        sheetName: 'Budget',
        edits: [
          { a1: 'B7', row: 6, col: 1, input: '=SUM(B1:B6)' },
          { a1: 'C1', row: 0, col: 2, input: '' },
        ],
      },
    })
    expect(access.edited).toHaveLength(0)
  })

  it('rejects oversized, malformed, and empty write requests', () => {
    const access = fakeAccess()
    const oversized = validateWriteCells(access, {
      edits: Array.from({ length: WRITE_CELLS_CAP + 1 }, (_, i) => ({ a1: `A${i + 1}`, input: 'x' })),
    })
    expect(oversized).toHaveProperty('error')
    expect(validateWriteCells(access, { edits: [] })).toHaveProperty('error')
    expect(validateWriteCells(access, { edits: [{ a1: 'A1:B2', input: 'x' }] })).toHaveProperty('error')
    expect(validateWriteCells(access, { edits: [{ a1: 'A1' }] })).toHaveProperty('error')
    expect(validateWriteCells(access, { sheet: 5, edits: [{ a1: 'A1', input: 'x' }] })).toHaveProperty('error')
  })

  it('applies an approved proposal as one batch', () => {
    const access = fakeAccess()
    const validated = validateWriteCells(access, {
      edits: [{ a1: 'A1', input: '42' }],
    })
    if (!('proposal' in validated)) throw new Error('expected a proposal')
    const result = applyWriteCells(access, validated.proposal)
    expect(access.edited).toEqual([{ sheet: 1, edits: [{ row: 0, col: 0, input: '42' }] }])
    expect(result).toEqual({ applied: true, sheet: 1, cells: [{ a1: 'A1', value: '' }] })
  })

  it('reports bad ranges, bad sheets, and unknown tools as model-readable errors', () => {
    expect(executeXlsxAgentTool(fakeAccess(), 'read_range', { range: 'nope' })).toHaveProperty(
      'error'
    )
    expect(
      executeXlsxAgentTool(fakeAccess(), 'read_range', { sheet: 9, range: 'A1' })
    ).toHaveProperty('error')
    expect(executeXlsxAgentTool(fakeAccess(), 'write_cell', {})).toHaveProperty('error')
  })

  it('validates create_chart specs into proposals and rejects bad ones', () => {
    const access = fakeAccess()
    const validated = validateCreateChart(access, {
      chart_type: 'column',
      title: 'Doanh thu',
      anchor: 'e2:l18',
      categories: 'A2:A10',
      series: [{ name: 'Revenue', values: 'C2:C10' }],
    })
    if (!('proposal' in validated)) throw new Error(`expected a proposal: ${JSON.stringify(validated)}`)
    expect(validated.proposal).toEqual({
      sheet: 1,
      sheetName: 'Summary',
      chartType: 'column',
      title: 'Doanh thu',
      anchor: 'E2:L18',
      categories: 'A2:A10',
      series: [{ name: 'Revenue', values: 'C2:C10' }],
    })

    expect(validateCreateChart(access, { chart_type: 'radar', anchor: 'A1:C3', series: [{ values: 'A1:A2' }] })).toHaveProperty('error')
    expect(validateCreateChart(access, { chart_type: 'pie', anchor: 'A1', series: [{ values: 'A1:A2' }] })).toHaveProperty('error')
    expect(validateCreateChart(access, { chart_type: 'pie', anchor: 'A1:C3', series: [] })).toHaveProperty('error')
    expect(
      validateCreateChart(access, {
        chart_type: 'pie',
        anchor: 'A1:C3',
        series: [{ values: 'A1:A2' }, { values: 'B1:B2' }],
      })
    ).toHaveProperty('error')
    expect(
      validateCreateChart(access, { chart_type: 'line', anchor: 'A1:C3', series: [{ values: 'nope' }] })
    ).toHaveProperty('error')
    expect(
      validateCreateChart(access, { sheet: 9, chart_type: 'line', anchor: 'A1:C3', series: [{ values: 'A1:A2' }] })
    ).toHaveProperty('error')
  })

  it('applies an approved chart proposal through the workbook access', () => {
    const access = fakeAccess()
    const validated = validateCreateChart(access, {
      chart_type: 'pie',
      anchor: 'D2:J14',
      series: [{ values: 'B2:B5' }],
    })
    if (!('proposal' in validated)) throw new Error('expected a proposal')
    const result = applyCreateChart(access, validated.proposal)
    expect(access.charts).toEqual([
      { sheet: 1, chartType: 'pie', anchor: 'D2:J14', series: [{ values: 'B2:B5' }] },
    ])
    expect(result).toEqual({ applied: true, sheet: 1, chartType: 'pie', anchor: 'D2:J14' })
  })

  it('accepts sheet-qualified chart ranges and strips the anchor qualifier', () => {
    const access = fakeAccess()
    const validated = validateCreateChart(access, {
      chart_type: 'line',
      anchor: 'Summary!D2:K16',
      categories: "'Budget'!A2:A10",
      series: [{ values: 'Budget!C2:C10' }],
    })
    if (!('proposal' in validated)) throw new Error(`expected a proposal: ${JSON.stringify(validated)}`)
    expect(validated.proposal.anchor).toBe('D2:K16')
    expect(validated.proposal.categories).toBe('Budget!A2:A10')
    expect(validated.proposal.series).toEqual([{ values: 'Budget!C2:C10' }])
    expect(
      validateCreateChart(access, {
        chart_type: 'line',
        anchor: 'A1:C3',
        series: [{ values: "'Broken!C2:C10" }],
      })
    ).toHaveProperty('error')
    const spaced = validateCreateChart(access, {
      chart_type: 'line',
      anchor: 'A1:C3',
      series: [{ values: 'Budget! C2:C10' }],
    })
    if (!('proposal' in spaced)) throw new Error(JSON.stringify(spaced))
    expect(spaced.proposal.series).toEqual([{ values: 'Budget!C2:C10' }])
  })

  it('routes the chart to the anchor-qualified sheet and rejects unknown names', () => {
    const access = fakeAccess()
    const validated = validateCreateChart(access, {
      chart_type: 'line',
      anchor: 'budget!D2:K16',
      series: [{ values: 'A1:A2' }],
    })
    if (!('proposal' in validated)) throw new Error(JSON.stringify(validated))
    expect(validated.proposal.sheet).toBe(0)
    expect(validated.proposal.sheetName).toBe('Budget')
    expect(
      validateCreateChart(access, {
        chart_type: 'line',
        sheet: 1,
        anchor: 'Budget!D2:K16',
        series: [{ values: 'A1:A2' }],
      })
    ).toHaveProperty('error')
    expect(
      validateCreateChart(access, {
        chart_type: 'line',
        anchor: 'Nowhere!D2:K16',
        series: [{ values: 'A1:A2' }],
      })
    ).toHaveProperty('error')
    expect(
      validateCreateChart(access, {
        chart_type: 'line',
        anchor: 'A1:C3',
        series: [{ values: 'Dtaa!C2:C10' }],
      })
    ).toHaveProperty('error')
    expect(
      validateCreateChart(access, {
        chart_type: 'line',
        anchor: 'A1:C3',
        categories: 'Nope!A1:A2',
        series: [{ values: 'A1:A2' }],
      })
    ).toHaveProperty('error')
  })
})
