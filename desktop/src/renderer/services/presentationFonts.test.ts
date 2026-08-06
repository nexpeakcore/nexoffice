import { describe, expect, test } from 'bun:test'
import type { DeckSnapshot, ShapeSnapshot } from '@betteroffice/pptx'
import {
  MAX_PRESENTATION_FONT_FACES,
  baseFontRequests,
  collectFontRequests,
} from './presentationFonts.js'

function run(fontFamily: string | null, bold = false, italic = false): unknown {
  return { text: 'x', style: { fontFamily, bold, italic } }
}

function shape(runs: unknown[], children: unknown[] = []): ShapeSnapshot {
  return {
    textStories: [{ paragraphs: [{ runs }] }],
    children,
  } as unknown as ShapeSnapshot
}

function deck(shapes: unknown[]): DeckSnapshot {
  return { slides: [{ shapes }] } as unknown as DeckSnapshot
}

const families = (snapshot: DeckSnapshot): string[] =>
  collectFontRequests(snapshot).map((request) => request.family)

describe('collectFontRequests', () => {
  test('always carries the base faces, whatever the deck names', () => {
    const collected = collectFontRequests(deck([]))
    for (const request of baseFontRequests()) {
      expect(collected).toContainEqual(request)
    }
  })

  // The defect this exists for: the editor loaded the base set alone, so a
  // deck naming a CJK family had its Chinese measured with a Latin face while
  // the PDF measured it with the right one — the same run 33% wider on paper.
  test('adds a family the deck names but the base set does not carry', () => {
    expect(families(deck([shape([run('SimSun')])]))).toContain('SimSun')
  })

  test('reaches families named inside nested groups', () => {
    const nested = shape([], [shape([], [shape([run('Meiryo')])])])
    expect(families(deck([nested]))).toContain('Meiryo')
  })

  test('asks for each style of a family separately', () => {
    const collected = collectFontRequests(
      deck([shape([run('SimSun'), run('SimSun', true), run('SimSun', false, true)])]),
    )
    const simsun = collected.filter((request) => request.family === 'SimSun')
    expect(simsun).toEqual([
      { family: 'SimSun', bold: false, italic: false },
      { family: 'SimSun', bold: true, italic: false },
      { family: 'SimSun', bold: false, italic: true },
    ])
  })

  test('names the same family once however it is cased', () => {
    const collected = collectFontRequests(
      deck([shape([run('SimSun'), run('simsun'), run('SIMSUN')])]),
    )
    expect(collected.filter((request) => request.family.toLowerCase() === 'simsun')).toHaveLength(1)
  })

  // A theme reference is resolved by the engine itself, and a blank names
  // nothing; registering a face under either would shadow the real one.
  test('skips theme references and blanks', () => {
    const collected = families(deck([shape([run('+mj-lt'), run('+mn-lt'), run('  '), run(null)])]))
    expect(collected).not.toContain('+mj-lt')
    expect(collected).not.toContain('+mn-lt')
    expect(collected.some((family) => family.trim() === '')).toBe(false)
  })

  test('trims a family the deck padded', () => {
    expect(families(deck([shape([run('  SimSun  ')])]))).toContain('SimSun')
  })

  // The engine refuses registrations past its own ceiling, so a deck naming
  // thousands of families must not push the base faces out or overrun it.
  test('stops at the engine ceiling and keeps the base faces', () => {
    const many = Array.from({ length: 1_000 }, (_, index) => run(`Family ${index}`))
    const collected = collectFontRequests(deck([shape(many)]))
    expect(collected.length).toBeLessThanOrEqual(MAX_PRESENTATION_FONT_FACES)
    for (const request of baseFontRequests()) {
      expect(collected).toContainEqual(request)
    }
  })
})
