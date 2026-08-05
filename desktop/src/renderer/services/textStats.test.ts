import { describe, expect, test } from 'bun:test'
import { countCharacters, countWords } from './textStats.js'

describe('countWords', () => {
  test('counts plain ASCII words', () => {
    expect(countWords('one two three')).toBe(3)
  })

  test('is empty for empty and whitespace-only input', () => {
    expect(countWords('')).toBe(0)
    expect(countWords('   ')).toBe(0)
    expect(countWords('\n\t \r\n')).toBe(0)
  })

  test('collapses runs of whitespace and line breaks', () => {
    expect(countWords('  one\t\ttwo\n\nthree  ')).toBe(3)
  })

  test('ignores punctuation between words', () => {
    expect(countWords('Hello, world! Really?')).toBe(3)
    expect(countWords('...')).toBe(0)
    expect(countWords('— – ·')).toBe(0)
  })

  test('counts accented Vietnamese text as one word per token', () => {
    const greeting = 'Xin chào thế giới'
    expect(countWords(greeting.normalize('NFC'))).toBe(4)
    expect(countWords(greeting.normalize('NFD'))).toBe(4)
    const sentence = 'Chúng tôi đã sửa lỗi này rồi'
    expect(countWords(sentence.normalize('NFC'))).toBe(7)
    expect(countWords(sentence.normalize('NFD'))).toBe(7)
  })

  test('keeps a decomposed accent inside its word', () => {
    const word = 'Tiếng'
    expect(countWords(word.normalize('NFC'))).toBe(1)
    expect(countWords(word.normalize('NFD'))).toBe(1)
    expect(countWords('café'.normalize('NFD'))).toBe(1)
  })

  test('counts other accented and non-Latin scripts', () => {
    expect(countWords('Grüße über Straße')).toBe(3)
    expect(countWords('naïve café résumé')).toBe(3)
    expect(countWords('Ελληνικά κείμενο')).toBe(2)
    expect(countWords('Привет мир')).toBe(2)
    expect(countWords('שלום עולם')).toBe(2)
  })

  test('keeps apostrophes and hyphens inside a single word', () => {
    expect(countWords("don't")).toBe(1)
    expect(countWords('don’t')).toBe(1)
    expect(countWords('mother-in-law')).toBe(1)
    expect(countWords("it's a well-known fact")).toBe(4)
    expect(countWords("l'État d'urgence")).toBe(2)
  })

  test('does not let a leading or trailing apostrophe or hyphen start a word', () => {
    expect(countWords("'quoted'")).toBe(1)
    expect(countWords('- dash -')).toBe(1)
    expect(countWords('--')).toBe(0)
  })

  test('counts every CJK character as its own word', () => {
    expect(countWords('中文')).toBe(1)
    expect(countWords('中文 测试')).toBe(2)
    expect(countWords('日本語のテキスト')).toBe(1)
  })

  test('counts numbers and alphanumeric tokens', () => {
    expect(countWords('42')).toBe(1)
    expect(countWords('1 2 3')).toBe(3)
    expect(countWords('3.14')).toBe(2)
    expect(countWords('Q4 2025 revenue')).toBe(3)
    expect(countWords('COVID-19')).toBe(1)
  })

  test('counts across slide and paragraph breaks without fusing words', () => {
    expect(countWords('first slide\n\nsecond slide')).toBe(4)
  })
})

describe('countCharacters', () => {
  test('excludes all whitespace', () => {
    expect(countCharacters('one two')).toBe(6)
    expect(countCharacters(' a \t b \n c ')).toBe(3)
  })

  test('is zero for empty and whitespace-only input', () => {
    expect(countCharacters('')).toBe(0)
    expect(countCharacters('   \n\t')).toBe(0)
  })

  test('counts punctuation and digits', () => {
    expect(countCharacters('Hi!')).toBe(3)
    expect(countCharacters('3.14')).toBe(4)
  })

  test('counts an accented character once, composed or decomposed', () => {
    const word = 'chào'
    expect(countCharacters(word.normalize('NFC'))).toBe(4)
    expect(countCharacters(word.normalize('NFD'))).toBe(4)
    expect(countCharacters('Grüße')).toBe(5)
  })

  test('counts CJK characters', () => {
    expect(countCharacters('中文 测试')).toBe(4)
  })
})
