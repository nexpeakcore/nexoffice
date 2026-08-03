import en from './en.json'
import de from './de.json'
import fr from './fr.json'
import he from './he.json'
import hi from './hi.json'
import id from './id.json'
import pl from './pl.json'
import ptBR from './pt-BR.json'
import tr from './tr.json'
import vi from './vi.json'
import zhCN from './zh-CN.json'

export const SUPPORTED_LOCALES = ['en', 'de', 'fr', 'he', 'hi', 'id', 'pl', 'pt-BR', 'tr', 'vi', 'zh-CN'] as const

export type LocaleCode = (typeof SUPPORTED_LOCALES)[number]

export const DEFAULT_LOCALE: LocaleCode = 'en'

export const LOCALE_NATIVE_NAMES: Record<LocaleCode, string> = {
  en: 'English',
  de: 'Deutsch',
  fr: 'Français',
  he: 'עברית',
  hi: 'हिन्दी',
  id: 'Bahasa Indonesia',
  pl: 'Polski',
  'pt-BR': 'Português (Brasil)',
  tr: 'Türkçe',
  vi: 'Tiếng Việt',
  'zh-CN': '简体中文',
}

const RTL_LOCALES: ReadonlySet<LocaleCode> = new Set(['he'])

export function isRtlLocale(locale: LocaleCode): boolean {
  return RTL_LOCALES.has(locale)
}

export function isSupportedLocale(value: string): value is LocaleCode {
  return (SUPPORTED_LOCALES as readonly string[]).includes(value)
}

const LEGACY_LANGUAGE_CODES: Record<string, string> = {
  iw: 'he',
  in: 'id',
}

const BASE_LANGUAGE_FALLBACKS: Record<string, LocaleCode> = {
  pt: 'pt-BR',
  zh: 'zh-CN',
}

export function matchLocale(raw: string): LocaleCode {
  const normalized = raw.trim().replace(/_/g, '-')
  if (!normalized) return DEFAULT_LOCALE

  const lower = normalized.toLowerCase()
  for (const locale of SUPPORTED_LOCALES) {
    if (locale.toLowerCase() === lower) return locale
  }

  const rawBase = lower.split('-')[0] ?? lower
  const base = LEGACY_LANGUAGE_CODES[rawBase] ?? rawBase
  if (isSupportedLocale(base)) return base
  const mapped = BASE_LANGUAGE_FALLBACKS[base]
  if (mapped) return mapped
  return DEFAULT_LOCALE
}

export type LocaleStrings = typeof en

type AnyRecord = Record<string, unknown>

function isRecord(value: unknown): value is AnyRecord {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function deepMerge(base: AnyRecord, override: AnyRecord): AnyRecord {
  const result: AnyRecord = { ...base }
  for (const key of Object.keys(override)) {
    const overrideValue = override[key]
    if (overrideValue == null) continue
    if (isRecord(base[key]) && isRecord(overrideValue)) {
      result[key] = deepMerge(base[key], overrideValue)
    } else {
      result[key] = overrideValue
    }
  }
  return result
}

const OVERRIDES: Record<LocaleCode, AnyRecord> = {
  en,
  de,
  fr,
  he,
  hi,
  id,
  pl,
  'pt-BR': ptBR,
  tr,
  vi,
  'zh-CN': zhCN,
}

const mergedCache = new Map<LocaleCode, LocaleStrings>()

export function getStrings(locale: LocaleCode): LocaleStrings {
  const cached = mergedCache.get(locale)
  if (cached) return cached
  const merged = (locale === 'en' ? en : deepMerge(en, OVERRIDES[locale])) as LocaleStrings
  mergedCache.set(locale, merged)
  return merged
}

export type TranslateVars = Record<string, string | number>

export type Translator = (key: string, vars?: TranslateVars) => string

function lookup(strings: AnyRecord, key: string): string | undefined {
  let current: unknown = strings
  for (const part of key.split('.')) {
    if (!isRecord(current)) return undefined
    current = current[part]
  }
  return typeof current === 'string' ? current : undefined
}

export function createTranslator(locale: LocaleCode): Translator {
  const strings = getStrings(locale) as AnyRecord
  return (key, vars) => {
    const template = lookup(strings, key)
    if (template === undefined) return key
    if (!vars) return template
    return template.replace(/\{(\w+)\}/g, (match, name: string) =>
      name in vars ? String(vars[name]) : match,
    )
  }
}
