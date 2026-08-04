#!/usr/bin/env node
/**
 * i18n consistency check for the desktop chrome — zero-dep.
 *
 * TypeScript can't catch these because `t()` takes a plain string and returns
 * one: a typo or a key that only exists in one locale compiles fine and then
 * renders the raw key (`status.ready`) to the user at runtime.
 *
 * Two checks:
 *   1. every `t('some.key')` in src/ resolves in EVERY locale file;
 *   2. all locale files share en.json's exact key set (nothing missing, nothing extra).
 *
 * Dynamic keys (`t(status.key)`, `t(KIND_LABEL_KEY[kind])`) are skipped — the
 * value is only known at runtime — so keep those namespaces covered by check 2.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const SRC = join(ROOT, 'src')
const LOCALES_DIR = join(SRC, 'i18n')
const BASE = 'en'
const LOCALES = ['en', 'de', 'fr', 'he', 'hi', 'id', 'pl', 'pt-BR', 'tr', 'vi', 'zh-CN']

const T_CALL = /(?<![A-Za-z0-9_.])t\('([a-zA-Z0-9_.]+)'/g

function walk(dir) {
  const out = []
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry)
    if (statSync(path).isDirectory()) out.push(...walk(path))
    else if (/\.tsx?$/.test(path)) out.push(path)
  }
  return out
}

function flatten(value, prefix = '') {
  const keys = []
  for (const [key, child] of Object.entries(value)) {
    const full = prefix ? `${prefix}.${key}` : key
    if (child && typeof child === 'object') keys.push(...flatten(child, full))
    else keys.push(full)
  }
  return keys
}

function resolve(dict, key) {
  let current = dict
  for (const part of key.split('.')) {
    if (!current || typeof current !== 'object' || !(part in current)) return undefined
    current = current[part]
  }
  return typeof current === 'string' ? current : undefined
}

const dicts = Object.fromEntries(
  LOCALES.map((locale) => [locale, JSON.parse(readFileSync(join(LOCALES_DIR, `${locale}.json`), 'utf8'))]),
)

const problems = []

for (const file of walk(SRC)) {
  const source = readFileSync(file, 'utf8')
  for (const [, key] of source.matchAll(T_CALL)) {
    for (const locale of LOCALES) {
      if (resolve(dicts[locale], key) === undefined) {
        problems.push(`${relative(ROOT, file)}: t('${key}') missing from ${locale}.json`)
      }
    }
  }
}

const baseKeys = new Set(flatten(dicts[BASE]))
for (const locale of LOCALES) {
  if (locale === BASE) continue
  const keys = new Set(flatten(dicts[locale]))
  for (const key of baseKeys) {
    if (!keys.has(key)) problems.push(`${locale}.json: missing "${key}" (present in ${BASE}.json)`)
  }
  for (const key of keys) {
    if (!baseKeys.has(key)) problems.push(`${locale}.json: extra "${key}" (absent from ${BASE}.json)`)
  }
}

if (problems.length > 0) {
  console.error(`i18n check failed — ${problems.length} problem(s):`)
  for (const problem of problems) console.error(`  ${problem}`)
  process.exit(1)
}

console.log(`i18n check passed — ${LOCALES.length} locales, ${baseKeys.size} keys`)
