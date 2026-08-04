import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react'
import {
  createTranslator,
  DEFAULT_LOCALE,
  isRtlLocale,
  type LocaleCode,
  type Translator,
} from '../i18n/index.js'

interface LocaleContextValue {
  locale: LocaleCode
  t: Translator
}

const LocaleContext = createContext<LocaleContextValue>({
  locale: DEFAULT_LOCALE,
  t: createTranslator(DEFAULT_LOCALE),
})

export function LocaleProvider({ children }: { children: ReactNode }) {
  const [locale, setLocale] = useState<LocaleCode>(DEFAULT_LOCALE)

  useEffect(() => {
    let mounted = true
    void window.nexoffice.getLocale().then((initial) => {
      if (mounted) setLocale(initial)
    })
    const unsubscribe = window.nexoffice.onLocaleChanged(setLocale)
    return () => {
      mounted = false
      unsubscribe()
    }
  }, [])

  useEffect(() => {
    document.documentElement.lang = locale
    document.documentElement.dir = isRtlLocale(locale) ? 'rtl' : 'ltr'
  }, [locale])

  const value = useMemo<LocaleContextValue>(
    () => ({ locale, t: createTranslator(locale) }),
    [locale],
  )

  return <LocaleContext.Provider value={value}>{children}</LocaleContext.Provider>
}

export function useI18n(): LocaleContextValue {
  return useContext(LocaleContext)
}
