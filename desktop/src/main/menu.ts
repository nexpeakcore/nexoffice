import { basename } from 'node:path'
import { app, Menu, shell, type MenuItemConstructorOptions } from 'electron'
import {
  LOCALE_NATIVE_NAMES,
  SUPPORTED_LOCALES,
  type LocaleCode,
  type Translator,
} from '../i18n/index.js'
import type { DocumentKind, EditCapabilities, MenuAction } from '../shared/ipc.js'

type Dispatch = (action: MenuAction) => void

export interface MenuContext {
  t: Translator
  locale: LocaleCode
  dispatch: Dispatch
  documentKind: DocumentKind | null
  editCapabilities: EditCapabilities
  recents: readonly string[]
  onOpenRecent: (path: string) => void
  onClearRecents: () => void
  onSelectLocale: (locale: LocaleCode) => void
  onCheckForUpdates: () => void
  onShowAbout: () => void
  onShowProcesses: () => void
}

const isMac = process.platform === 'darwin'

function openRecentSubmenu(context: MenuContext): MenuItemConstructorOptions[] {
  const { t } = context
  const entries: MenuItemConstructorOptions[] = context.recents.length
    ? context.recents.map((path) => ({
        label: basename(path),
        toolTip: path,
        click: () => context.onOpenRecent(path),
      }))
    : [{ label: t('menu.file.noRecentFiles'), enabled: false }]

  return [
    ...entries,
    { type: 'separator' },
    {
      label: t('menu.file.clearRecent'),
      enabled: context.recents.length > 0,
      click: context.onClearRecents,
    },
  ]
}

function languageSubmenu(context: MenuContext): MenuItemConstructorOptions[] {
  return SUPPORTED_LOCALES.map((locale) => ({
    label: LOCALE_NATIVE_NAMES[locale],
    type: 'radio',
    checked: locale === context.locale,
    click: () => context.onSelectLocale(locale),
  }))
}

export function buildMenu(context: MenuContext): Menu {
  const { t } = context
  const send = (action: MenuAction) => () => context.dispatch(action)
  const hasZoom = context.documentKind !== null
  const hasFind = context.documentKind === 'docx'
  const hasFreeze = context.documentKind === 'xlsx'
  const hasProofing = context.documentKind === 'docx' || context.documentKind === 'pptx'
  const canExportPdf = context.documentKind !== null
  // What the focused editor says its clipboard verbs can act on right now: the
  // pptx editor has no shape clipboard, so cut, copy, paste and delete would
  // silently do nothing over a shape selection.
  const edits = context.editCapabilities

  const appMenu: MenuItemConstructorOptions[] = isMac
    ? [
        {
          label: app.name,
          submenu: [
            { role: 'about', label: t('menu.help.about') },
            { label: t('menu.app.checkForUpdates'), click: context.onCheckForUpdates },
            { type: 'separator' },
            { role: 'services', label: t('menu.app.services') },
            { type: 'separator' },
            { role: 'hide', label: t('menu.app.hide') },
            { role: 'hideOthers', label: t('menu.app.hideOthers') },
            { role: 'unhide', label: t('menu.app.unhide') },
            { type: 'separator' },
            { role: 'quit', label: t('menu.app.quit') },
          ],
        },
      ]
    : []

  const template: MenuItemConstructorOptions[] = [
    ...appMenu,
    {
      label: t('menu.file.label'),
      submenu: [
        { label: t('menu.file.new'), accelerator: 'CmdOrCtrl+N', click: send('file:new') },
        { label: t('menu.file.open'), accelerator: 'CmdOrCtrl+O', click: send('file:open') },
        { label: t('menu.file.openRecent'), submenu: openRecentSubmenu(context) },
        { type: 'separator' },
        { label: t('menu.file.save'), accelerator: 'CmdOrCtrl+S', click: send('file:save') },
        {
          label: t('menu.file.saveAs'),
          accelerator: 'CmdOrCtrl+Shift+S',
          click: send('file:saveAs'),
        },
        { type: 'separator' },
        {
          label: t('menu.file.exportPdf'),
          accelerator: 'CmdOrCtrl+Shift+E',
          enabled: canExportPdf,
          click: send('file:exportPdf'),
        },
        {
          label: t('menu.file.print'),
          accelerator: 'CmdOrCtrl+P',
          enabled: canExportPdf,
          click: send('file:print'),
        },
        { type: 'separator' },
        isMac
          ? { role: 'close', label: t('menu.file.closeWindow') }
          : { role: 'quit', label: t('menu.file.exit') },
      ],
    },
    {
      label: t('menu.edit.label'),
      submenu: [
        { label: t('menu.edit.undo'), accelerator: 'CmdOrCtrl+Z', click: send('edit:undo') },
        {
          label: t('menu.edit.redo'),
          accelerator: process.platform === 'win32' ? 'Ctrl+Y' : 'Shift+CmdOrCtrl+Z',
          click: send('edit:redo'),
        },
        { type: 'separator' },
        {
          label: t('menu.edit.cut'),
          accelerator: 'CmdOrCtrl+X',
          enabled: edits.cut,
          click: send('edit:cut'),
        },
        {
          label: t('menu.edit.copy'),
          accelerator: 'CmdOrCtrl+C',
          enabled: edits.copy,
          click: send('edit:copy'),
        },
        {
          label: t('menu.edit.paste'),
          accelerator: 'CmdOrCtrl+V',
          enabled: edits.paste,
          click: send('edit:paste'),
        },
        ...(isMac
          ? ([
              {
                role: 'pasteAndMatchStyle',
                label: t('menu.edit.pasteAndMatchStyle'),
                enabled: edits.paste,
              },
              { label: t('menu.edit.delete'), enabled: edits.delete, click: send('edit:delete') },
              {
                label: t('menu.edit.selectAll'),
                accelerator: 'CmdOrCtrl+A',
                enabled: edits.selectAll,
                click: send('edit:selectAll'),
              },
            ] as MenuItemConstructorOptions[])
          : ([
              { label: t('menu.edit.delete'), enabled: edits.delete, click: send('edit:delete') },
              { type: 'separator' },
              {
                label: t('menu.edit.selectAll'),
                accelerator: 'CmdOrCtrl+A',
                enabled: edits.selectAll,
                click: send('edit:selectAll'),
              },
            ] as MenuItemConstructorOptions[])),
        { type: 'separator' },
        {
          label: t('menu.edit.find'),
          accelerator: 'CmdOrCtrl+F',
          enabled: hasFind,
          click: send('edit:find'),
        },
      ],
    },
    {
      label: t('menu.view.label'),
      submenu: [
        {
          label: t('menu.view.zoomIn'),
          accelerator: 'CmdOrCtrl+Plus',
          enabled: hasZoom,
          click: send('view:zoomIn'),
        },
        {
          label: t('menu.view.zoomIn'),
          accelerator: 'CmdOrCtrl+=',
          enabled: hasZoom,
          click: send('view:zoomIn'),
          visible: false,
          acceleratorWorksWhenHidden: true,
        },
        {
          label: t('menu.view.zoomOut'),
          accelerator: 'CmdOrCtrl+-',
          enabled: hasZoom,
          click: send('view:zoomOut'),
        },
        {
          label: t('menu.view.actualSize'),
          accelerator: 'CmdOrCtrl+0',
          enabled: hasZoom,
          click: send('view:zoomReset'),
        },
        { type: 'separator' },
        {
          label: t('menu.view.wordCount'),
          accelerator: 'CmdOrCtrl+Shift+G',
          enabled: hasProofing,
          click: send('view:wordCount'),
        },
        {
          label: t('menu.view.spellCheck'),
          accelerator: 'CmdOrCtrl+Shift+;',
          enabled: hasProofing,
          click: send('view:spellCheck'),
        },
        {
          label: t('menu.view.aiAssistant'),
          accelerator: 'CmdOrCtrl+Shift+A',
          enabled: context.documentKind === 'xlsx',
          click: send('view:aiAssistant'),
        },
        { type: 'separator' },
        { label: t('menu.view.freezeTopRow'), enabled: hasFreeze, click: send('view:freezeTopRow') },
        {
          label: t('menu.view.freezeFirstColumn'),
          enabled: hasFreeze,
          click: send('view:freezeFirstColumn'),
        },
        { label: t('menu.view.unfreeze'), enabled: hasFreeze, click: send('view:unfreeze') },
        { type: 'separator' },
        { label: t('menu.view.language'), submenu: languageSubmenu(context) },
        { type: 'separator' },
        { role: 'togglefullscreen', label: t('menu.view.fullScreen') },
        ...(app.isPackaged
          ? []
          : ([{ role: 'reload' }, { role: 'forceReload' }, { role: 'toggleDevTools' }] as MenuItemConstructorOptions[])),
      ],
    },
    {
      label: t('menu.window.label'),
      submenu: isMac
        ? [
            { role: 'minimize', label: t('menu.window.minimize') },
            { role: 'zoom', label: t('menu.window.zoom') },
            { type: 'separator' },
            { role: 'front', label: t('menu.window.front') },
            { type: 'separator' },
            { role: 'window' },
          ]
        : [
            { role: 'minimize', label: t('menu.window.minimize') },
            { role: 'zoom', label: t('menu.window.zoom') },
            { role: 'close', label: t('menu.file.closeWindow') },
          ],
    },
    {
      role: 'help',
      label: t('menu.help.label'),
      submenu: [
        {
          label: t('menu.help.website'),
          click: () => {
            void shell.openExternal('https://nexpeak.net')
          },
        },
        { type: 'separator' },
        { label: t('menu.help.processes'), click: context.onShowProcesses },
        ...(isMac
          ? []
          : ([
              { type: 'separator' },
              { label: t('menu.help.checkForUpdates'), click: context.onCheckForUpdates },
              { label: t('menu.help.about'), click: context.onShowAbout },
            ] as MenuItemConstructorOptions[])),
      ],
    },
  ]

  const menu = Menu.buildFromTemplate(template)
  Menu.setApplicationMenu(menu)
  return menu
}
