import { basename } from 'node:path'
import { app, Menu, shell, type MenuItemConstructorOptions } from 'electron'
import type { MenuAction } from '../shared/ipc.js'

type Dispatch = (action: MenuAction) => void

export interface MenuContext {
  dispatch: Dispatch
  recents: readonly string[]
  onOpenRecent: (path: string) => void
  onClearRecents: () => void
  onCheckForUpdates: () => void
  onShowAbout: () => void
}

const isMac = process.platform === 'darwin'

function openRecentSubmenu(context: MenuContext): MenuItemConstructorOptions[] {
  const entries: MenuItemConstructorOptions[] = context.recents.length
    ? context.recents.map((path) => ({
        label: basename(path),
        toolTip: path,
        click: () => context.onOpenRecent(path),
      }))
    : [{ label: 'No Recent Files', enabled: false }]

  return [
    ...entries,
    { type: 'separator' },
    { label: 'Clear Recent', enabled: context.recents.length > 0, click: context.onClearRecents },
  ]
}

export function buildMenu(context: MenuContext): Menu {
  const send = (action: MenuAction) => () => context.dispatch(action)

  const appMenu: MenuItemConstructorOptions[] = isMac
    ? [
        {
          label: app.name,
          submenu: [
            { role: 'about' },
            { label: 'Check for Updates…', click: context.onCheckForUpdates },
            { type: 'separator' },
            { role: 'services' },
            { type: 'separator' },
            { role: 'hide' },
            { role: 'hideOthers' },
            { role: 'unhide' },
            { type: 'separator' },
            { role: 'quit' },
          ],
        },
      ]
    : []

  const template: MenuItemConstructorOptions[] = [
    ...appMenu,
    {
      label: '&File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: send('file:new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: send('file:open') },
        { label: 'Open Recent', submenu: openRecentSubmenu(context) },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: send('file:save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: send('file:saveAs') },
        { type: 'separator' },
        { label: 'Export as PDF…', accelerator: 'CmdOrCtrl+Shift+E', click: send('file:exportPdf') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    {
      label: '&Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: send('edit:undo') },
        {
          label: 'Redo',
          accelerator: process.platform === 'win32' ? 'Ctrl+Y' : 'Shift+CmdOrCtrl+Z',
          click: send('edit:redo'),
        },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'CmdOrCtrl+X', click: send('edit:cut') },
        { label: 'Copy', accelerator: 'CmdOrCtrl+C', click: send('edit:copy') },
        { label: 'Paste', accelerator: 'CmdOrCtrl+V', click: send('edit:paste') },
        ...(isMac
          ? ([{ role: 'pasteAndMatchStyle' }, { role: 'delete' }, { role: 'selectAll' }] as MenuItemConstructorOptions[])
          : ([{ role: 'delete' }, { type: 'separator' }, { role: 'selectAll' }] as MenuItemConstructorOptions[])),
        { type: 'separator' },
        { label: 'Find…', accelerator: 'CmdOrCtrl+F', click: send('edit:find') },
      ],
    },
    {
      label: '&View',
      submenu: [
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+Plus', click: send('view:zoomIn') },
        {
          label: 'Zoom In',
          accelerator: 'CmdOrCtrl+=',
          click: send('view:zoomIn'),
          visible: false,
          acceleratorWorksWhenHidden: true,
        },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', click: send('view:zoomOut') },
        { label: 'Actual Size', accelerator: 'CmdOrCtrl+0', click: send('view:zoomReset') },
        { type: 'separator' },
        { label: 'Word Count', accelerator: 'CmdOrCtrl+Shift+G', click: send('view:wordCount') },
        { label: 'Spell Check', accelerator: 'CmdOrCtrl+Shift+;', click: send('view:spellCheck') },
        { type: 'separator' },
        { label: 'Freeze Top Row', click: send('view:freezeTopRow') },
        { label: 'Freeze First Column', click: send('view:freezeFirstColumn') },
        { label: 'Unfreeze Panes', click: send('view:unfreeze') },
        { type: 'separator' },
        { role: 'togglefullscreen' },
        ...(app.isPackaged
          ? []
          : ([{ role: 'reload' }, { role: 'forceReload' }, { role: 'toggleDevTools' }] as MenuItemConstructorOptions[])),
      ],
    },
    {
      label: '&Window',
      submenu: isMac
        ? [
            { role: 'minimize' },
            { role: 'zoom' },
            { type: 'separator' },
            { role: 'front' },
            { type: 'separator' },
            { role: 'window' },
          ]
        : [{ role: 'minimize' }, { role: 'zoom' }, { role: 'close' }],
    },
    {
      role: 'help',
      submenu: [
        {
          label: 'NexOffice Documentation',
          click: () => {
            void shell.openExternal('https://betteroffice.dev')
          },
        },
        ...(isMac
          ? []
          : ([
              { type: 'separator' },
              { label: 'Check for Updates…', click: context.onCheckForUpdates },
              { label: 'About NexOffice', click: context.onShowAbout },
            ] as MenuItemConstructorOptions[])),
      ],
    },
  ]

  const menu = Menu.buildFromTemplate(template)
  Menu.setApplicationMenu(menu)
  return menu
}
