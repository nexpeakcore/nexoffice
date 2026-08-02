import { app, Menu, shell, type MenuItemConstructorOptions } from 'electron'
import type { MenuAction } from '../shared/ipc.js'

type Dispatch = (action: MenuAction) => void

const isMac = process.platform === 'darwin'

export function buildMenu(dispatch: Dispatch): Menu {
  const send = (action: MenuAction) => () => dispatch(action)

  const appMenu: MenuItemConstructorOptions[] = isMac
    ? [
        {
          label: app.name,
          submenu: [
            { role: 'about' },
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
        ...(isMac
          ? ([
              { role: 'recentDocuments', submenu: [{ role: 'clearRecentDocuments' }] },
            ] as MenuItemConstructorOptions[])
          : []),
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
        { role: 'undo' },
        { role: 'redo' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
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
      ],
    },
  ]

  const menu = Menu.buildFromTemplate(template)
  Menu.setApplicationMenu(menu)
  return menu
}
