/// <reference types="vite/client" />

import type { NexOfficeApi } from '../preload/index.js'

declare global {
  interface Window {
    nexoffice: NexOfficeApi
  }
}

export {}
