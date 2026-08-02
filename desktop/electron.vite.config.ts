import { resolve } from 'node:path'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

const rootDir = import.meta.dirname

export default defineConfig({
  main: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve(rootDir, 'src/main/index.ts'),
      },
    },
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: resolve(rootDir, 'src/preload/index.ts'),
        output: { format: 'cjs', entryFileNames: 'index.cjs' },
      },
    },
  },
  renderer: {
    root: resolve(rootDir, 'src/renderer'),
    publicDir: resolve(rootDir, 'src/renderer/public'),
    plugins: [react(), tailwindcss()],
    resolve: {
      alias: {
        '@shared': resolve(rootDir, 'src/shared'),
        '@renderer': resolve(rootDir, 'src/renderer'),
      },
    },
    build: {
      target: 'esnext',
      rollupOptions: {
        input: resolve(rootDir, 'src/renderer/index.html'),
      },
    },
  },
})
