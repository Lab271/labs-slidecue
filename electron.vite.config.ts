// SPDX-License-Identifier: Apache-2.0
// SPDX-FileCopyrightText: 2025-2026 Schuberg Philis / Lab271
import { defineConfig, externalizeDepsPlugin } from 'electron-vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'path';
import { viteStaticCopy } from 'vite-plugin-static-copy';

export default defineConfig({
  main: {
    plugins: [
      externalizeDepsPlugin(),
      viteStaticCopy({
        // electron-vite builds the main process in Vite's `ssr` environment, but
        // vite-plugin-static-copy v4 gained an `environment` option that defaults
        // to `client` and silently no-ops its writeBundle hook everywhere else.
        // Without this the target never fired and out/main/powerpoint-bridge.ps1
        // was never emitted, leaving the __dirname lookup in
        // src/main/pptx/windows-powershell.ts resolving against a file that did
        // not exist.
        environment: 'ssr',
        targets: [
          {
            src: 'src/main/pptx/powerpoint-bridge.ps1',
            // `dest` alone reproduces the matched path's directories underneath
            // it, which would land the script in out/main/src/main/pptx/. The
            // __dirname lookup expects it flat next to index.js, so strip the
            // leading segments.
            dest: '.',
            rename: { stripBase: true }
          }
        ]
      })
    ],
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/main/index.ts'),
        },
      },
    },
  },
  preload: {
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/preload/index.ts'),
        },
      },
    },
  },
  renderer: {
    root: 'src/renderer',
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/renderer/index.html'),
        },
      },
    },
    plugins: [react()],
  },
});
