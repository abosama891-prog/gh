import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [
    react(),
    {
      name: 'version-service-worker-per-build',
      apply: 'build',
      async closeBundle() {
        const serviceWorkerPath = resolve('dist/sw.js');
        const serviceWorker = await readFile(serviceWorkerPath, 'utf8');
        if (!serviceWorker.includes('__BUILD_VERSION__')) {
          throw new Error('Service worker build version placeholder is missing');
        }
        await writeFile(serviceWorkerPath, serviceWorker.replace('__BUILD_VERSION__', Date.now().toString(36)));
      },
    },
  ],
  server: {
    port: 3000,
    open: true,
    watch: {
      ignored: ['**/dist/**'],
    },
  },
  build: {
    outDir: 'dist',
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.jsx', '.js'],
    alias: {
      '@': '/src',
    },
  },
});