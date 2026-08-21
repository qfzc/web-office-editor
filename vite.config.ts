import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

const streamBrowserify = fileURLToPath(
  new URL('./node_modules/stream-browserify/index.js', import.meta.url),
);
const eventsBrowserify = fileURLToPath(
  new URL('./node_modules/events/events.js', import.meta.url),
);

export default defineConfig({
  resolve: {
    alias: {
      events: eventsBrowserify,
      stream: streamBrowserify,
    },
  },
  build: {
    lib: {
      entry: 'src/index.ts',
      formats: ['es'],
      fileName: 'index',
    },
    rollupOptions: {
      output: {
        chunkFileNames: 'chunks/[name]-[hash].js',
        entryFileNames: '[name].js',
      },
    },
    chunkSizeWarningLimit: 3000,
    sourcemap: false,
    target: 'es2022',
  },
  test: {
    environment: 'jsdom',
    coverage: {
      reporter: ['text', 'html'],
    },
  },
});
