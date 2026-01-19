import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';

export default defineConfig({
  plugins: [react()],

  // Critical for Electron file:// protocol
  base: './',

  // Build configuration
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    sourcemap: true,
    rollupOptions: {
      output: {
        manualChunks: undefined
      }
    }
  },

  // Dev server
  server: {
    port: 5173,
    strictPort: true, // Fail if port unavailable
    cors: true
  },

  // Path aliases
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
});
