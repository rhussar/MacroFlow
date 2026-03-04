import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import path from 'node:path';

export default defineConfig({
  plugins: [
    react(),
    tailwindcss()
  ],

  // Critical for Electron file:// protocol
  base: './',

  // Build configuration
  build: {
    outDir: 'dist',
    assetsDir: 'assets',
    sourcemap: true,
    assetsInlineLimit: 10240, // inline PNGs under 10KB as base64
    rollupOptions: {
      output: {
        manualChunks: undefined
      }
    }
  },

  // Dev server
  server: {
    port: 5173,
    strictPort: true,
    cors: true
  },

  // Path aliases
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src')
    }
  }
});
