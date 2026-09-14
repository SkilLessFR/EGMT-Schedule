import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',
    watch: {
      ignored: [
        '**/*.xlsx',
        '**/*.xls',
        '**/*.xlsm',
        '**/*.csv',
        '**/*.png',
        '**/*.jpg',
        '**/*.jpeg',
        '**/*.gif',
        '**/*.webp',
        '**/*.pdf',
        '**/*.tmp',
        '**/*.crdownload',
        '**/*.part',
        '**/node_bin/**',
        '**/.git/**',
        '**/dist/**',
      ],
    },
  },
});
