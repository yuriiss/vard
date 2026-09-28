import { defineConfig } from 'vitest/config';
import react from '@vitejs/plugin-react';

// `base: './'` keeps asset paths relative so the build works from any sub-path
// (GitHub Pages, SharePoint, a file share, ...).
export default defineConfig({
  base: './',
  plugins: [react()],
  // exceljs is lazy-loaded only when exporting.
  build: { chunkSizeWarningLimit: 1000 },
  test: {
    include: ['src/**/*.test.ts'],
  },
});
