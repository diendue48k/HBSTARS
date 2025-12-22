
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  base: '/HBSTARS/', // Phải trùng với tên repository trên GitHub
  build: {
    outDir: 'dist',
    sourcemap: false
  }
});