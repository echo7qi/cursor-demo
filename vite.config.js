import { defineConfig } from 'vite';

/** 静态 HTML 看板：本地预览用，避免直接 file:// 限制 */
export default defineConfig({
  root: '.',
  server: {
    port: 5174,
    strictPort: false,
    open: '/index.html',
  },
  publicDir: false,
});
