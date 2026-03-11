import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';

/**
 * Chrome extension pages don't support the `crossorigin` attribute on scripts.
 * When present, the script executes in a reduced-privilege context and loses
 * access to chrome.* APIs (e.g. chrome.storage becomes undefined).
 */
function stripCrossorigin(): Plugin {
  return {
    name: 'strip-crossorigin',
    enforce: 'post',
    transformIndexHtml(html) {
      return html.replace(/ crossorigin/g, '');
    },
  };
}

export default defineConfig({
  plugins: [react(), stripCrossorigin()],
  // Relative base so extension page paths resolve correctly and don't
  // trigger CORS-style loading that strips chrome.* API access.
  base: './',
  publicDir: 'public',
  build: {
    outDir: 'dist',
    rollupOptions: {
      // Use `popup.html` as a separate HTML entry so the extension's popup is emitted
      input: {
        popup: 'popup.html',
        offscreen: 'offscreen.html',
        background: 'src/background.ts',
      },
      output: {
        entryFileNames: '[name].js',
        chunkFileNames: '[name].js',
        assetFileNames: '[name].[ext]'
      }
    }
  }
});
