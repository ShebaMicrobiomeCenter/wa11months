import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig} from 'vite';

export default defineConfig(() => {
  return {
    base: '/wa11months/',
    plugins: [
      react(),
      tailwindcss(),
      {
        name: 'redirect-base',
        configureServer(server) {
          server.middlewares.use((req, res, next) => {
            const base = '/wa11months';
            const url = new URL(req.url || '', `http://${req.headers.host}`);
            if (url.pathname === base) {
              res.writeHead(301, {Location: base + '/' + url.search});
              res.end();
            } else {
              next();
            }
          });
        },
      },
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
