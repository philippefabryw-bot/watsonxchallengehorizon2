import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      // Proxy /wx-api/* → watsonx AU-SYD, stripping the /wx-api prefix.
      // This runs server-side so no CORS headers are needed.
      '/wx-api': {
        target: 'https://au-syd.ml.cloud.ibm.com',
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/wx-api/, ''),
        secure: true,
      },
    },
  },
});
