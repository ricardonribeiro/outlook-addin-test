import { defineConfig } from 'vite';
import { readFileSync } from 'fs';
import { homedir } from 'os';
import { join, resolve } from 'path';

function httpsConfig() {
  const certDir = join(homedir(), '.office-addin-dev-certs');
  try {
    return {
      key: readFileSync(join(certDir, 'localhost.key')),
      cert: readFileSync(join(certDir, 'localhost.crt')),
    };
  } catch {
    // Certs not installed yet — fall back to Vite self-signed cert.
    // Run `npx office-addin-dev-certs install` to get certs trusted by Outlook.
    console.warn('[vite] Dev certs not found; falling back to self-signed. Run: npx office-addin-dev-certs install');
    return true;
  }
}

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        commands: resolve(__dirname, 'commands.html'),
      },
    },
    outDir: 'dist',
    emptyOutDir: true,
  },
  server: {
    port: 3000,
    https: httpsConfig(),
    headers: {
      // Required for Office.js to load the add-in in Outlook on the web
      'Access-Control-Allow-Origin': '*',
    },
  },
  preview: {
    port: 3000,
    https: httpsConfig(),
  },
});
