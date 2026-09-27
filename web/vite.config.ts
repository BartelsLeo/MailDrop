import { resolve } from "node:path";
import { defineConfig } from "vite";

// HTTPS ist Pflicht: Outlook lädt Add-ins nur über HTTPS mit vertrauenswürdigem Zertifikat.
// Lokal liefert "office-addin-dev-certs" ein Entwicklerzertifikat (einmalig
// "npx office-addin-dev-certs install" auf dem Test-PC). Im Build (Azure) wird es nicht gebraucht.
export default defineConfig(async ({ command }) => {
  let https: { key: Buffer; cert: Buffer; ca: Buffer } | undefined;
  if (command === "serve") {
    const devCerts = await import("office-addin-dev-certs");
    https = await devCerts.getHttpsServerOptions();
  }
  return {
    root: resolve(import.meta.dirname, "src"),
    publicDir: resolve(import.meta.dirname, "public"),
    envDir: import.meta.dirname,
    server: { port: 3000, strictPort: true, https },
    build: {
      outDir: resolve(import.meta.dirname, "dist"),
      emptyOutDir: true,
      rollupOptions: { input: { taskpane: resolve(import.meta.dirname, "src/taskpane.html") } },
    },
  };
});
