import { defineConfig } from "vitest/config";

// Eigene Test-Konfiguration: Die Vite-Konfiguration (root = src, HTTPS-Entwicklerzertifikat)
// ist für den Add-in-Server gedacht und wird für die Tests nicht gebraucht.
export default defineConfig({
  test: {
    root: import.meta.dirname,
    include: ["test/**/*.test.ts"],
    testTimeout: 120_000,
  },
});
