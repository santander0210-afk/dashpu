import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

const projectRoot = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  base: "./",
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      "@": path.resolve(projectRoot, "client/src"),
    },
  },
  root: path.resolve(projectRoot, "client"),
  publicDir: path.resolve(projectRoot, "client/public"),
  build: {
    outDir: path.resolve(projectRoot, "dist"),
    emptyOutDir: true,
  },
  server: {
    host: true,
    port: 5173,
    allowedHosts: ["localhost", "127.0.0.1", ".manus.computer"],
  },
});
