import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import basicSsl from "@vitejs/plugin-basic-ssl";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const officeCertificateDirectory = path.join(os.homedir(), ".office-addin-dev-certs");
const officeCertificatePath = path.join(officeCertificateDirectory, "localhost.crt");
const officeKeyPath = path.join(officeCertificateDirectory, "localhost.key");
const hasTrustedOfficeCertificate =
  fs.existsSync(officeCertificatePath) && fs.existsSync(officeKeyPath);

export default defineConfig({
  root: fileURLToPath(new URL(".", import.meta.url)),
  plugins: [react(), ...(hasTrustedOfficeCertificate ? [] : [basicSsl()])],
  server: {
    host: "localhost",
    port: 3000,
    strictPort: true,
    https: hasTrustedOfficeCertificate
      ? {
          cert: fs.readFileSync(officeCertificatePath),
          key: fs.readFileSync(officeKeyPath),
        }
      : undefined,
    proxy: {
      "/api": "http://127.0.0.1:10200",
      "/health": "http://127.0.0.1:10200",
      "/healthz": "http://127.0.0.1:10200",
    },
  },
  build: { outDir: "dist", emptyOutDir: true },
});
