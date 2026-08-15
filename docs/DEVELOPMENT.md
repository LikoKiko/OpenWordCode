# Development

```powershell
npm install
npm run dev
```

The Core is TypeScript executed with `tsx`. The task pane is a Vite React app with a local HTTPS certificate. Vite proxies `/api`, `/health`, and `/healthz` to the Core during development.

The codebase intentionally uses relative source imports between package boundaries in the MVP so the project can be run without a package publishing step. The directories remain separate and enforce the architectural ownership: provider code cannot import the Word adapter, and the UI calls adapters through their interface.

For an isolated Core data directory:

```powershell
$env:OPENWORDCODE_DATA_DIR = "$PWD\.local-openwordcode"
npm run start
```
