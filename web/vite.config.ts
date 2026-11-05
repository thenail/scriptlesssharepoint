import react from "@vitejs/plugin-react";
import { defineConfig, type Plugin } from "vitest/config";

/** The page's content security policy, in builds only (the dev server injects inline scripts). */
const csp: Plugin = {
  name: "csp",
  apply: "build",
  transformIndexHtml: (html) =>
    html.replace(
      "<meta charset=\"UTF-8\" />",
      `<meta charset="UTF-8" />
    <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; form-action 'none'" />`,
    ),
};

// The UI is served from the package folder (WebView2's virtual host), so
// paths are relative. `yarn dev` serves it on a fixed port for SSP_WEB_URL.
export default defineConfig({
  base: "./",
  plugins: [react(), csp],
  build: {
    outDir: "dist",
    emptyOutDir: true,
    sourcemap: false,
    chunkSizeWarningLimit: 1500,
  },
  server: { port: 5279, strictPort: true },
  test: { environment: "node" },
});
