import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "node:path";
import { fileURLToPath } from "node:url";
import fs from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const keyPath = path.resolve(__dirname, "certs/localhost-key.pem");
const certPath = path.resolve(__dirname, "certs/localhost.pem");
const https =
  fs.existsSync(keyPath) && fs.existsSync(certPath)
    ? {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath),
      }
    : true;

export default defineConfig({
  plugins: [react()],
  root: path.resolve(__dirname, "src"),
  publicDir: path.resolve(__dirname, "public"),
  server: {
    port: 3000,
    https,
    strictPort: true,
    // Office 插件依赖 URL hash 里的 _host_Info；HMR 整页刷新会弄丢，导致 Word API 消失
    hmr: false,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    proxy: {
      // 浏览器/Word WebView 直连 DeepSeek 会 CORS；开发态走本地代理
      "/llm": {
        target: "https://api.deepseek.com",
        changeOrigin: true,
        secure: true,
        rewrite: (p) => p.replace(/^\/llm/, ""),
      },
    },
  },
  build: {
    outDir: path.resolve(__dirname, "dist"),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, "src/taskpane/index.html"),
        commands: path.resolve(__dirname, "src/commands/index.html"),
      },
    },
  },
});
