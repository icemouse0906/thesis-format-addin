import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const src = path.join(root, "node_modules/@microsoft/office-js/dist");
const dest = path.join(root, "public/office-js");

if (!fs.existsSync(src)) {
  console.warn("[copy-office-js] @microsoft/office-js 未安装，跳过");
  process.exit(0);
}

fs.rmSync(dest, { recursive: true, force: true });
fs.mkdirSync(dest, { recursive: true });
fs.cpSync(src, dest, { recursive: true });
console.log("[copy-office-js] synced -> public/office-js");
