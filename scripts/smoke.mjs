import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const parseSrc = readFileSync(path.join(root, "src/shared/parseRequirements.ts"), "utf8");
const templatesSrc = readFileSync(path.join(root, "src/shared/templates.ts"), "utf8");

for (const [token, hay] of [
  ["export function parseRequirementsToRules", parseSrc],
  ["export const BUILTIN_TEMPLATES", templatesSrc],
  ["generic-cn-thesis", templatesSrc],
]) {
  if (!hay.includes(token)) {
    console.error("FAIL missing", token);
    process.exit(1);
  }
}

function size(text) {
  if (/小四/.test(text)) return 12;
  if (/三号/.test(text) && !/小三/.test(text)) return 16;
  return null;
}

if (size("正文：宋体小四") !== 12 || size("一级标题：黑体三号") !== 16) {
  console.error("FAIL size mapping");
  process.exit(1);
}

console.log("PASS smoke checks");
