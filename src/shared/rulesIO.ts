import type { FormatRules } from "./schema";
import { genericCnThesis } from "./templates";

export function exportRulesJson(rules: FormatRules): string {
  return JSON.stringify(rules, null, 2);
}

export function downloadRulesJson(rules: FormatRules, filename?: string) {
  const blob = new Blob([exportRulesJson(rules)], {
    type: "application/json;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename ?? `${rules.id || "thesis-rules"}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

export function importRulesJson(text: string): FormatRules {
  const data = JSON.parse(text) as Partial<FormatRules>;
  if (!data || typeof data !== "object") {
    throw new Error("不是有效的 JSON 对象");
  }
  if (!data.styles?.normal || !data.page) {
    throw new Error("缺少 page / styles.normal 等必要字段");
  }

  const base: FormatRules = JSON.parse(JSON.stringify(genericCnThesis));
  return {
    ...base,
    ...data,
    id: data.id || `import-${Date.now()}`,
    name: data.name || "导入的规则",
    page: { ...base.page, ...data.page },
    styles: {
      normal: { ...base.styles.normal, ...data.styles.normal, font: { ...base.styles.normal.font, ...data.styles?.normal?.font }, paragraph: { ...base.styles.normal.paragraph, ...data.styles?.normal?.paragraph } },
      heading1: { ...base.styles.heading1, ...data.styles.heading1, font: { ...base.styles.heading1.font, ...data.styles?.heading1?.font }, paragraph: { ...base.styles.heading1.paragraph, ...data.styles?.heading1?.paragraph } },
      heading2: { ...base.styles.heading2, ...data.styles.heading2, font: { ...base.styles.heading2.font, ...data.styles?.heading2?.font }, paragraph: { ...base.styles.heading2.paragraph, ...data.styles?.heading2?.paragraph } },
      heading3: { ...base.styles.heading3, ...data.styles.heading3, font: { ...base.styles.heading3.font, ...data.styles?.heading3?.font }, paragraph: { ...base.styles.heading3.paragraph, ...data.styles?.heading3?.paragraph } },
    },
    caption: data.caption
      ? {
          font: { ...base.caption!.font, ...data.caption.font },
          paragraph: { ...base.caption!.paragraph, ...data.caption.paragraph },
        }
      : base.caption,
    notes: data.notes ?? [`导入于 ${new Date().toLocaleString()}`],
  };
}

export async function readFileAsText(file: File): Promise<string> {
  return await file.text();
}
