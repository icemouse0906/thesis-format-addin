import type { FormatRules } from "./schema";

type RecordValue = Record<string, unknown>;
function object(value: unknown, path: string): RecordValue {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error(`${path} 必须是对象`);
  return value as RecordValue;
}
function text(value: unknown, path: string, required = false) {
  if (value === undefined && !required) return;
  if (typeof value !== "string" || !value.trim() || value.length > 2000) throw new Error(`${path} 必须是 1–2000 字的文本`);
}
function number(value: unknown, path: string, min: number, max: number, required = false) {
  if (value === undefined && !required) return;
  if (typeof value !== "number" || !Number.isFinite(value) || value < min || value > max) throw new Error(`${path} 必须在 ${min}–${max} 之间`);
}
function choice(value: unknown, path: string, values: string[]) {
  if (value !== undefined && (typeof value !== "string" || !values.includes(value))) throw new Error(`${path} 取值无效`);
}
function boolean(value: unknown, path: string) {
  if (value !== undefined && typeof value !== "boolean") throw new Error(`${path} 必须为布尔值`);
}
function style(value: unknown, path: string) {
  const rule = object(value, path);
  const font = object(rule.font, `${path}.font`);
  text(font.name, `${path}.font.name`, true);
  text(font.nameAscii, `${path}.font.nameAscii`);
  number(font.sizePt, `${path}.font.sizePt`, 1, 1638, true);
  for (const key of ["bold", "italic"]) boolean(font[key], `${path}.font.${key}`);
  const p = object(rule.paragraph, `${path}.paragraph`);
  choice(p.alignment, `${path}.paragraph.alignment`, ["left", "center", "right", "justify"]);
  number(p.lineSpacing, `${path}.paragraph.lineSpacing`, 0.5, 10);
  for (const key of ["spaceBeforePt", "spaceAfterPt"]) number(p[key], `${path}.paragraph.${key}`, 0, 1584);
  for (const key of ["firstLineIndentPt", "leftIndentPt"]) number(p[key], `${path}.paragraph.${key}`, -1584, 1584);
  boolean(p.pageBreakBefore, `${path}.paragraph.pageBreakBefore`);
}

/** 在 JSON 导入、AI 输出与 Word 写入边界拒绝非法规则，避免部分套用后才报错。 */
export function validateRules(value: unknown): asserts value is FormatRules {
  const r = object(value, "规则");
  text(r.id, "id", true);
  text(r.name, "name", true);
  text(r.school, "school");
  text(r.description, "description");
  const page = object(r.page, "page");
  choice(page.paper, "page.paper", ["A4", "Letter"]);
  for (const key of ["marginTopCm", "marginBottomCm", "marginLeftCm", "marginRightCm"]) number(page[key], `page.${key}`, 0, 20, true);
  number(page.gutterCm, "page.gutterCm", 0, 10);
  const width = page.paper === "Letter" ? 21.59 : 21;
  const height = page.paper === "Letter" ? 27.94 : 29.7;
  if (Number(page.marginLeftCm) + Number(page.marginRightCm) + Number(page.gutterCm ?? 0) >= width ||
      Number(page.marginTopCm) + Number(page.marginBottomCm) >= height) throw new Error("页边距不能占满整张纸");
  const styles = object(r.styles, "styles");
  for (const key of ["normal", "heading1", "heading2", "heading3"]) style(styles[key], `styles.${key}`);
  if (r.caption !== undefined) style(r.caption, "caption");
  if (r.notes !== undefined && (!Array.isArray(r.notes) || r.notes.some((v) => typeof v !== "string"))) throw new Error("notes 必须是文本数组");
  if (r.headerFooter !== undefined) {
    const hf = object(r.headerFooter, "headerFooter");
    for (const key of ["headerText", "footerText"]) {
      if (hf[key] !== "") text(hf[key], `headerFooter.${key}`);
    }
    if (hf.pageNumber !== undefined) {
      const pn = object(hf.pageNumber, "headerFooter.pageNumber");
      boolean(pn.enabled, "pageNumber.enabled");
      number(pn.startAt, "pageNumber.startAt", 0, 32767);
      if (pn.startAt !== undefined && !Number.isInteger(pn.startAt)) throw new Error("页码起始值必须为整数");
      choice(pn.alignment, "pageNumber.alignment", ["left", "center", "right"]);
      choice(pn.format, "pageNumber.format", ["arabic", "roman"]);
    }
  }
}
