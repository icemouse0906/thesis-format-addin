import type { FormatRules, StyleRule } from "./schema";
import { genericCnThesis } from "./templates";

export type ParseResult = {
  rules: FormatRules;
  matched: string[];
  unmatched: string[];
};

const PT = /([\d.]+)\s*(?:pt|磅)/i;

function deepClone<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

function pickFontSizePt(text: string): number | undefined {
  if (/小四/.test(text)) return 12;
  if (/四号/.test(text) && !/小四/.test(text)) return 14;
  if (/小三/.test(text)) return 15;
  if (/三号/.test(text) && !/小三/.test(text)) return 16;
  if (/小二/.test(text)) return 18;
  if (/二号/.test(text) && !/小二/.test(text)) return 22;
  if (/五号/.test(text)) return 10.5;
  const m = text.match(PT);
  if (m) return Number(m[1]);
  return undefined;
}

function pickFontName(text: string): string | undefined {
  if (/黑体/.test(text)) return "黑体";
  if (/楷体|楷书/.test(text)) return "楷体";
  if (/仿宋/.test(text)) return "仿宋";
  if (/微软雅黑/.test(text)) return "微软雅黑";
  if (/宋体/.test(text)) return "宋体";
  if (/Times\s*New\s*Roman/i.test(text)) return "Times New Roman";
  if (/Arial/i.test(text)) return "Arial";
  return undefined;
}

function pickLineSpacing(text: string): number | undefined {
  const m = text.match(/([\d.]+)\s*倍(?:行距)?/);
  if (m) return Number(m[1]);
  if (/单倍行距|单倍/.test(text)) return 1;
  if (/固定值/.test(text)) {
    const pt = text.match(PT);
    if (pt) return Number(pt[1]) / 12; // 近似换算
  }
  if (/1\.5|1．5/.test(text) && /行距/.test(text)) return 1.5;
  if (/2\.0|2倍|两倍/.test(text) && /行距/.test(text)) return 2;
  return undefined;
}

function pickAlignment(text: string): StyleRule["paragraph"]["alignment"] | undefined {
  if (/居中/.test(text)) return "center";
  if (/两端对齐|左右对齐/.test(text)) return "justify";
  if (/右对齐/.test(text)) return "right";
  if (/左对齐/.test(text)) return "left";
  return undefined;
}

function applyToStyle(style: StyleRule, chunk: string, matched: string[]) {
  const size = pickFontSizePt(chunk);
  if (size != null) {
    style.font.sizePt = size;
    matched.push(`字号 → ${size}pt`);
  }
  const font = pickFontName(chunk);
  if (font) {
    style.font.name = font;
    matched.push(`字体 → ${font}`);
  }
  if (/加粗|粗体|bold/i.test(chunk)) {
    style.font.bold = true;
    matched.push("加粗");
  }
  const ls = pickLineSpacing(chunk);
  if (ls != null) {
    style.paragraph.lineSpacing = ls;
    matched.push(`行距 → ${ls} 倍`);
  }
  const align = pickAlignment(chunk);
  if (align) {
    style.paragraph.alignment = align;
    matched.push(`对齐 → ${align}`);
  }
  if (/首行缩进\s*2|缩进\s*两字符|缩进两字|首行缩进两/.test(chunk)) {
    style.paragraph.firstLineIndentPt = 24;
    matched.push("首行缩进 2 字符");
  } else {
    const indentMatch = chunk.match(/首行缩进\s*([\d.]+)/);
    if (indentMatch) {
      const n = Number(indentMatch[1]);
      style.paragraph.firstLineIndentPt = n * 12;
      matched.push(`首行缩进 ${n} 字符`);
    }
  }
  if (/无首行缩进|不缩进|不首行缩进/.test(chunk)) {
    style.paragraph.firstLineIndentPt = 0;
    matched.push("取消首行缩进");
  }
}

/**
 * 规则型解析：把用户粘贴的格式要求映射到 FormatRules。
 * 不依赖外部 LLM，今晚即可离线使用；后续可换成后端 LLM 填同一 Schema。
 */
export function parseRequirementsToRules(input: string): ParseResult {
  const text = input.trim();
  const rules = deepClone(genericCnThesis);
  rules.id = `custom-${Date.now()}`;
  rules.name = "自定义解析结果";
  rules.school = "用户输入";
  rules.description = "由格式要求文本自动解析生成（未匹配项保留通用模板默认值）。";
  rules.notes = [];

  const matched: string[] = [];
  const unmatched: string[] = [];

  if (!text) {
    return { rules, matched, unmatched: ["输入为空，使用通用模板默认值"] };
  }

  // 页边距：上/下/左/右
  const marginBlock = text.match(/页边距[\s\S]{0,80}/);
  if (marginBlock) {
    const block = marginBlock[0];
    const top = block.match(/上\s*([0-9.]+)/);
    const bottom = block.match(/下\s*([0-9.]+)/);
    const left = block.match(/左\s*([0-9.]+)/);
    const right = block.match(/右\s*([0-9.]+)/);
    if (top) {
      rules.page.marginTopCm = Number(top[1]);
      matched.push(`上边距 ${top[1]}cm`);
    }
    if (bottom) {
      rules.page.marginBottomCm = Number(bottom[1]);
      matched.push(`下边距 ${bottom[1]}cm`);
    }
    if (left) {
      rules.page.marginLeftCm = Number(left[1]);
      matched.push(`左边距 ${left[1]}cm`);
    }
    if (right) {
      rules.page.marginRightCm = Number(right[1]);
      matched.push(`右边距 ${right[1]}cm`);
    }
  } else {
    // 单独出现「左边距 3cm」等
    const pairs: Array<
      ["marginTopCm" | "marginBottomCm" | "marginLeftCm" | "marginRightCm", RegExp]
    > = [
      ["marginTopCm", /上(?:边距|页边距)?\s*[:：]?\s*([\d.]+)\s*(?:cm|厘米)?/i],
      ["marginBottomCm", /下(?:边距|页边距)?\s*[:：]?\s*([\d.]+)\s*(?:cm|厘米)?/i],
      ["marginLeftCm", /左(?:边距|页边距)?\s*[:：]?\s*([\d.]+)\s*(?:cm|厘米)?/i],
      ["marginRightCm", /右(?:边距|页边距)?\s*[:：]?\s*([\d.]+)\s*(?:cm|厘米)?/i],
    ];
    for (const [key, re] of pairs) {
      const m = text.match(re);
      if (m) {
        rules.page[key] = Number(m[1]);
        matched.push(`${key}=${m[1]}cm`);
      }
    }
  }

  // 全局行距
  const globalLs = pickLineSpacing(text);
  if (globalLs != null) {
    rules.styles.normal.paragraph.lineSpacing = globalLs;
    rules.styles.heading1.paragraph.lineSpacing = globalLs;
    rules.styles.heading2.paragraph.lineSpacing = globalLs;
    rules.styles.heading3.paragraph.lineSpacing = globalLs;
    matched.push(`全局行距 ${globalLs} 倍`);
  }

  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  for (const line of lines) {
    if (/正文|体例|正文字体|正文要求/.test(line) || (/宋体|小四/.test(line) && !/标题|章|节/.test(line))) {
      applyToStyle(rules.styles.normal, line, matched);
      continue;
    }
    if (/一级标题|章标题|标题\s*1|一级\s*标题/.test(line)) {
      applyToStyle(rules.styles.heading1, line, matched);
      continue;
    }
    if (/二级标题|节标题|标题\s*2|二级\s*标题/.test(line)) {
      applyToStyle(rules.styles.heading2, line, matched);
      continue;
    }
    if (/三级标题|标题\s*3|三级\s*标题/.test(line)) {
      applyToStyle(rules.styles.heading3, line, matched);
      continue;
    }
    if (/页眉/.test(line)) {
      const t = line.replace(/页眉[:：]?\s*/, "").trim();
      if (t && rules.headerFooter) {
        rules.headerFooter.headerText = t;
        matched.push(`页眉：${t}`);
      }
      continue;
    }
    if (/页码/.test(line)) {
      if (rules.headerFooter?.pageNumber) {
        rules.headerFooter.pageNumber.enabled = !/无页码|不要页码/.test(line);
        if (/居中/.test(line)) rules.headerFooter.pageNumber.alignment = "center";
        if (/右侧|靠右/.test(line)) rules.headerFooter.pageNumber.alignment = "right";
        matched.push("页码规则已更新");
      }
      continue;
    }
    // 未明确分类的行：若含格式关键词则尝试套到正文
    if (/字体|字号|行距|对齐|缩进|黑体|宋体|小四|三号/.test(line)) {
      applyToStyle(rules.styles.normal, line, matched);
    } else if (line.length > 4) {
      unmatched.push(line.slice(0, 80));
    }
  }

  // 若全文提到「A4」
  if (/\bA4\b|纸张.?A4|A4纸/i.test(text)) {
    rules.page.paper = "A4";
    matched.push("纸张 A4");
  }

  // 去重 matched
  const uniq = [...new Set(matched)];
  rules.notes = [
    `已解析 ${uniq.length} 条线索`,
    ...uniq.slice(0, 12),
    ...(unmatched.length ? [`有 ${unmatched.length} 行未能归类，已保留模板默认值`] : []),
  ];

  return { rules, matched: uniq, unmatched };
}
