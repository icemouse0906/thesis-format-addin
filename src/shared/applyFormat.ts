import type {
  ApplyOptions,
  FormatRules,
  ParagraphSpec,
  StyleRule,
} from "./schema";
import { getWordApi } from "./officeHost";
import {
  Align,
  StyleBuiltIn,
  StyleLocalized,
  setParagraphStyle,
  type StyleKind,
} from "./wordConsts";
import { detectCaption, detectHeadingLevel } from "./headingDetect";

function W(): typeof Word {
  return getWordApi();
}

export type ApplyReport = {
  ok: boolean;
  steps: string[];
  warnings: string[];
  error?: string;
};

const CM_TO_POINT = 28.35;
const A4_WIDTH_PT = 595.3;
const A4_HEIGHT_PT = 841.9;

type ParaProxy = Word.Paragraph & {
  alignment?: string;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  firstLineIndent?: number;
  leftIndent?: number;
  font?: Word.Font;
  space1?: () => void;
  space1Pt5?: () => void;
  space2?: () => void;
  paragraphFormat?: Word.ParagraphFormat;
};

function alignmentValue(
  a: ParagraphSpec["alignment"]
): string | undefined {
  switch (a) {
    case "left":
      return Align.left;
    case "center":
      return Align.centered;
    case "right":
      return Align.right;
    case "justify":
      return Align.justified;
    default:
      return undefined;
  }
}

/** 安全写属性：任一失败不抛出 */
function safeSet(label: string, fn: () => void, warnings: string[]) {
  try {
    fn();
  } catch (e) {
    warnings.push(`${label}: ${String(e)}`);
  }
}

/**
 * Mac/Win 通用：直接写 Paragraph 上的属性（不要用 paragraphFormat，Mac 上常为 undefined）
 * 参考：paragraph.alignment / paragraph.font / paragraph.space1Pt5()
 */
function formatParagraphDirect(
  paragraph: Word.Paragraph,
  rule: StyleRule,
  warnings: string[]
) {
  const p = paragraph as ParaProxy;
  const fontSpec = rule.font;
  const paraSpec = rule.paragraph;

  // —— 字体：优先 paragraph.font，其次 range.font ——
  const applyFontTo = (font: Word.Font | undefined, tag: string) => {
    if (!font) return false;
    safeSet(`${tag}.name`, () => {
      font.name = fontSpec.name;
      font.nameFarEast = fontSpec.name;
      if (fontSpec.nameAscii) font.nameAscii = fontSpec.nameAscii;
    }, warnings);
    safeSet(`${tag}.size`, () => {
      font.size = fontSpec.sizePt;
    }, warnings);
    if (fontSpec.bold != null) {
      safeSet(`${tag}.bold`, () => {
        font.bold = fontSpec.bold!;
      }, warnings);
    }
    if (fontSpec.italic != null) {
      safeSet(`${tag}.italic`, () => {
        font.italic = fontSpec.italic!;
      }, warnings);
    }
    return true;
  };

  if (!applyFontTo(p.font, "paragraph.font")) {
    try {
      const range = paragraph.getRange("Whole");
      applyFontTo(range.font, "range.font");
    } catch (e) {
      warnings.push(`font: ${String(e)}`);
    }
  }

  // —— 段落：直接属性 ——
  const align = alignmentValue(paraSpec.alignment);
  if (align != null) {
    safeSet("alignment", () => {
      (p as unknown as { alignment: string }).alignment = align;
    }, warnings);
  }

  if (paraSpec.firstLineIndentPt != null) {
    safeSet("firstLineIndent", () => {
      p.firstLineIndent = paraSpec.firstLineIndentPt!;
    }, warnings);
  }
  if (paraSpec.leftIndentPt != null) {
    safeSet("leftIndent", () => {
      p.leftIndent = paraSpec.leftIndentPt!;
    }, warnings);
  }
  if (paraSpec.spaceBeforePt != null) {
    safeSet("spaceBefore", () => {
      p.spaceBefore = paraSpec.spaceBeforePt!;
    }, warnings);
  }
  if (paraSpec.spaceAfterPt != null) {
    safeSet("spaceAfter", () => {
      p.spaceAfter = paraSpec.spaceAfterPt!;
    }, warnings);
  }

  // 行距：优先专用方法，再退回 lineSpacing（值为倍数*12）
  if (paraSpec.lineSpacing != null) {
    const mult = paraSpec.lineSpacing;
    if (Math.abs(mult - 1.5) < 0.05 && typeof p.space1Pt5 === "function") {
      safeSet("space1Pt5", () => p.space1Pt5!(), warnings);
    } else if (Math.abs(mult - 1) < 0.05 && typeof p.space1 === "function") {
      safeSet("space1", () => p.space1!(), warnings);
    } else if (Math.abs(mult - 2) < 0.05 && typeof p.space2 === "function") {
      safeSet("space2", () => p.space2!(), warnings);
    } else {
      // API：lineSpacing 单位为磅，UI 显示为 /12 的倍数
      safeSet("lineSpacing", () => {
        p.lineSpacing = mult * 12;
      }, warnings);
    }
  }

  // 若存在 paragraphFormat 再补写一遍（Win 上有时更完整）
  const pf = p.paragraphFormat as
    | (Word.ParagraphFormat & Record<string, unknown>)
    | undefined;
  if (pf) {
    if (align != null) {
      safeSet("pf.alignment", () => {
        (pf as { alignment: string }).alignment = align;
      }, warnings);
    }
    if (paraSpec.lineSpacing != null) {
      safeSet("pf.lineSpacing", () => {
        pf.lineSpacing = paraSpec.lineSpacing! * 12;
      }, warnings);
    }
    if (paraSpec.firstLineIndentPt != null) {
      safeSet("pf.firstLineIndent", () => {
        pf.firstLineIndent = paraSpec.firstLineIndentPt!;
      }, warnings);
    }
  }
}

async function ensureAndStyleBuiltIn(
  context: Word.RequestContext,
  kind: StyleKind,
  rule: StyleRule,
  steps: string[],
  label: string
) {
  const doc = context.document as Word.Document & {
    getStyles?: () => Word.StyleCollection;
  };

  if (typeof doc.getStyles !== "function") {
    steps.push(`跳过样式表「${label}」（无 getStyles，改用段落直接格式）`);
    return;
  }

  const names = [...new Set([StyleBuiltIn[kind], ...StyleLocalized[kind]])];
  try {
    let style: Word.Style | null = null;
    for (const name of names) {
      const s = doc.getStyles().getByNameOrNullObject(name);
      s.load("isNullObject");
      await context.sync();
      if (!s.isNullObject) {
        style = s;
        break;
      }
    }
    if (!style) {
      steps.push(`跳过 ${label}：未找到样式`);
      return;
    }

    style.font.name = rule.font.name;
    style.font.nameFarEast = rule.font.name;
    if (rule.font.nameAscii) style.font.nameAscii = rule.font.nameAscii;
    style.font.size = rule.font.sizePt;
    if (rule.font.bold != null) style.font.bold = rule.font.bold;
    await context.sync();
    steps.push(`已更新样式表：${label}`);
  } catch (e) {
    steps.push(`样式表「${label}」跳过：${String(e)}`);
  }
}


function isHeadingStyleName(name: string, level?: 1 | 2 | 3): boolean {
  const n = name.toLowerCase();
  if (level === 1) return n.includes("heading 1") || n.includes("标题 1");
  if (level === 2) return n.includes("heading 2") || n.includes("标题 2");
  if (level === 3) return n.includes("heading 3") || n.includes("标题 3");
  return n.includes("heading") || n.includes("标题");
}

async function getTargetParagraphs(
  context: Word.RequestContext,
  selectionOnly: boolean
): Promise<Word.Paragraph[]> {
  if (selectionOnly) {
    const paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load("items");
    await context.sync();
    return paragraphs.items;
  }
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  return paragraphs.items.slice(0, 2000);
}

async function autoMapHeadings(
  context: Word.RequestContext,
  paragraphs: Word.Paragraph[],
  steps: string[],
  warnings: string[]
) {
  for (const p of paragraphs) p.load("text");
  await context.sync();

  let h1 = 0;
  let h2 = 0;
  let h3 = 0;

  for (const p of paragraphs) {
    const level = detectHeadingLevel(p.text || "");
    if (!level) continue;
    const kind: StyleKind =
      level === 1 ? "heading1" : level === 2 ? "heading2" : "heading3";
    const ok = await setParagraphStyle(p, kind, context);
    if (level === 1) h1++;
    else if (level === 2) h2++;
    else h3++;
    if (!ok) {
      warnings.push(`标题样式名不可用（将直接格式化）：${(p.text || "").slice(0, 20)}`);
    }
  }
  steps.push(`自动识别标题：H1=${h1}, H2=${h2}, H3=${h3}`);
}

async function clearDirectFormatting(
  context: Word.RequestContext,
  paragraphs: Word.Paragraph[],
  steps: string[]
) {
  for (const p of paragraphs) p.load("style");
  await context.sync();

  let n = 0;
  for (const p of paragraphs) {
    try {
      const current = (p.style || "").trim();
      if (current) {
        p.style = current;
        n++;
      }
    } catch {
      /* ignore */
    }
  }
  await context.sync();
  steps.push(`已整理 ${n} 段样式引用`);
}

async function removeExtraBlankLines(
  context: Word.RequestContext,
  paragraphs: Word.Paragraph[],
  steps: string[]
) {
  for (const p of paragraphs) p.load("text");
  await context.sync();

  let blankRun = 0;
  let removed = 0;
  const toDelete: Word.Paragraph[] = [];

  for (const p of paragraphs) {
    const empty = !p.text || p.text.replace(/\r/g, "").trim() === "";
    if (empty) {
      blankRun++;
      if (blankRun > 1) toDelete.push(p);
    } else blankRun = 0;
  }

  for (const p of toDelete) {
    try {
      p.delete();
      removed++;
    } catch {
      /* ignore */
    }
  }
  await context.sync();
  steps.push(`已压缩多余空行：删除 ${removed} 段`);
}

async function applyPageSetup(
  context: Word.RequestContext,
  rules: FormatRules,
  options: ApplyOptions,
  steps: string[],
  warnings: string[]
) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  let applied = 0;
  for (const section of sections.items) {
    const pageSetup = (section as unknown as { pageSetup?: Word.PageSetup })
      .pageSetup;
    if (!pageSetup) continue;
    try {
      if (options.applyPage) {
        pageSetup.topMargin = rules.page.marginTopCm * CM_TO_POINT;
        pageSetup.bottomMargin = rules.page.marginBottomCm * CM_TO_POINT;
        pageSetup.leftMargin = rules.page.marginLeftCm * CM_TO_POINT;
        pageSetup.rightMargin = rules.page.marginRightCm * CM_TO_POINT;
      }
      if (options.applyPaperSize && rules.page.paper === "A4") {
        pageSetup.pageWidth = A4_WIDTH_PT;
        pageSetup.pageHeight = A4_HEIGHT_PT;
      }
      applied++;
    } catch {
      /* ignore */
    }
  }

  if (applied === 0) {
    warnings.push(
      `页边距需手动设置：上${rules.page.marginTopCm} 下${rules.page.marginBottomCm} 左${rules.page.marginLeftCm} 右${rules.page.marginRightCm} cm（版式→页边距）`
    );
    steps.push("页边距/纸张：当前 Word 无 pageSetup，已跳过");
    return;
  }
  await context.sync();
  steps.push(
    `页边距已应用：上${rules.page.marginTopCm} 下${rules.page.marginBottomCm} 左${rules.page.marginLeftCm} 右${rules.page.marginRightCm} cm`
  );
}

async function applyHeaderFooter(
  context: Word.RequestContext,
  rules: FormatRules,
  steps: string[],
  warnings: string[]
) {
  const hf = rules.headerFooter;
  if (!hf) return;

  /**
   * 重要：Mac Word 上 footer.clear() + insertOoxml(PAGE域) 会触发 GeneralException，
   * 并可能导致「第x页内容不可用 / 文档打不开」。
   * 因此这里绝不写 OOXML 页码域；仅在有页眉纯文本时做安全写入。
   */
  if (hf.pageNumber?.enabled) {
    warnings.push(
      "页码未自动插入（Mac 上自动写页码易损坏文档）。请手动：插入 → 页码 → 页面底端。"
    );
    steps.push("页码：已跳过（安全策略）");
  }

  if (!hf.headerText) {
    return;
  }

  try {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (const section of sections.items) {
      try {
        const header = section.getHeader("Primary");
        // 不 clear 整节页眉，改为在末尾追加一行，降低破坏风险
        const para = header.insertParagraph(hf.headerText, "End");
        (para as unknown as { alignment: string }).alignment = "Centered";
        await context.sync();
      } catch (e) {
        warnings.push(`页眉写入跳过：${String(e)}`);
      }
    }
    steps.push(`页眉文本已尝试写入：${hf.headerText}`);
  } catch (e) {
    warnings.push(`页眉处理跳过：${String(e)}`);
  }
}

async function restyleAllParagraphs(
  context: Word.RequestContext,
  paragraphs: Word.Paragraph[],
  rules: FormatRules,
  options: ApplyOptions,
  steps: string[],
  warnings: string[]
) {
  for (const p of paragraphs) {
    p.load("style,text");
  }
  await context.sync();

  const h1Rule: StyleRule = {
    ...rules.styles.heading1,
    paragraph: {
      ...rules.styles.heading1.paragraph,
      pageBreakBefore:
        options.pageBreakBeforeH1 ||
        rules.styles.heading1.paragraph.pageBreakBefore,
    },
  };

  let n = 0;
  let fail = 0;

  for (const p of paragraphs) {
    const styleName = p.style || "";
    const text = p.text || "";
    const detected = detectHeadingLevel(text);

    let rule: StyleRule | null = null;
    if (options.applyHeadings && (isHeadingStyleName(styleName, 1) || detected === 1)) {
      rule = h1Rule;
    } else if (
      options.applyHeadings &&
      (isHeadingStyleName(styleName, 2) || detected === 2)
    ) {
      rule = rules.styles.heading2;
    } else if (
      options.applyHeadings &&
      (isHeadingStyleName(styleName, 3) || detected === 3)
    ) {
      rule = rules.styles.heading3;
    } else if (
      options.applyCaptions &&
      rules.caption &&
      (detectCaption(text) || styleName.includes("题注") || /caption/i.test(styleName))
    ) {
      rule = rules.caption;
    } else if (options.applyNormal && !detected) {
      rule = rules.styles.normal;
    }

    if (!rule) continue;

    try {
      formatParagraphDirect(p, rule, warnings);
      n++;
    } catch (e) {
      fail++;
      warnings.push(`段落失败：${text.slice(0, 16)} → ${String(e)}`);
    }
  }

  // 分批 sync，避免一次队列过大
  await context.sync();
  steps.push(`已直接格式化 ${n} 个段落` + (fail ? `（${fail} 段失败）` : ""));
}

async function applyCaptionsToParagraphs(
  context: Word.RequestContext,
  paragraphs: Word.Paragraph[],
  rules: FormatRules,
  steps: string[],
  warnings: string[]
) {
  if (!rules.caption) return;
  for (const p of paragraphs) p.load("text");
  await context.sync();

  let n = 0;
  for (const p of paragraphs) {
    if (!detectCaption(p.text || "")) continue;
    await setParagraphStyle(p, "caption", context);
    formatParagraphDirect(p, rules.caption, warnings);
    n++;
  }
  await context.sync();
  steps.push(`题注格式：${n} 段`);
}

export { isOfficeWordHost } from "./officeHost";

export async function applyFormatToDocument(
  rules: FormatRules,
  options: ApplyOptions
): Promise<ApplyReport> {
  const steps: string[] = [];
  const warnings: string[] = [];

  try {
    getWordApi();
  } catch (e) {
    return {
      ok: false,
      steps,
      warnings,
      error: e instanceof Error ? e.message : String(e),
    };
  }

  try {
    return await W().run(async (context) => {
      steps.push(`开始套用「${rules.name}」`);
      steps.push(options.selectionOnly ? "范围：仅选区" : "范围：全文");

      let paragraphs = await getTargetParagraphs(
        context,
        options.selectionOnly
      );

      if (options.clearDirectFormatting) {
        await clearDirectFormatting(context, paragraphs, steps);
      }

      if (options.autoDetectHeadings) {
        await autoMapHeadings(context, paragraphs, steps, warnings);
        // 标题样式可能已变，重新取段落
        paragraphs = await getTargetParagraphs(context, options.selectionOnly);
      }

      if ((options.applyPage || options.applyPaperSize) && !options.selectionOnly) {
        try {
          await applyPageSetup(context, rules, options, steps, warnings);
        } catch (e) {
          warnings.push(`页面设置：${String(e)}`);
        }
      }

      // 样式表能改就改，不能改也不阻断
      if (options.applyNormal) {
        await ensureAndStyleBuiltIn(
          context,
          "normal",
          rules.styles.normal,
          steps,
          "正文"
        );
      }
      if (options.applyHeadings) {
        await ensureAndStyleBuiltIn(
          context,
          "heading1",
          rules.styles.heading1,
          steps,
          "标题1"
        );
        await ensureAndStyleBuiltIn(
          context,
          "heading2",
          rules.styles.heading2,
          steps,
          "标题2"
        );
        await ensureAndStyleBuiltIn(
          context,
          "heading3",
          rules.styles.heading3,
          steps,
          "标题3"
        );
      }

      // 核心路径：段落直接属性（Mac 可用）
      if (options.applyNormal || options.applyHeadings || options.applyCaptions) {
        paragraphs = await getTargetParagraphs(context, options.selectionOnly);
        await restyleAllParagraphs(
          context,
          paragraphs,
          rules,
          options,
          steps,
          warnings
        );
      }

      if (options.applyCaptions && rules.caption) {
        paragraphs = await getTargetParagraphs(context, options.selectionOnly);
        await applyCaptionsToParagraphs(
          context,
          paragraphs,
          rules,
          steps,
          warnings
        );
      }

      if (options.removeExtraBlankLines) {
        paragraphs = await getTargetParagraphs(context, options.selectionOnly);
        await removeExtraBlankLines(context, paragraphs, steps);
      }

      if (options.applyHeaderFooter && rules.headerFooter && !options.selectionOnly) {
        await applyHeaderFooter(context, rules, steps, warnings);
      }

      // 过滤噪音警告：同类型只保留前几条
      const trimmedWarnings = trimWarnings(warnings, 12);

      steps.push("套用完成");
      return { ok: true, steps, warnings: trimmedWarnings };
    });
  } catch (e) {
    return {
      ok: false,
      steps,
      warnings: trimWarnings(warnings, 12),
      error: e instanceof Error ? e.message : String(e),
    };
  }
}

function trimWarnings(warnings: string[], max: number): string[] {
  if (warnings.length <= max) return warnings;
  return [
    ...warnings.slice(0, max),
    `…另有 ${warnings.length - max} 条类似提示已省略`,
  ];
}

export function simulateApply(
  rules: FormatRules,
  options: ApplyOptions
): ApplyReport {
  return {
    ok: true,
    steps: [
      `【预览】${rules.name}`,
      options.applyNormal ? "正文直接格式" : "",
      options.applyHeadings ? "标题直接格式 + 自动识别" : "",
      options.applyPage ? "页边距（若 API 可用）" : "",
    ].filter(Boolean),
    warnings: ["浏览器预览不会改文档"],
  };
}
