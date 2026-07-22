import { getWordApi } from "./officeHost";
import type { FormatRules } from "./schema";
import { detectCaption, detectHeadingLevel } from "./headingDetect";

export type OutlineItem = {
  level: 1 | 2 | 3 | "caption" | "body";
  text: string;
  index: number;
};

export type StructureReport = {
  ok: boolean;
  paragraphCount: number;
  heading1: number;
  heading2: number;
  heading3: number;
  captions: number;
  outline: OutlineItem[];
  error?: string;
};

export type ComplianceItem = {
  id: string;
  label: string;
  status: "pass" | "warn" | "fail" | "skip";
  detail: string;
};

export type ComplianceReport = {
  ok: boolean;
  score: number;
  items: ComplianceItem[];
  error?: string;
};

function W() {
  return getWordApi();
}

/** 扫描文档结构大纲（装逼 + 实用） */
export async function scanDocumentStructure(): Promise<StructureReport> {
  try {
    return await W().run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const items = paragraphs.items.slice(0, 1500);
      for (const p of items) p.load("text");
      await context.sync();

      let heading1 = 0;
      let heading2 = 0;
      let heading3 = 0;
      let captions = 0;
      const outline: OutlineItem[] = [];

      items.forEach((p, index) => {
        const text = (p.text || "").replace(/\r/g, "").trim();
        if (!text) return;
        if (detectCaption(text)) {
          captions++;
          if (outline.length < 40) {
            outline.push({ level: "caption", text: text.slice(0, 48), index });
          }
          return;
        }
        const level = detectHeadingLevel(text);
        if (level === 1) {
          heading1++;
          if (outline.length < 40) outline.push({ level: 1, text: text.slice(0, 48), index });
        } else if (level === 2) {
          heading2++;
          if (outline.length < 40) outline.push({ level: 2, text: text.slice(0, 48), index });
        } else if (level === 3) {
          heading3++;
          if (outline.length < 40) outline.push({ level: 3, text: text.slice(0, 48), index });
        }
      });

      return {
        ok: true,
        paragraphCount: items.length,
        heading1,
        heading2,
        heading3,
        captions,
        outline,
      };
    });
  } catch (e) {
    return {
      ok: false,
      paragraphCount: 0,
      heading1: 0,
      heading2: 0,
      heading3: 0,
      captions: 0,
      outline: [],
      error: e instanceof Error ? e.message : String(e),
    };
  }
}

type ParaProbe = Word.Paragraph & {
  font?: { size?: number; name?: string; bold?: boolean };
  firstLineIndent?: number;
  lineSpacing?: number;
  alignment?: string;
};

/**
 * 对照当前规则做抽样合规检查（读文档实际字体/缩进等）
 */
export async function checkCompliance(
  rules: FormatRules
): Promise<ComplianceReport> {
  const items: ComplianceItem[] = [];

  try {
    return await W().run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const list = paragraphs.items.slice(0, 800);
      for (const p of list) {
        p.load("text,font/size,font/name,font/bold,firstLineIndent,lineSpacing,alignment");
      }
      await context.sync();

      let bodyChecked = 0;
      let bodySizeOk = 0;
      let bodyIndentOk = 0;
      let h1Checked = 0;
      let h1BoldOk = 0;
      let captionCount = 0;

      const expectBody = rules.styles.normal.font.sizePt;
      const expectIndent = rules.styles.normal.paragraph.firstLineIndentPt ?? 0;
      const expectH1Size = rules.styles.heading1.font.sizePt;

      for (const raw of list) {
        const p = raw as ParaProbe;
        const text = (p.text || "").replace(/\r/g, "").trim();
        if (!text) continue;

        if (detectCaption(text)) {
          captionCount++;
          continue;
        }

        const level = detectHeadingLevel(text);
        const size = p.font?.size;
        const bold = p.font?.bold;
        const indent = p.firstLineIndent;

        if (level === 1) {
          h1Checked++;
          if (bold === true || (typeof size === "number" && Math.abs(size - expectH1Size) <= 1.5)) {
            h1BoldOk++;
          }
          continue;
        }
        if (level) continue;

        bodyChecked++;
        if (typeof size === "number" && Math.abs(size - expectBody) <= 0.6) bodySizeOk++;
        if (
          typeof indent === "number" &&
          (expectIndent === 0
            ? Math.abs(indent) < 2
            : Math.abs(indent - expectIndent) <= 6)
        ) {
          bodyIndentOk++;
        }
      }

      const push = (
        id: string,
        label: string,
        status: ComplianceItem["status"],
        detail: string
      ) => items.push({ id, label, status, detail });

      if (bodyChecked === 0) {
        push("body", "正文字号", "skip", "未采样到正文段落");
      } else {
        const ratio = bodySizeOk / bodyChecked;
        push(
          "body-size",
          `正文字号≈${expectBody}pt`,
          ratio >= 0.7 ? "pass" : ratio >= 0.4 ? "warn" : "fail",
          `${bodySizeOk}/${bodyChecked} 段接近目标`
        );
        const ir = bodyIndentOk / bodyChecked;
        push(
          "body-indent",
          "正文首行缩进",
          ir >= 0.6 ? "pass" : ir >= 0.3 ? "warn" : "fail",
          `${bodyIndentOk}/${bodyChecked} 段接近目标`
        );
      }

      if (h1Checked === 0) {
        push("h1", "一级标题样式", "warn", "未识别到一级标题（可开自动识别后再查）");
      } else {
        const hr = h1BoldOk / h1Checked;
        push(
          "h1",
          `一级标题（目标约 ${expectH1Size}pt/加粗）`,
          hr >= 0.6 ? "pass" : "warn",
          `${h1BoldOk}/${h1Checked} 段看起来像标题`
        );
      }

      push(
        "caption",
        "图/表题注",
        captionCount > 0 ? "pass" : "skip",
        captionCount > 0 ? `发现 ${captionCount} 条题注候选` : "未发现「图X/表X」题注"
      );

      push(
        "margin",
        "页边距",
        "skip",
        `目标 上${rules.page.marginTopCm}/下${rules.page.marginBottomCm}/左${rules.page.marginLeftCm}/右${rules.page.marginRightCm} cm（Mac 常需手动）`
      );

      push(
        "page-number",
        "页码",
        "skip",
        "建议手动：插入 → 页码（插件默认不自动写页码）"
      );

      const scored = items.filter((i) => i.status === "pass" || i.status === "fail" || i.status === "warn");
      const points = scored.reduce((s, i) => {
        if (i.status === "pass") return s + 1;
        if (i.status === "warn") return s + 0.5;
        return s;
      }, 0);
      const score =
        scored.length === 0 ? 0 : Math.round((points / scored.length) * 100);

      return { ok: true, score, items };
    });
  } catch (e) {
    return {
      ok: false,
      score: 0,
      items,
      error: e instanceof Error ? e.message : String(e),
    };
  }
}
