/** 论文格式规则 Schema —— 模板与自然语言解析的统一中间层 */

export type FontSpec = {
  name: string;
  nameAscii?: string;
  sizePt: number;
  bold?: boolean;
  italic?: boolean;
};

export type ParagraphSpec = {
  alignment?: "left" | "center" | "right" | "justify";
  lineSpacing?: number; // 倍数，如 1.5
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  firstLineIndentPt?: number;
  leftIndentPt?: number;
  pageBreakBefore?: boolean;
};

export type StyleRule = {
  font: FontSpec;
  paragraph: ParagraphSpec;
};

export type PageRule = {
  paper?: "A4" | "Letter";
  marginTopCm: number;
  marginBottomCm: number;
  marginLeftCm: number;
  marginRightCm: number;
  gutterCm?: number;
};

export type PageNumberRule = {
  enabled: boolean;
  startAt?: number;
  alignment?: "left" | "center" | "right";
  format?: "arabic" | "roman";
};

export type HeaderFooterRule = {
  headerText?: string;
  footerText?: string;
  pageNumber?: PageNumberRule;
};

export type CaptionRule = {
  font: FontSpec;
  paragraph: ParagraphSpec;
};

export type FormatRules = {
  id: string;
  name: string;
  school?: string;
  description?: string;
  page: PageRule;
  styles: {
    normal: StyleRule;
    heading1: StyleRule;
    heading2: StyleRule;
    heading3: StyleRule;
  };
  /** 图题 / 表题 */
  caption?: CaptionRule;
  headerFooter?: HeaderFooterRule;
  notes?: string[];
};

export type ApplyOptions = {
  /** 仅处理当前选区；关闭则处理全文 */
  selectionOnly: boolean;
  applyPage: boolean;
  applyPaperSize: boolean;
  applyNormal: boolean;
  applyHeadings: boolean;
  applyHeaderFooter: boolean;
  autoDetectHeadings: boolean;
  /** 套用前清除选区/正文直接格式，避免旧字号残留 */
  clearDirectFormatting: boolean;
  /** 压缩连续空行 */
  removeExtraBlankLines: boolean;
  /** 图表题注样式 */
  applyCaptions: boolean;
  /** 一级标题前强制分页 */
  pageBreakBeforeH1: boolean;
};

export const DEFAULT_APPLY_OPTIONS: ApplyOptions = {
  selectionOnly: false,
  applyPage: true,
  applyPaperSize: true,
  applyNormal: true,
  applyHeadings: true,
  // Mac 上写页码 OOXML 易损坏文档，默认关闭
  applyHeaderFooter: false,
  autoDetectHeadings: true,
  clearDirectFormatting: true,
  removeExtraBlankLines: false,
  applyCaptions: true,
  pageBreakBeforeH1: false,
};
