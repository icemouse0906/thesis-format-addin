/**
 * Mac Word 上 Word.BuiltInStyleName / Alignment 等枚举经常是 undefined，
 * 直接使用官方字符串值更稳妥。
 */

export const Align = {
  left: "Left",
  centered: "Centered",
  right: "Right",
  justified: "Justified",
} as const;

export const StyleBuiltIn = {
  normal: "Normal",
  heading1: "Heading 1",
  heading2: "Heading 2",
  heading3: "Heading 3",
  caption: "Caption",
} as const;

/** 中文版 Word 常见样式名（优先试本地化名） */
export const StyleLocalized = {
  normal: ["正文", "Normal", "Body Text"],
  heading1: ["标题 1", "标题1", "Heading 1", "Heading1"],
  heading2: ["标题 2", "标题2", "Heading 2", "Heading2"],
  heading3: ["标题 3", "标题3", "Heading 3", "Heading3"],
  caption: ["题注", "Caption"],
} as const;

export const RangeLoc = {
  whole: "Whole",
  start: "Start",
  end: "End",
} as const;

export const InsertLoc = {
  start: "Start",
  end: "End",
  before: "Before",
  after: "After",
  replace: "Replace",
} as const;

export const HeaderFooterKind = {
  primary: "Primary",
  firstPage: "FirstPage",
  evenPages: "EvenPages",
} as const;

export type StyleKind = keyof typeof StyleBuiltIn;

/** 给段落设置样式；优先中文名，失败再试英文 */
export async function setParagraphStyle(
  paragraph: Word.Paragraph,
  kind: StyleKind,
  context: Word.RequestContext
): Promise<boolean> {
  const names = [...new Set([...StyleLocalized[kind], StyleBuiltIn[kind]])];

  for (const name of names) {
    try {
      paragraph.style = name;
      await context.sync();
      return true;
    } catch {
      /* 试下一个名字 */
    }
  }
  return false;
}
