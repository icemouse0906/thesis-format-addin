import type { FormatRules } from "./schema";

/** 通用中文硕博论文常用排版（可作多校起点） */
export const genericCnThesis: FormatRules = {
  id: "generic-cn-thesis",
  name: "通用中文学位论文",
  school: "通用模板",
  description: "覆盖多数高校常见要求：宋体小四正文、黑体标题、1.5 倍行距、A4 页边距。",
  page: {
    paper: "A4",
    marginTopCm: 2.5,
    marginBottomCm: 2.5,
    marginLeftCm: 3.0,
    marginRightCm: 2.5,
    gutterCm: 0,
  },
  styles: {
    normal: {
      font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 12 },
      paragraph: {
        alignment: "justify",
        lineSpacing: 1.5,
        firstLineIndentPt: 24,
        spaceBeforePt: 0,
        spaceAfterPt: 0,
      },
    },
    heading1: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 16, bold: true },
      paragraph: {
        alignment: "center",
        lineSpacing: 1.5,
        spaceBeforePt: 24,
        spaceAfterPt: 18,
        firstLineIndentPt: 0,
      },
    },
    heading2: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 14, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 18,
        spaceAfterPt: 12,
        firstLineIndentPt: 0,
      },
    },
    heading3: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 12, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        firstLineIndentPt: 0,
      },
    },
  },
  headerFooter: {
    headerText: "",
    pageNumber: {
      enabled: true,
      startAt: 1,
      alignment: "center",
      format: "arabic",
    },
  },
  caption: {
    font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 10.5 },
    paragraph: {
      alignment: "center",
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 6,
      firstLineIndentPt: 0,
    },
  },
  notes: [
    "正文：宋体小四（12pt）+ Times New Roman，两端对齐，首行缩进 2 字符，1.5 倍行距。",
    "一级标题：黑体三号居中；二级：黑体四号；三级：黑体小四。",
    "题注：宋体五号居中（图X / 表X）。",
    "页边距：上 2.5 / 下 2.5 / 左 3.0 / 右 2.5 cm（A4）。",
  ],
};

/** 偏工科常见：左装订稍宽、标题更醒目 */
export const engineeringThesis: FormatRules = {
  id: "engineering-thesis",
  name: "工科论文常用格式",
  school: "工科通用",
  description: "左页边距略宽便于装订；正文小四宋体；章标题黑体三号居中。",
  page: {
    paper: "A4",
    marginTopCm: 2.54,
    marginBottomCm: 2.54,
    marginLeftCm: 3.17,
    marginRightCm: 2.54,
    gutterCm: 0,
  },
  styles: {
    normal: {
      font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 12 },
      paragraph: {
        alignment: "justify",
        lineSpacing: 1.5,
        firstLineIndentPt: 24,
      },
    },
    heading1: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 16, bold: true },
      paragraph: {
        alignment: "center",
        lineSpacing: 1.5,
        spaceBeforePt: 24,
        spaceAfterPt: 18,
        firstLineIndentPt: 0,
      },
    },
    heading2: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 14, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        firstLineIndentPt: 0,
      },
    },
    heading3: {
      font: { name: "楷体", nameAscii: "Times New Roman", sizePt: 12, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 6,
        spaceAfterPt: 6,
        firstLineIndentPt: 0,
      },
    },
  },
  headerFooter: {
    pageNumber: {
      enabled: true,
      alignment: "center",
      format: "arabic",
    },
  },
  caption: {
    font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 10.5 },
    paragraph: {
      alignment: "center",
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 6,
      firstLineIndentPt: 0,
    },
  },
  notes: [
    "页边距接近 Word 默认「适中」并加宽左侧，适合打印装订。",
    "三级标题使用楷体加粗，便于与二级黑体区分。",
  ],
};

/** 偏文科常见：行距更松、段落间距更大 */
export const artsThesis: FormatRules = {
  id: "arts-thesis",
  name: "文科论文常用格式",
  school: "文科通用",
  description: "固定值行距约 20 磅等效 1.5 倍；标题层次更疏朗。",
  page: {
    paper: "A4",
    marginTopCm: 2.8,
    marginBottomCm: 2.5,
    marginLeftCm: 2.8,
    marginRightCm: 2.5,
  },
  styles: {
    normal: {
      font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 12 },
      paragraph: {
        alignment: "justify",
        lineSpacing: 1.5,
        firstLineIndentPt: 24,
        spaceAfterPt: 0,
      },
    },
    heading1: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 18, bold: true },
      paragraph: {
        alignment: "center",
        lineSpacing: 1.5,
        spaceBeforePt: 24,
        spaceAfterPt: 24,
        firstLineIndentPt: 0,
      },
    },
    heading2: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 15, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 18,
        spaceAfterPt: 12,
        firstLineIndentPt: 0,
      },
    },
    heading3: {
      font: { name: "黑体", nameAscii: "Arial", sizePt: 12, bold: true },
      paragraph: {
        alignment: "left",
        lineSpacing: 1.5,
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        firstLineIndentPt: 0,
      },
    },
  },
  headerFooter: {
    pageNumber: {
      enabled: true,
      alignment: "center",
      format: "arabic",
    },
  },
  caption: {
    font: { name: "宋体", nameAscii: "Times New Roman", sizePt: 10.5 },
    paragraph: {
      alignment: "center",
      lineSpacing: 1.5,
      spaceBeforePt: 6,
      spaceAfterPt: 12,
      firstLineIndentPt: 0,
    },
  },
  notes: ["一级标题略大（18pt），适合章节标题独立成行的文科写法。"],
};

export const BUILTIN_TEMPLATES: FormatRules[] = [
  genericCnThesis,
  engineeringThesis,
  artsThesis,
];

export function getTemplateById(id: string): FormatRules | undefined {
  return BUILTIN_TEMPLATES.find((t) => t.id === id);
}
