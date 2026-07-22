import type { FormatRules, StyleRule } from "./schema";
import { genericCnThesis } from "./templates";
import type { LlmSettings } from "./settings";
import type { ParseResult } from "./parseRequirements";

function deepClone<T>(v: T): T {
  return JSON.parse(JSON.stringify(v)) as T;
}

const SYSTEM_PROMPT = `你是中文学位论文/实验报告排版规则解析器。
用户会粘贴学校格式要求（可能很乱、很长）。你的任务是抽取排版规则，输出【唯一一个 JSON 对象】，不要 Markdown，不要解释。

JSON 必须符合下列 TypeScript 形状（缺省字段用 null，不要发明无关键）：
{
  "name": string,
  "school": string | null,
  "description": string | null,
  "page": {
    "paper": "A4" | "Letter" | null,
    "marginTopCm": number | null,
    "marginBottomCm": number | null,
    "marginLeftCm": number | null,
    "marginRightCm": number | null
  },
  "styles": {
    "normal": Style,
    "heading1": Style,
    "heading2": Style,
    "heading3": Style
  },
  "caption": Style | null,
  "notes": string[]
}
Style = {
  "font": { "name": string|null, "nameAscii": string|null, "sizePt": number|null, "bold": boolean|null },
  "paragraph": {
    "alignment": "left"|"center"|"right"|"justify"|null,
    "lineSpacing": number|null,
    "firstLineIndentPt": number|null,
    "spaceBeforePt": number|null,
    "spaceAfterPt": number|null
  }
}

字号换算（必须遵守）：
小五=9, 五号=10.5, 小四=12, 四号=14, 小三=15, 三号=16, 小二=18, 二号=22。
首行缩进「两字符/2字符」→ firstLineIndentPt=24（按小四近似）。
行距「1.5倍」→ lineSpacing=1.5。
中文正文默认 name=宋体、nameAscii=Times New Roman；标题常用黑体。
只输出 JSON。`;

type PartialStyle = {
  font?: {
    name?: string | null;
    nameAscii?: string | null;
    sizePt?: number | null;
    bold?: boolean | null;
  };
  paragraph?: {
    alignment?: "left" | "center" | "right" | "justify" | null;
    lineSpacing?: number | null;
    firstLineIndentPt?: number | null;
    spaceBeforePt?: number | null;
    spaceAfterPt?: number | null;
  };
};

type LlmJson = {
  name?: string;
  school?: string | null;
  description?: string | null;
  page?: {
    paper?: "A4" | "Letter" | null;
    marginTopCm?: number | null;
    marginBottomCm?: number | null;
    marginLeftCm?: number | null;
    marginRightCm?: number | null;
  };
  styles?: {
    normal?: PartialStyle;
    heading1?: PartialStyle;
    heading2?: PartialStyle;
    heading3?: PartialStyle;
  };
  caption?: PartialStyle | null;
  notes?: string[];
};

function mergeStyle(base: StyleRule, patch?: PartialStyle | null): StyleRule {
  if (!patch) return deepClone(base);
  const out = deepClone(base);
  if (patch.font) {
    if (patch.font.name) out.font.name = patch.font.name;
    if (patch.font.nameAscii) out.font.nameAscii = patch.font.nameAscii;
    if (typeof patch.font.sizePt === "number") out.font.sizePt = patch.font.sizePt;
    if (typeof patch.font.bold === "boolean") out.font.bold = patch.font.bold;
  }
  if (patch.paragraph) {
    if (patch.paragraph.alignment) out.paragraph.alignment = patch.paragraph.alignment;
    if (typeof patch.paragraph.lineSpacing === "number") {
      out.paragraph.lineSpacing = patch.paragraph.lineSpacing;
    }
    if (typeof patch.paragraph.firstLineIndentPt === "number") {
      out.paragraph.firstLineIndentPt = patch.paragraph.firstLineIndentPt;
    }
    if (typeof patch.paragraph.spaceBeforePt === "number") {
      out.paragraph.spaceBeforePt = patch.paragraph.spaceBeforePt;
    }
    if (typeof patch.paragraph.spaceAfterPt === "number") {
      out.paragraph.spaceAfterPt = patch.paragraph.spaceAfterPt;
    }
  }
  return out;
}

function mergeLlmIntoRules(data: LlmJson): FormatRules {
  const rules = deepClone(genericCnThesis);
  rules.id = `deepseek-${Date.now()}`;
  rules.name = data.name?.trim() || "DeepSeek 解析结果";
  rules.school = data.school?.trim() || "AI 解析";
  rules.description =
    data.description?.trim() ||
    "由 DeepSeek 从格式要求文本解析生成，未提及项保留通用模板默认值。";

  if (data.page) {
    if (data.page.paper) rules.page.paper = data.page.paper;
    if (typeof data.page.marginTopCm === "number") {
      rules.page.marginTopCm = data.page.marginTopCm;
    }
    if (typeof data.page.marginBottomCm === "number") {
      rules.page.marginBottomCm = data.page.marginBottomCm;
    }
    if (typeof data.page.marginLeftCm === "number") {
      rules.page.marginLeftCm = data.page.marginLeftCm;
    }
    if (typeof data.page.marginRightCm === "number") {
      rules.page.marginRightCm = data.page.marginRightCm;
    }
  }

  if (data.styles) {
    rules.styles.normal = mergeStyle(rules.styles.normal, data.styles.normal);
    rules.styles.heading1 = mergeStyle(rules.styles.heading1, data.styles.heading1);
    rules.styles.heading2 = mergeStyle(rules.styles.heading2, data.styles.heading2);
    rules.styles.heading3 = mergeStyle(rules.styles.heading3, data.styles.heading3);
  }

  if (data.caption && rules.caption) {
    rules.caption = mergeStyle(rules.caption, data.caption);
  }

  const notes = Array.isArray(data.notes) ? data.notes.filter(Boolean) : [];
  rules.notes = [
    "来源：DeepSeek",
    ...notes.slice(0, 10),
    `正文 ${rules.styles.normal.font.name} ${rules.styles.normal.font.sizePt}pt / 行距 ${rules.styles.normal.paragraph.lineSpacing}`,
  ];
  return rules;
}

function extractJson(text: string): unknown {
  const trimmed = text.trim();
  const fence = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const raw = fence ? fence[1].trim() : trimmed;
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start < 0 || end <= start) {
    throw new Error("模型未返回 JSON 对象");
  }
  return JSON.parse(raw.slice(start, end + 1));
}

function chatUrl(baseUrl: string): string {
  const b = baseUrl.replace(/\/$/, "");
  if (b.endsWith("/chat/completions")) return b;
  return `${b}/chat/completions`;
}

/**
 * 调用 DeepSeek（OpenAI 兼容）把格式要求解析为 FormatRules。
 * 开发环境 baseUrl 默认 /llm，由 Vite 代理到 api.deepseek.com，避免浏览器 CORS。
 */
export async function parseRequirementsWithDeepSeek(
  input: string,
  settings: LlmSettings,
  signal?: AbortSignal
): Promise<ParseResult> {
  const text = input.trim();
  if (!text) {
    return {
      rules: deepClone(genericCnThesis),
      matched: [],
      unmatched: ["输入为空"],
    };
  }
  if (!settings.apiKey.trim()) {
    throw new Error("未配置 API Key。请在「AI 设置」中粘贴 DeepSeek API Key。");
  }

  const res = await fetch(chatUrl(settings.baseUrl), {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${settings.apiKey.trim()}`,
    },
    body: JSON.stringify({
      model: settings.model || "deepseek-chat",
      temperature: 0.1,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        {
          role: "user",
          content: `请解析下列论文/报告格式要求：\n\n${text.slice(0, 12000)}`,
        },
      ],
    }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(
      `DeepSeek 请求失败 HTTP ${res.status}${body ? `：${body.slice(0, 200)}` : ""}`
    );
  }

  const payload = (await res.json()) as {
    choices?: Array<{ message?: { content?: string } }>;
  };
  const content = payload.choices?.[0]?.message?.content;
  if (!content) throw new Error("DeepSeek 返回空内容");

  const data = extractJson(content) as LlmJson;
  const rules = mergeLlmIntoRules(data);

  const matched = [
    `AI 模型：${settings.model}`,
    `模板名：${rules.name}`,
    `正文：${rules.styles.normal.font.name} ${rules.styles.normal.font.sizePt}pt`,
    `标题1：${rules.styles.heading1.font.name} ${rules.styles.heading1.font.sizePt}pt`,
    `页边距：${rules.page.marginTopCm}/${rules.page.marginBottomCm}/${rules.page.marginLeftCm}/${rules.page.marginRightCm} cm`,
  ];

  return { rules, matched, unmatched: [] };
}
