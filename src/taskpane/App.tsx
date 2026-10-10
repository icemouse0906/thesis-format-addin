import { useEffect, useMemo, useState } from "react";
import { BUILTIN_TEMPLATES, getTemplateById } from "../shared/templates";
import {
  DEFAULT_APPLY_OPTIONS,
  type ApplyOptions,
  type FormatRules,
} from "../shared/schema";
import { parseRequirementsToRules } from "../shared/parseRequirements";
import { parseRequirementsWithDeepSeek } from "../shared/deepseekParse";
import {
  applyFormatToDocument,
  simulateApply,
  type ApplyReport,
} from "../shared/applyFormat";
import {
  getHostDiagnostics,
  waitForWordApi,
  type HostDiagnostics,
} from "../shared/officeHost";
import {
  loadLlmSettings,
  saveLlmSettings,
  maskApiKey,
  type LlmSettings,
} from "../shared/settings";
import {
  scanDocumentStructure,
  checkCompliance,
  type StructureReport,
  type ComplianceReport,
} from "../shared/documentIntel";
import { downloadRulesJson, importRulesJson, readFileAsText } from "../shared/rulesIO";

type Mode = "template" | "parse" | "ai";

const SAMPLE_REQUIREMENTS = `正文：宋体小四，两端对齐，1.5倍行距，首行缩进2字符
一级标题：黑体三号居中加粗
二级标题：黑体四号加粗
三级标题：黑体小四加粗
页边距：上2.5 下2.5 左3.0 右2.5 cm
纸张A4
页码居中`;

const OPTION_GROUPS: Array<{
  title: string;
  items: Array<[keyof ApplyOptions, string, string?]>;
}> = [
  {
    title: "范围",
    items: [["selectionOnly", "仅当前选区", "勾选后只改选中文字，不改全文"]],
  },
  {
    title: "页面",
    items: [
      ["applyPage", "页边距"],
      ["applyPaperSize", "纸张 A4"],
      [
        "applyHeaderFooter",
        "页眉页脚/页码",
        "Mac 上默认关闭，避免损坏文档；页码请手动插入",
      ],
    ],
  },
  {
    title: "正文与标题",
    items: [
      ["applyNormal", "正文样式（宋体/行距/缩进）"],
      ["applyHeadings", "标题 1–3 样式"],
      ["autoDetectHeadings", "自动识别标题（第X章 / 1.1 / 一、）"],
      ["pageBreakBeforeH1", "一级标题前分页"],
    ],
  },
  {
    title: "常用增强",
    items: [
      ["applyCaptions", "图题 / 表题格式"],
      ["clearDirectFormatting", "清除直接格式干扰"],
      ["removeExtraBlankLines", "压缩多余空行"],
    ],
  },
];

export function App() {
  const [mode, setMode] = useState<Mode>("template");
  const [templateId, setTemplateId] = useState(BUILTIN_TEMPLATES[0].id);
  const [rawText, setRawText] = useState(SAMPLE_REQUIREMENTS);
  const [options, setOptions] = useState<ApplyOptions>({ ...DEFAULT_APPLY_OPTIONS });
  const [busy, setBusy] = useState(false);
  const [aiBusy, setAiBusy] = useState(false);
  const [report, setReport] = useState<ApplyReport | null>(null);
  const [parseNotes, setParseNotes] = useState<string[]>([]);
  const [diag, setDiag] = useState<HostDiagnostics>(() => getHostDiagnostics());
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [llm, setLlm] = useState<LlmSettings>(() => loadLlmSettings());
  const [aiRules, setAiRules] = useState<FormatRules | null>(null);
  const [draftKey, setDraftKey] = useState(() => loadLlmSettings().apiKey);
  const [intelBusy, setIntelBusy] = useState(false);
  const [structure, setStructure] = useState<StructureReport | null>(null);
  const [compliance, setCompliance] = useState<ComplianceReport | null>(null);
  const [importedRules, setImportedRules] = useState<FormatRules | null>(null);

  useEffect(() => {
    const refresh = () => setDiag(getHostDiagnostics());
    refresh();
    const t = window.setInterval(refresh, 1000);
    return () => window.clearInterval(t);
  }, []);

  const inWord = diag.hasWordRun;
  const officeLoaded = diag.hasOffice;

  const localParsed = useMemo(
    () => parseRequirementsToRules(rawText).rules,
    [rawText]
  );

  const activeRules: FormatRules = useMemo(() => {
    if (importedRules) return importedRules;
    if (mode === "template") {
      return getTemplateById(templateId) ?? BUILTIN_TEMPLATES[0];
    }
    if (mode === "ai" && aiRules) return aiRules;
    return localParsed;
  }, [mode, templateId, aiRules, localParsed, importedRules]);

  function onLocalParse() {
    const result = parseRequirementsToRules(rawText);
    setAiRules(null);
    setParseNotes([
      "引擎：本地规则",
      ...result.matched.map((m) => `✓ ${m}`),
      ...result.unmatched.slice(0, 8).map((u) => `? 未归类：${u}`),
    ]);
    setReport(null);
  }

  async function onDeepSeekParse() {
    setAiBusy(true);
    setReport(null);
    try {
      const settings = loadLlmSettings();
      const result = await parseRequirementsWithDeepSeek(rawText, settings);
      setAiRules(result.rules);
      setMode("ai");
      setParseNotes(["引擎：DeepSeek", ...result.matched.map((m) => `✓ ${m}`)]);
    } catch (e) {
      setParseNotes([
        "引擎：DeepSeek 失败",
        e instanceof Error ? e.message : String(e),
        "可改用「本地规则解析」，或检查 API Key / npm run dev 代理。",
      ]);
    } finally {
      setAiBusy(false);
    }
  }

  function saveKey() {
    const next = saveLlmSettings({ apiKey: draftKey.trim(), enabled: true });
    setLlm(next);
    setParseNotes([`已保存 API Key（${maskApiKey(next.apiKey)}）；调用 AI 时发送到所配置的服务。`]);
  }

  async function runApply(real: boolean) {
    setBusy(true);
    setReport(null);
    setConfirmOpen(false);
    try {
      const rules = activeRules;

      if (mode === "parse" && !aiRules) {
        const parsed = parseRequirementsToRules(rawText);
        setParseNotes([
          "引擎：本地规则",
          ...parsed.matched.map((m) => `✓ ${m}`),
        ]);
      }

      if (!real) {
        setReport(simulateApply(rules, options));
        return;
      }

      try {
        await waitForWordApi(6000);
      } catch {
        /* ignore */
      }

      const result = await applyFormatToDocument(rules, options);
      setReport(result);
      setDiag(getHostDiagnostics());
    } catch (e) {
      setReport({
        ok: false,
        steps: [],
        warnings: [],
        error: e instanceof Error ? e.message : String(e),
      });
      setDiag(getHostDiagnostics());
    } finally {
      setBusy(false);
    }
  }

  function toggleOption<K extends keyof ApplyOptions>(key: K) {
    setOptions((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  async function onScanStructure() {
    setIntelBusy(true);
    try {
      await waitForWordApi(4000).catch(() => undefined);
      const report = await scanDocumentStructure();
      setStructure(report);
    } finally {
      setIntelBusy(false);
    }
  }

  async function onCheckCompliance() {
    setIntelBusy(true);
    try {
      await waitForWordApi(4000).catch(() => undefined);
      const report = await checkCompliance(activeRules);
      setCompliance(report);
    } finally {
      setIntelBusy(false);
    }
  }

  async function onImportRulesFile(file: File | null) {
    if (!file) return;
    try {
      const text = await readFileAsText(file);
      const rules = importRulesJson(text);
      setImportedRules(rules);
      setAiRules(null);
      setParseNotes([`已导入规则：${rules.name}`]);
    } catch (e) {
      setParseNotes([
        "导入失败",
        e instanceof Error ? e.message : String(e),
      ]);
    }
  }

  return (
    <div className="app">
      <header className="header">
        <div className="brand">论文格式助手</div>
        <div className={`badge ${inWord ? "ok" : officeLoaded ? "warn" : "warn"}`}>
          {inWord ? "Word 已连接" : officeLoaded ? "Office 已加载" : "未加载 Office.js"}
        </div>
      </header>

      <p className="diag">{diag.detail}</p>

      <nav className="tabs">
        <button
          className={mode === "template" ? "tab active" : "tab"}
          onClick={() => setMode("template")}
          type="button"
        >
          模板
        </button>
        <button
          className={mode === "parse" || mode === "ai" ? "tab active" : "tab"}
          onClick={() => setMode(aiRules ? "ai" : "parse")}
          type="button"
        >
          智能解析
        </button>
      </nav>

      {/* 始终可访问的 AI 设置条 */}
      <section className="panel">
        <div className="label">DeepSeek API（自备 Key）</div>
        <p className="desc">
          Key 保存在本机 localStorage；AI 解析时，Key 和输入的格式要求会发送到配置的服务。
          默认通过本地代理 <code>/llm</code> 请求 DeepSeek，费用以服务商账单为准。
        </p>
        <input
          className="select"
          type="password"
          placeholder="粘贴 sk- 开头的 DeepSeek API Key"
          value={draftKey}
          onChange={(e) => setDraftKey(e.target.value)}
          autoComplete="off"
        />
        <div className="row-actions">
          <button className="btn ghost" type="button" onClick={saveKey}>
            保存 Key（当前 {maskApiKey(llm.apiKey)}）
          </button>
        </div>
        <label className="label" htmlFor="base">
          Base URL
        </label>
        <input
          id="base"
          className="select"
          value={llm.baseUrl}
          onChange={(e) => {
            const next = saveLlmSettings({ baseUrl: e.target.value });
            setLlm(next);
          }}
        />
        <p className="desc muted-sm">
          默认 <code>/llm</code>（推荐）。若你有自己的中转，可改成
          https://api.deepseek.com 或可信的 HTTPS 中转地址。中转方会收到 Key 和格式要求。
        </p>
      </section>

      {mode === "template" ? (
        <section className="panel">
          <label className="label" htmlFor="tpl">
            选择模板
          </label>
          <select
            id="tpl"
            className="select"
            value={templateId}
            onChange={(e) => setTemplateId(e.target.value)}
          >
            {BUILTIN_TEMPLATES.map((t) => (
              <option key={t.id} value={t.id}>
                {t.name}（{t.school}）
              </option>
            ))}
          </select>
          <p className="desc">{activeRules.description}</p>
        </section>
      ) : (
        <section className="panel">
          <label className="label" htmlFor="req">
            粘贴学校格式要求（可很长、很乱）
          </label>
          <textarea
            id="req"
            className="textarea"
            value={rawText}
            onChange={(e) => {
              setRawText(e.target.value);
              setAiRules(null);
            }}
            rows={8}
            placeholder="粘贴培养手册/格式说明全文…"
          />
          <div className="row-actions">
            <button
              className="btn primary"
              type="button"
              disabled={aiBusy || busy}
              onClick={onDeepSeekParse}
            >
              {aiBusy ? "DeepSeek 解析中…" : "DeepSeek 智能解析"}
            </button>
            <button
              className="btn ghost"
              type="button"
              disabled={aiBusy || busy}
              onClick={onLocalParse}
            >
              本地规则解析
            </button>
          </div>
          {aiRules && (
            <p className="desc">
              当前将套用：<strong>{aiRules.name}</strong>（DeepSeek）
            </p>
          )}
          {parseNotes.length > 0 && (
            <ul className="notes">
              {parseNotes.map((n) => (
                <li key={n}>{n}</li>
              ))}
            </ul>
          )}
        </section>
      )}

      {OPTION_GROUPS.map((group) => (
        <section className="panel" key={group.title}>
          <div className="label">{group.title}</div>
          <div className="checks">
            {group.items.map(([key, label, hint]) => (
              <label key={key} className="check" title={hint}>
                <input
                  type="checkbox"
                  checked={options[key]}
                  onChange={() => toggleOption(key)}
                />
                <span>
                  {label}
                  {hint ? <span className="hint"> · {hint}</span> : null}
                </span>
              </label>
            ))}
          </div>
        </section>
      ))}

      <section className="panel preview">
        <div className="label">将套用的规则</div>
        <ul className="summary">
          <li>
            页边距：上{activeRules.page.marginTopCm} / 下{activeRules.page.marginBottomCm} /
            左{activeRules.page.marginLeftCm} / 右{activeRules.page.marginRightCm} cm
          </li>
          <li>
            正文：{activeRules.styles.normal.font.name}+
            {activeRules.styles.normal.font.nameAscii}{" "}
            {activeRules.styles.normal.font.sizePt}pt，行距{" "}
            {activeRules.styles.normal.paragraph.lineSpacing}
          </li>
          <li>
            标题：
            {activeRules.styles.heading1.font.sizePt}/
            {activeRules.styles.heading2.font.sizePt}/
            {activeRules.styles.heading3.font.sizePt}pt
          </li>
        </ul>
        {importedRules && (
          <p className="desc">
            使用导入规则「{importedRules.name}」{" "}
            <button
              type="button"
              className="linkish"
              onClick={() => setImportedRules(null)}
            >
              清除导入
            </button>
          </p>
        )}
      </section>

      <section className="panel lab">
        <div className="label">实验室 · 结构扫描 / 合规评分 / 规则交换</div>
        <p className="desc">
          演示向能力：先看文档骨架，再打合规分；规则可导出 JSON
          分享给同学。
        </p>
        <div className="stats">
          <div className="stat">
            <div className="stat-num">{structure?.paragraphCount ?? "—"}</div>
            <div className="stat-label">段落</div>
          </div>
          <div className="stat">
            <div className="stat-num">
              {structure
                ? `${structure.heading1}/${structure.heading2}/${structure.heading3}`
                : "—"}
            </div>
            <div className="stat-label">H1/H2/H3</div>
          </div>
          <div className="stat">
            <div className="stat-num">
              {compliance?.ok ? `${compliance.score}` : "—"}
            </div>
            <div className="stat-label">合规分</div>
          </div>
        </div>
        <div className="row-actions">
          <button
            className="btn ghost"
            type="button"
            disabled={intelBusy || busy}
            onClick={onScanStructure}
          >
            {intelBusy ? "分析中…" : "扫描文档结构"}
          </button>
          <button
            className="btn ghost"
            type="button"
            disabled={intelBusy || busy}
            onClick={onCheckCompliance}
          >
            对照规则做合规检查
          </button>
          <button
            className="btn ghost"
            type="button"
            disabled={busy}
            onClick={() => downloadRulesJson(activeRules)}
          >
            导出当前规则 JSON
          </button>
          <label className="btn ghost file-btn">
            导入规则 JSON
            <input
              type="file"
              accept="application/json,.json"
              hidden
              onChange={(e) => onImportRulesFile(e.target.files?.[0] ?? null)}
            />
          </label>
        </div>
        {structure && (
          <div className="intel-block">
            {structure.error ? (
              <p className="error">{structure.error}</p>
            ) : (
              <>
                <div className="label">结构大纲（截取）</div>
                <ul className="outline">
                  {structure.outline.length === 0 && <li>未识别到标题/题注</li>}
                  {structure.outline.map((o) => (
                    <li key={`${o.index}-${o.text}`} className={`lv-${o.level}`}>
                      <span className="tag">
                        {o.level === "caption" ? "题注" : `H${o.level}`}
                      </span>
                      {o.text}
                    </li>
                  ))}
                </ul>
              </>
            )}
          </div>
        )}
        {compliance && (
          <div className="intel-block">
            {compliance.error ? (
              <p className="error">{compliance.error}</p>
            ) : (
              <>
                <div className="label">合规报告 · {compliance.score} 分</div>
                <ul className="compliance">
                  {compliance.items.map((it) => (
                    <li key={it.id} className={`st-${it.status}`}>
                      <strong>
                        {it.status === "pass"
                          ? "通过"
                          : it.status === "warn"
                            ? "存疑"
                            : it.status === "fail"
                              ? "未达标"
                              : "跳过"}
                      </strong>
                      {" · "}
                      {it.label}
                      <div className="hint">{it.detail}</div>
                    </li>
                  ))}
                </ul>
              </>
            )}
          </div>
        )}
      </section>

      <div className="actions">
        <button
          className="btn primary"
          type="button"
          disabled={busy || aiBusy}
          onClick={() => setConfirmOpen(true)}
        >
          {busy ? "正在套用…" : "确认套用到文档"}
        </button>
        <button
          className="btn ghost"
          type="button"
          disabled={busy || aiBusy}
          onClick={() => runApply(false)}
        >
          仅预览将执行的操作
        </button>
      </div>

      {confirmOpen && (
        <div className="modal-backdrop" role="presentation">
          <div className="modal" role="dialog" aria-modal="true">
            <div className="label">确认写入当前 Word 文档？</div>
            <p className="desc">
              规则来源：{mode === "template" ? "固定模板" : aiRules ? "DeepSeek" : "本地解析"}
              。可用 ⌘Z 撤销。
            </p>
            <div className="modal-actions">
              <button
                className="btn ghost"
                type="button"
                disabled={busy}
                onClick={() => setConfirmOpen(false)}
              >
                取消
              </button>
              <button
                className="btn primary"
                type="button"
                disabled={busy}
                onClick={() => runApply(true)}
              >
                {busy ? "执行中…" : "确认执行"}
              </button>
            </div>
          </div>
        </div>
      )}

      {report && (
        <section className={`panel report ${report.ok ? "ok" : "err"}`}>
          <div className="label">{report.ok ? "执行结果" : "执行失败"}</div>
          {report.error && <p className="error">{report.error}</p>}
          <ul className="notes">
            {report.steps.map((s) => (
              <li key={s}>{s}</li>
            ))}
          </ul>
          {report.warnings.length > 0 && (
            <>
              <div className="label">注意</div>
              <ul className="notes warn">
                {report.warnings.map((w) => (
                  <li key={w}>{w}</li>
                ))}
              </ul>
            </>
          )}
        </section>
      )}

      <footer className="footer">
        技术路径：自然语言/手册 →（DeepSeek）→ FormatRules →
        Word 确定性套用。页码请手动插入。
      </footer>
    </div>
  );
}
