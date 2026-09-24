# ThesisFormat Add-in

> **Word 论文格式助手** — 把杂乱的学校排版要求，变成可执行的结构化规则，再一键套用到已写好的 Word 文档。

[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Office.js](https://img.shields.io/badge/Office.js-Word_Add--in-orange)](https://learn.microsoft.com/office/dev/add-ins/)
[![Vite](https://img.shields.io/badge/Vite-6-646CFF?logo=vite&logoColor=white)](https://vitejs.dev/)
[![DeepSeek](https://img.shields.io/badge/LLM-DeepSeek_API-411D31)](https://www.deepseek.com/)
[![Platform](https://img.shields.io/badge/Platform-macOS_Word-000000?logo=apple)](https://www.microsoft.com/microsoft-365/word)

面向「高校论文格式繁琐」痛点：学生已经写完正文，却卡在宋体小四、1.5 倍行距、标题层级、页边距等校级规范上。本项目用 **Office Add-in（任务窗格）** 把「读懂要求」和「改文档」拆开——AI / 规则引擎只负责解析，Word 侧用确定性 API 执行，结果可预览、可勾选、可撤销。

---

## 为什么不是「让 AI 直接改 Word」？

| 方案 | 问题 |
|------|------|
| 每次让大模型直接改文档 | 不稳定、难复现、难调试 |
| 只手写死每校模板 | 质量高但扩展慢 |
| **LLM / 规则 → FormatRules Schema → Office.js 执行** | 可解释、可版本化、可跨校复用 |

核心设计一句话：

```text
格式手册 / 自然语言
        ↓  DeepSeek 或本地规则解析
   FormatRules（JSON Schema）
        ↓  确定性套用引擎
   Word 文档（段落字体 / 标题 / 题注…）
        ↓
   结构扫描 + 合规评分报告
```

---

## 功能亮点

### 1. 双引擎解析
- **DeepSeek 智能解析**：粘贴培养手册式长文本，自备 API Key，输出结构化规则  
- **本地规则解析**：离线可用，覆盖常见「小四 / 三号 / 1.5 倍行距」表述  

### 2. 固定学校模板
内置通用中文学位论文 / 工科 / 文科模板，适合快速演示与冷启动。

### 3. Mac Word 实战兼容
针对 Mac 上 `getStyles` / `pageSetup` / 页码 OOXML 等能力缺口做了降级：
- 段落直接属性套用（`font` / `alignment` / `space1Pt5()`）
- 默认不自动写页码，避免损坏文档  
- 页边距给出明确手动指引  

### 4. 实验室能力（演示向）
- **文档结构扫描**：H1/H2/H3 / 题注大纲  
- **合规评分**：对照当前规则抽样检查字号、缩进、标题  
- **规则 JSON 导入导出**：模板可分享、可版本管理  

---

## 技术栈

- **前端**：React 18 + TypeScript + Vite  
- **宿主**：Microsoft Word（Office.js Task Pane Add-in）  
- **AI**：DeepSeek Chat Completions（OpenAI 兼容协议，用户自备 Key）  
- **工程**：本地 HTTPS 证书、Vite 代理解决浏览器 CORS、postinstall 同步 office-js  

---

## 快速开始（macOS + Microsoft Word）

### 环境
- Node.js 20+
- 已安装 Microsoft Word for Mac

### 安装与启动

```bash
git clone https://github.com/icemouse0906/thesis-format-addin.git
cd thesis-format-addin
npm install
npm run dev
```

开发服务：`https://localhost:3000/taskpane/index.html`

### 信任本地证书（首次）

```bash
npm run trust-cert
```

### 侧载到 Word

```bash
npm run sideload
```

然后 **完全退出 Word（⌘Q）再打开**，在功能区 **开始** 中找到 **「格式助手」**。

### DeepSeek（可选）

1. 在 [DeepSeek 开放平台](https://platform.deepseek.com/) 创建 API Key  
2. 插件内粘贴并保存（仅存本机 `localStorage`）  
3. 「智能解析」→ **DeepSeek 智能解析** → 确认套用  

开发态默认 Base URL 为 `/llm`（Vite 代理到官方 API，避免 CORS）。

---

## 项目结构

```text
src/
  shared/
    schema.ts            # FormatRules 统一中间层
    templates.ts         # 内置学校/通用模板
    parseRequirements.ts # 本地规则解析
    deepseekParse.ts     # DeepSeek → FormatRules
    applyFormat.ts       # Office.js 套用引擎（Mac 兼容）
    documentIntel.ts     # 结构扫描 / 合规评分
    rulesIO.ts           # 规则 JSON 导入导出
    officeHost.ts        # Word 宿主探测
  taskpane/              # 任务窗格 UI
  commands/              # Ribbon 命令占位
manifest.xml             # Office Add-in 清单
scripts/                 # 侧载 / 证书 / office-js 同步
```

---

## 脚本

| 命令 | 说明 |
|------|------|
| `npm run dev` | HTTPS 开发服（含 `/llm` 代理） |
| `npm run build` | 类型检查 + 生产构建 |
| `npm run sideload` | 复制 manifest 到 Word `wef` 目录 |
| `npm run trust-cert` | 安装/信任本地开发证书 |
| `npm run smoke` | 本地冒烟检查 |

---



---

## 已知限制

- Mac Word 对页边距 / 样式表 / 页码域支持弱于 Windows；页码请手动插入。  
- 标题识别依赖文本启发式（`一、` / `1.1` / `第X章` 等），极端排版需人工微调。  
- DeepSeek 直连受 CORS 限制，生产环境需自建代理或保持开发代理方案。  

---

## License

MIT
