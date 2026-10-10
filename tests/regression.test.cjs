const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const ts = require('typescript');
// 直接执行项目 TypeScript 模块，而非复制一份实现做“冒烟”。
require.extensions['.ts'] = (module, filename) => {
  const source = ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020 } }).outputText;
  module._compile(source, filename);
};
const { parseRequirementsToRules } = require('../src/shared/parseRequirements.ts');
const { BUILTIN_TEMPLATES } = require('../src/shared/templates.ts');
const { validateRules } = require('../src/shared/validateRules.ts');
const { importRulesJson, exportRulesJson } = require('../src/shared/rulesIO.ts');
const { parseRequirementsWithDeepSeek } = require('../src/shared/deepseekParse.ts');
const { applyFormatToDocument } = require('../src/shared/applyFormat.ts');
const settings = { apiKey: 'test-only', baseUrl: '/llm', model: 'deepseek-chat', enabled: true };
const clone = () => JSON.parse(JSON.stringify(BUILTIN_TEMPLATES[0]));
test('内置模板全部通过校验，导入补齐默认值后稳定往返', () => {
  for (const rule of BUILTIN_TEMPLATES) {
    validateRules(rule);
    const imported = importRulesJson(exportRulesJson(rule));
    assert.equal(imported.id, rule.id);
    assert.equal(imported.styles.normal.font.sizePt, rule.styles.normal.font.sizePt);
    for (const [key, value] of Object.entries(rule.page)) assert.equal(imported.page[key], value);
    assert.deepEqual(importRulesJson(exportRulesJson(imported)), imported);
  }
});
test('执行真实本地解析器，识别小五、正文与标题字号', () => {
  const result = parseRequirementsToRules('正文：宋体小五，1.5倍行距\n一级标题：黑体三号，居中');
  assert.equal(result.rules.styles.normal.font.sizePt, 9);
  assert.equal(result.rules.styles.heading1.font.sizePt, 16);
  assert.equal(result.rules.styles.heading1.paragraph.alignment, 'center');
  validateRules(result.rules);
});
test('非法导入在进入 Word 之前被拒绝', async () => {
  for (const mutate of [r => r.styles.normal.font.sizePt = -2, r => r.styles.normal.paragraph.alignment = 'banana', r => r.page.marginLeftCm = 22, r => r.notes = [{}], r => r.name = {}, r => r.styles.normal.font.bold = 'false']) {
    const rule = clone(); mutate(rule);
    assert.throws(() => importRulesJson(JSON.stringify(rule)));
    await assert.rejects(applyFormatToDocument(rule, {}));
  }
});
test('AI 只允许本地代理或 HTTPS，不向不安全端点发送密钥', async () => {
  const previous = global.fetch; let calls = 0;
  global.fetch = async () => { calls++; throw new Error('should not call'); };
  try {
    for (const baseUrl of ['http://example.com', '//example.com', 'https://user:pass@example.com', 'https://example.com?key=abc']) {
      await assert.rejects(parseRequirementsWithDeepSeek('正文小四', { ...settings, baseUrl }), /Base URL/);
    }
    assert.equal(calls, 0);
  } finally { global.fetch = previous; }
});
test('AI 正常响应可生成规则，异常字号和 HTTP 响应不会直接套用或泄露响应正文', async () => {
  const previous = global.fetch;
  try {
    global.fetch = async (url, options) => {
      assert.equal(url, '/llm/chat/completions'); assert.equal(options.redirect, 'error');
      return { ok: true, json: async () => ({ choices: [{ message: { content: JSON.stringify({ name: 'test', styles: { normal: { font: { sizePt: 12 } } } }) } }] }) };
    };
    assert.equal((await parseRequirementsWithDeepSeek('正文小四', settings)).rules.styles.normal.font.sizePt, 12);
    global.fetch = async () => ({ ok: true, json: async () => ({ choices: [{ message: { content: '{"styles":{"normal":{"font":{"sizePt":-10}}}}' } }] }) });
    await assert.rejects(parseRequirementsWithDeepSeek('正文小四', settings), /sizePt/);
    global.fetch = async () => ({ ok: false, status: 401, text: async () => 'sensitive-provider-response' });
    await assert.rejects(parseRequirementsWithDeepSeek('正文小四', settings), error => error.message === 'DeepSeek 请求失败 HTTP 401');
  } finally { global.fetch = previous; }
});
test('AI 超时会取消请求并释放计时器；调用方也可取消', async () => {
  const original = { fetch: global.fetch, setTimeout: global.setTimeout, clearTimeout: global.clearTimeout };
  let timeout, cleared = 0;
  global.setTimeout = fn => { timeout = fn; return 1; };
  global.clearTimeout = () => { cleared++; };
  global.fetch = async (_, { signal }) => new Promise((resolve, reject) => {
    if (signal.aborted) reject(new Error('aborted'));
    signal.addEventListener('abort', () => reject(new Error('aborted')), { once: true });
  });
  try {
    const pending = parseRequirementsWithDeepSeek('正文小四', settings); timeout();
    await assert.rejects(pending, /超时/); assert.equal(cleared, 1);
    const controller = new AbortController(); const second = parseRequirementsWithDeepSeek('正文小四', settings, controller.signal); controller.abort();
    await assert.rejects(second, /aborted/); assert.equal(cleared, 2);
  } finally { Object.assign(global, original); }
});
