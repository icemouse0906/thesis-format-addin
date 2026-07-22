/** 安全读取 Office/Word 全局，避免 Vite/ESM 把裸标识符 Word 编译坏 */

export type HostDiagnostics = {
  hasOffice: boolean;
  hasWord: boolean;
  hasWordRun: boolean;
  host: string | null;
  hostName: string | null;
  platform: string | null;
  officeReady: boolean;
  hasHostInfoHash: boolean;
  wordApiSupported: boolean | null;
  detail: string;
};

type OfficeGlobals = {
  Office?: typeof Office;
  Word?: typeof Word;
  __officeInfo?: { host?: string; platform?: string };
  __thesisOfficeReady?: boolean;
};

function g(): OfficeGlobals {
  return globalThis as unknown as OfficeGlobals;
}

function hasHostInfoInUrl(): boolean {
  try {
    const blob = `${window.location.hash}${window.location.search}`;
    return /_host_Info/i.test(blob);
  } catch {
    return false;
  }
}

export function getWordApi(): typeof Word {
  const WordApi = g().Word;
  if (!WordApi || typeof WordApi.run !== "function") {
    const d = getHostDiagnostics();
    throw new Error(
      `Word API 不可用（${d.detail}）。请：①关闭任务窗格 ②⌘Q 退出 Word ③重新打开文档并点「格式助手」。不要用浏览器打开本页。`
    );
  }
  return WordApi;
}

export function getOfficeApi(): typeof Office | undefined {
  return g().Office;
}

export function getHostDiagnostics(): HostDiagnostics {
  const env = g();
  const OfficeApi = env.Office;
  let host: string | null = null;
  let hostName: string | null = null;
  let platform: string | null = null;
  let wordApiSupported: boolean | null = null;

  try {
    host = (OfficeApi?.context?.host as unknown as string) ?? null;
  } catch {
    host = null;
  }
  try {
    hostName =
      (OfficeApi?.context as unknown as { diagnostics?: { hostName?: string } })
        ?.diagnostics?.hostName ?? null;
  } catch {
    hostName = null;
  }
  try {
    platform = (OfficeApi?.context?.platform as unknown as string) ?? null;
  } catch {
    platform = null;
  }
  try {
    wordApiSupported =
      OfficeApi?.context?.requirements?.isSetSupported("WordApi", "1.1") ?? null;
  } catch {
    wordApiSupported = null;
  }

  if (!host && env.__officeInfo?.host) {
    host = String(env.__officeInfo.host);
  }
  if (!platform && env.__officeInfo?.platform) {
    platform = String(env.__officeInfo.platform);
  }

  const hasOffice = typeof OfficeApi !== "undefined" && !!OfficeApi;
  const hasWord = !!env.Word;
  const hasWordRun = typeof env.Word?.run === "function";
  const officeReady = !!env.__thesisOfficeReady || !!env.__officeInfo;
  const hasHostInfoHash = hasHostInfoInUrl();

  const connected =
    hasWordRun ||
    String(host).toLowerCase() === "word" ||
    String(hostName).toLowerCase() === "word";

  const detail = [
    `Office=${hasOffice ? "Y" : "N"}`,
    `Word.run=${hasWordRun ? "Y" : "N"}`,
    `host=${host ?? "-"}`,
    `req=${wordApiSupported == null ? "-" : wordApiSupported ? "Y" : "N"}`,
    `hash=${hasHostInfoHash ? "Y" : "N"}`,
    `platform=${platform ?? "-"}`,
  ].join(" | ");

  return {
    hasOffice,
    hasWord,
    hasWordRun,
    host,
    hostName,
    platform,
    officeReady,
    hasHostInfoHash,
    wordApiSupported,
    detail: connected ? `已连接 · ${detail}` : `未连接 · ${detail}`,
  };
}

export function isOfficeWordHost(): boolean {
  try {
    const d = getHostDiagnostics();
    return (
      d.hasWordRun ||
      String(d.host).toLowerCase() === "word" ||
      String(d.hostName).toLowerCase() === "word"
    );
  } catch {
    return false;
  }
}

export function markOfficeReady(info?: { host?: string; platform?: string }) {
  const env = g();
  env.__thesisOfficeReady = true;
  if (info) env.__officeInfo = info;
}

/** 等待 Word.run 出现（宿主脚本偶发晚于 onReady） */
export function waitForWordApi(timeoutMs = 8000): Promise<typeof Word> {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const tick = () => {
      const WordApi = g().Word;
      if (WordApi && typeof WordApi.run === "function") {
        resolve(WordApi);
        return;
      }
      if (Date.now() - start >= timeoutMs) {
        reject(new Error(getHostDiagnostics().detail));
        return;
      }
      window.setTimeout(tick, 200);
    };
    tick();
  });
}
