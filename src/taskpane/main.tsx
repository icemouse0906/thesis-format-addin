import React from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import {
  markOfficeReady,
  waitForWordApi,
  getHostDiagnostics,
} from "../shared/officeHost";
import "./styles.css";

const rootEl = document.getElementById("root");
if (!rootEl) {
  throw new Error("root element missing");
}

const root = createRoot(rootEl);

function render() {
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

function waitForOfficeGlobal(timeoutMs = 15000): Promise<typeof Office> {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const tick = () => {
      const OfficeApi = (globalThis as unknown as { Office?: typeof Office }).Office;
      if (OfficeApi && typeof OfficeApi.onReady === "function") {
        resolve(OfficeApi);
        return;
      }
      if (Date.now() - start >= timeoutMs) {
        reject(new Error("Office.js 未加载"));
        return;
      }
      window.setTimeout(tick, 50);
    };
    tick();
  });
}

function withTimeout<T>(p: Promise<T>, ms: number, label: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const t = window.setTimeout(() => reject(new Error(`${label} 超时 ${ms}ms`)), ms);
    p.then(
      (v) => {
        window.clearTimeout(t);
        resolve(v);
      },
      (e) => {
        window.clearTimeout(t);
        reject(e);
      }
    );
  });
}

async function boot() {
  try {
    const OfficeApi = await waitForOfficeGlobal();
    console.info("[论文格式助手] Office 全局已出现", getHostDiagnostics());

    try {
      const info = await withTimeout(OfficeApi.onReady(), 12000, "Office.onReady");
      markOfficeReady({
        host: info.host != null ? String(info.host) : undefined,
        platform: info.platform != null ? String(info.platform) : undefined,
      });
      console.info("[论文格式助手] Office ready", info, getHostDiagnostics());
    } catch (e) {
      console.warn("[论文格式助手] onReady 未完成（若在浏览器中属正常）", e);
      markOfficeReady();
    }

    // 先渲染 UI，Word API 在后台等到再刷新诊断
    render();

    try {
      await waitForWordApi(12000);
      console.info("[论文格式助手] Word API ready", getHostDiagnostics());
    } catch (e) {
      console.warn("[论文格式助手] Word API 等待超时", e, getHostDiagnostics());
    }
    return;
  } catch (e) {
    console.error("[论文格式助手] 启动失败", e);
  }

  render();
}

boot();
