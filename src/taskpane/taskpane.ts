/*
 * Excel 智能助手 Taskpane entry
 * - Mount React
 * - Provide a stable API to React via window.ExcelAssistant
 * - Probe backend health and expose connection status
 * - Use same-origin endpoints (/api, /chat, ...) so webpack proxy can forward to :3001
 */

import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { ErrorBoundary } from "../components/ErrorBoundary";
import { GlobalErrorHandler } from "../utils/GlobalErrorHandler";
import { Logger } from "../utils/Logger";

// 瀵煎叆鍏ㄥ眬绫诲瀷瀹氫箟
import type { BackendHealth, AssistantStatus, AuditLogEntry } from "../types/global";

// 初始化全局错误处理
GlobalErrorHandler.initialize();

const STATUS: AssistantStatus = {
  backend: {
    ok: false,
  },
};

function nowIso() {
  return new Date().toISOString();
}

function showOutput(data: any) {
  const outputEl = document.getElementById("output");
  try {
    if (outputEl) outputEl.textContent = JSON.stringify(data, null, 2);
    console.log("output:", data);
  } catch {
    if (outputEl) outputEl.textContent = String(data);
    console.error("showOutput error");
  }
}

function auditLog(action: string, payload: any) {
  try {
    const key = "excelAssistantAudit";
    const raw = localStorage.getItem(key);
    const arr = raw ? (JSON.parse(raw) as AuditLogEntry[]) : [];
    const entry: AuditLogEntry = { time: nowIso(), action, payload };
    arr.push(entry);
    localStorage.setItem(key, JSON.stringify(arr));
  } catch {
    console.error("audit log failed");
  }
}

function auditList(): AuditLogEntry[] {
  try {
    const raw = localStorage.getItem("excelAssistantAudit");
    return raw ? (JSON.parse(raw) as AuditLogEntry[]) : [];
  } catch {
    return [];
  }
}

function auditClear() {
  try {
    localStorage.removeItem("excelAssistantAudit");
  } catch {
    // ignore
  }
}

/**
 * 鉁?Backend calls: always same-origin
 * With webpack devServer proxy:
 *  - /api/* , /chat , /health will be forwarded to http://localhost:3001
 */
async function httpJson<T>(path: string, init?: RequestInit): Promise<T> {
  const resp = await fetch(path, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      ...(init?.headers || {}),
    },
  });

  const text = await resp.text();
  let data: any = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = { raw: text };
  }

  if (!resp.ok) {
    const msg =
      data?.error || data?.message || `HTTP ${resp.status} ${resp.statusText} for ${path}`;
    throw new Error(msg);
  }

  return data as T;
}

async function checkBackend(): Promise<AssistantStatus["backend"]> {
  try {
    const health = await httpJson<BackendHealth>("/api/health", {
      method: "GET",
    });
    STATUS.backend = {
      ok: true,
      lastCheckedAt: nowIso(),
      health,
    };
    return STATUS.backend;
  } catch (e: any) {
    STATUS.backend = {
      ok: false,
      lastCheckedAt: nowIso(),
      error: String(e?.message || e),
    };
    return STATUS.backend;
  }
}

async function chat(message: string, context: any = {}) {
  const payload = { message, context };
  const res = await httpJson<any>("/chat", {
    method: "POST",
    body: JSON.stringify(payload),
  });
  auditLog("chat", { message, response: res });
  return res;
}

async function perceiveExcel() {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      const selection = context.workbook.getSelectedRange();
      selection.load(["address", "values", "rowCount", "columnCount"]);

      const used = sheet.getUsedRange();
      used.load(["address", "rowCount", "columnCount"]);

      await context.sync();

      const result = {
        sheet: sheet.name,
        selection: {
          address: selection.address,
          rows: selection.rowCount,
          cols: selection.columnCount,
          valuesSample: selection.values ? selection.values.slice(0, 10) : [],
        },
        usedRange: {
          address: used.address,
          rows: used.rowCount,
          cols: used.columnCount,
        },
      };

      auditLog("perceive", result);
      return result;
    });
  } catch (error) {
    console.error("perceiveExcel error", error);
    return { error: String(error) };
  }
}

/**
 * Expose a stable API surface to the React app.
 * App.tsx can do:
 *   const api = window.ExcelAssistant;
 *   await api.checkBackend();
 *   const res = await api.chat("甯垜鎶婇€夊尯鍋氭垚琛ㄦ牸");
 */
function exposeApi() {
  window.ExcelAssistant = {
    status: STATUS,
    checkBackend,
    chat,
    perceiveExcel,
    audit: {
      log: auditLog,
      list: auditList,
      clear: auditClear,
    },
  };
}

function mountReact() {
  const appBody = document.getElementById("app-body");
  if (!appBody) return;

  appBody.style.display = "flex";
  appBody.innerHTML = "";

  const rootEl = document.createElement("div");
  rootEl.id = "react-root";
  rootEl.style.width = "100%";
  rootEl.style.height = "100%";
  appBody.appendChild(rootEl);

  const root = createRoot(rootEl);
  // 使用 ErrorBoundary 包裹 App 组件
  root.render(React.createElement(ErrorBoundary, { moduleName: "App" }, React.createElement(App)));

  Logger.info("Taskpane", "React 应用已挂载");
}

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  exposeApi();
  mountReact();

  // Optional legacy buttons (if they exist in HTML)
  document.getElementById("perceiveBtn")?.addEventListener("click", async () => {
    const result = await perceiveExcel();
    showOutput(result);
  });

  document.getElementById("showAuditBtn")?.addEventListener("click", async () => {
    showOutput(auditList());
  });

  document.getElementById("clearOutput")?.addEventListener("click", () => {
    const outEl = document.getElementById("output");
    if (outEl) outEl.textContent = "";
  });

  document.getElementById("copyOutput")?.addEventListener("click", () => {
    const out = document.getElementById("output")?.textContent || "";
    navigator.clipboard
      ?.writeText(out)
      .then(() => console.log("output copied"))
      .catch((e) => console.error("copy failed", e));
  });

  // 鉁?Startup: probe backend once so UI can show status immediately
  await checkBackend();
});
