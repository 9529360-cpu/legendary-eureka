/**
 * Excel 智能助手 AI Backend (compat)
 * - 兼容旧 eslint parser / Node：不使用 ?. / ?? / replaceAll
 * - 完整 CORS（含 OPTIONS 预检）
 * - /api/health 统一返回 JSON
 * - 速率限制启用
 * - 支持 DELETE /api/config/key
 * - Excel command 模板替换支持对象/数组
 * - JSON 容错（去 ```json 包裹）
 */

"use strict";

require("dotenv").config();

const express = require("express");
const axios = require("axios");
const helmet = require("helmet");
const compression = require("compression");

const app = express();
const port = Number(process.env.AI_BACKEND_PORT || process.env.PORT || 3001);
const NODE_ENV = process.env.NODE_ENV || "development";

// ? 必须用 let（因为支持动态更新 key）
let DEEPSEEK_API_KEY = process.env.DEEPSEEK_API_KEY || "";
const DEFAULT_DEEPSEEK_API_BASE = "https://api.deepseek.com";
const DEFAULT_DEEPSEEK_API_URL = `${DEFAULT_DEEPSEEK_API_BASE}/v1/chat/completions`;
const DEEPSEEK_API_BASE = process.env.DEEPSEEK_API_BASE || "";
const DEEPSEEK_API_URL = resolveDeepSeekApiUrl(
  process.env.DEEPSEEK_API_URL || "",
  DEEPSEEK_API_BASE
);
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || "deepseek-chat";

const RATE_LIMIT_WINDOW_MS = parseInt(process.env.RATE_LIMIT_WINDOW_MS || "900000", 10);
const RATE_LIMIT_MAX_REQUESTS = parseInt(process.env.RATE_LIMIT_MAX_REQUESTS || "100", 10);

const ALLOWED_ORIGINS = (process.env.CORS_ORIGINS ||
  process.env.CORS_ORIGIN ||
  "https://localhost:3000,https://127.0.0.1:3000,http://localhost:3000,http://127.0.0.1:3000")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

function resolveDeepSeekApiUrl(apiUrl, apiBase) {
  const trimmedUrl = String(apiUrl || "").trim();
  if (trimmedUrl) return trimmedUrl;

  const trimmedBase = String(apiBase || "").trim();
  const base = trimmedBase ? trimmedBase.replace(/\/+$/, "") : DEFAULT_DEEPSEEK_API_BASE;

  if (base.endsWith("/chat/completions")) return base;
  if (base.endsWith("/v1")) return base + "/chat/completions";
  return base + "/v1/chat/completions";
}


// ----------------------
// Logging
// ----------------------
const logLevels = { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3 };
const currentLogLevel = process.env.LOG_LEVEL || "INFO";

function log(level, message, data) {
  const levelNum = Object.prototype.hasOwnProperty.call(logLevels, level) ? logLevels[level] : 1;
  const currentLevelNum = Object.prototype.hasOwnProperty.call(logLevels, currentLogLevel)
    ? logLevels[currentLogLevel]
    : 1;
  if (levelNum < currentLevelNum) return;

  const ts = new Date().toISOString();
  // eslint-disable-next-line no-console
  console.log("[" + ts + "] [" + level + "] " + message, data || "");
}

// ----------------------
// Middleware
// ----------------------
app.use(helmet());
app.use(compression());
app.use(express.json({ limit: "2mb" }));

// ? CORS（包含 OPTIONS 预检）
app.use((req, res, next) => {
  const origin = req.headers.origin;

  if (!origin) {
    res.header("Access-Control-Allow-Origin", "*");
  } else if (ALLOWED_ORIGINS.indexOf(origin) >= 0) {
    res.header("Access-Control-Allow-Origin", origin);
    res.header("Vary", "Origin");
  } else {
    res.header("Access-Control-Allow-Origin", "*");
  }

  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept, Authorization"
  );
  res.header("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.header("Access-Control-Max-Age", "86400");

  if (req.method === "OPTIONS") {
    return res.sendStatus(204);
  }
  next();
});

// ----------------------
// Simple Rate Limit
// ----------------------
const rateLimitStore = new Map();

function getClientIp(req) {
  const xf = req.headers["x-forwarded-for"];
  if (xf && typeof xf === "string") {
    return xf.split(",")[0].trim();
  }
  if (Array.isArray(xf) && xf.length > 0) {
    return String(xf[0]).split(",")[0].trim();
  }
  return req.ip || (req.connection && req.connection.remoteAddress) || "unknown";
}

function rateLimitMiddleware(req, res, next) {
  const clientIp = getClientIp(req);
  const now = Date.now();

  const entry = rateLimitStore.get(clientIp);
  if (!entry) {
    rateLimitStore.set(clientIp, { count: 1, firstRequest: now });
    return next();
  }

  const timeDiff = now - entry.firstRequest;
  if (timeDiff > RATE_LIMIT_WINDOW_MS) {
    entry.count = 1;
    entry.firstRequest = now;
    return next();
  }

  entry.count += 1;
  if (entry.count > RATE_LIMIT_MAX_REQUESTS) {
    log("WARN", "Rate limit exceeded", {
      clientIp: clientIp,
      count: entry.count,
      limit: RATE_LIMIT_MAX_REQUESTS,
    });
    return res.status(429).json({
      success: false,
      error: "请求过于频繁，请稍后再试",
      retryAfter: Math.ceil((RATE_LIMIT_WINDOW_MS - timeDiff) / 1000),
    });
  }

  next();
}

// 只对 API 请求限流（健康检查不限制）
app.use(["/chat", "/batch", "/api"], rateLimitMiddleware);

// ----------------------
// Request Timeout Middleware
// ----------------------
const REQUEST_TIMEOUT_MS = parseInt(process.env.REQUEST_TIMEOUT_MS || "60000", 10);

function requestTimeoutMiddleware(req, res, next) {
  // 为每个请求设置超时
  req.setTimeout(REQUEST_TIMEOUT_MS, function() {
    log("WARN", "Request timeout", {
      method: req.method,
      path: req.path,
      timeout: REQUEST_TIMEOUT_MS,
    });
    
    if (!res.headersSent) {
      res.status(408).json({
        success: false,
        error: "请求超时，请稍后重试",
        timeout: REQUEST_TIMEOUT_MS / 1000,
      });
    }
  });
  next();
}

app.use(["/chat", "/batch"], requestTimeoutMiddleware);

// ----------------------
// Request Logging Middleware (Enhanced)
// ----------------------
function requestLoggingMiddleware(req, res, next) {
  const startTime = Date.now();
  const requestId = Math.random().toString(36).substring(7);
  req.requestId = requestId;

  // 响应完成时记录
  res.on("finish", function() {
    const duration = Date.now() - startTime;
    const logLevel = res.statusCode >= 400 ? "WARN" : "INFO";
    
    log(logLevel, "Request completed", {
      requestId: requestId,
      method: req.method,
      path: req.path,
      statusCode: res.statusCode,
      duration: duration + "ms",
    });
  });

  next();
}

app.use(requestLoggingMiddleware);

// ----------------------
// Health
// ----------------------
app.get("/health", (req, res) => {
  res.json({
    status: "ok",
    service: "Excel 智能助手 AI Backend",
    version: "2.0.1-compat",
    environment: NODE_ENV,
  });
});

app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    status: "ok",
    service: "Excel 智能助手 AI Backend",
    version: "2.0.1-compat",
    environment: NODE_ENV,
    port: port,
    allowedOrigins: ALLOWED_ORIGINS,
    model: DEEPSEEK_MODEL,
    configured: !!DEEPSEEK_API_KEY && DEEPSEEK_API_KEY.trim() !== "",
  });
});

// ----------------------
// API Key Store
// ----------------------
let apiKeyStore = {
  key: DEEPSEEK_API_KEY,
  lastUpdated: new Date(),
  isValid: false,
  validationResult: null,
};

async function validateApiKey(apiKey) {
  if (!apiKey || apiKey.trim() === "") {
    return { valid: false, error: "API密钥不能为空" };
  }

  try {
    const resp = await axios.post(
      DEEPSEEK_API_URL,
      {
        model: DEEPSEEK_MODEL,
        messages: [{ role: "user", content: "Hello" }],
        max_tokens: 5,
        stream: false,
      },
      {
        headers: {
          Authorization: "Bearer " + apiKey,
          "Content-Type": "application/json",
        },
        timeout: 10000,
      }
    );

    const model = resp && resp.data ? resp.data.model : undefined;
    const usage = resp && resp.data ? resp.data.usage : undefined;
    return { valid: true, model: model, usage: usage };
  } catch (error) {
    const resp = error && error.response ? error.response : null;
    const data = resp && resp.data ? resp.data : null;
    const msg =
      (data && data.error && data.error.message) ||
      (error && error.message) ||
      "unknown error";
    const code = (resp && resp.status) || (error && error.code);
    return { valid: false, error: msg, code: code };
  }
}

function updateApiKey(newApiKey) {
  DEEPSEEK_API_KEY = newApiKey;
  apiKeyStore = {
    key: newApiKey,
    lastUpdated: new Date(),
    isValid: false,
    validationResult: null,
  };
  return apiKeyStore;
}

// ----------------------
// Excel Operations
// ----------------------
const EXCEL_OPERATIONS = {
  // ===== 数据写入操作 =====
  insert_data: {
    description: "插入数据到指定范围",
    action: "insert",
    category: "write",
  },
  create_table: {
    description: "创建表格",
    action: "createTable",
    category: "write",
  },
  write_range: {
    description: "写入数据到指定范围",
    action: "writeRange",
    category: "write",
  },
  write_cell: {
    description: "写入单个单元格",
    action: "writeCell",
    category: "write",
  },
  
  // ===== 公式操作 =====
  set_formula: {
    description: "设置单元格公式",
    action: "setFormula",
    category: "formula",
  },
  fill_formula: {
    description: "填充公式到范围",
    action: "fillFormula",
    category: "formula",
  },
  
  // ===== 格式化操作 =====
  format_range: {
    description: "格式化单元格范围",
    action: "formatRange",
    category: "format",
  },
  conditional_format: {
    description: "条件格式化",
    action: "conditionalFormat",
    category: "format",
  },
  auto_fit: {
    description: "自动调整列宽/行高",
    action: "autoFit",
    category: "format",
  },
  
  // ===== 数据操作 =====
  sort_range: {
    description: "对范围进行排序",
    action: "sortRange",
    category: "data",
  },
  filter_range: {
    description: "对范围应用筛选器",
    action: "filterRange",
    category: "data",
  },
  remove_duplicates: {
    description: "删除重复项",
    action: "removeDuplicates",
    category: "data",
  },
  find_replace: {
    description: "查找替换",
    action: "findReplace",
    category: "data",
  },
  
  // ===== 图表操作 =====
  create_chart: {
    description: "创建图表",
    action: "createChart",
    category: "chart",
  },
  
  // ===== 工作表操作 =====
  create_sheet: {
    description: "创建新工作表",
    action: "createSheet",
    category: "sheet",
  },
  rename_sheet: {
    description: "重命名工作表",
    action: "renameSheet",
    category: "sheet",
  },
  copy_sheet: {
    description: "复制工作表",
    action: "copySheet",
    category: "sheet",
  },
  delete_sheet: {
    description: "删除工作表",
    action: "deleteSheet",
    category: "sheet",
    highRisk: true,
  },
  switch_sheet: {
    description: "切换到指定工作表",
    action: "switchSheet",
    category: "sheet",
  },
  
  // ===== 跨表操作 =====
  copy_to_sheet: {
    description: "复制数据到另一工作表",
    action: "copyToSheet",
    category: "cross-sheet",
  },
  merge_sheets: {
    description: "合并多个工作表数据",
    action: "mergeSheets",
    category: "cross-sheet",
  },
  
  // ===== 批量操作 =====
  batch_write: {
    description: "批量写入大量数据",
    action: "batchWrite",
    category: "batch",
  },
  batch_formula: {
    description: "批量设置公式",
    action: "batchFormula",
    category: "batch",
  },
  
  // ===== 清理操作 =====
  clear_range: {
    description: "清除范围内容",
    action: "clearRange",
    category: "clear",
    highRisk: true,
  },
  clear_format: {
    description: "清除格式",
    action: "clearFormat",
    category: "clear",
  },
  
  // ===== 分析操作 =====
  analyze_data: {
    description: "分析数据并生成洞察",
    action: "analyzeData",
    category: "analysis",
  },
  
  // ===== 数据透视表 =====
  create_pivot_table: {
    description: "创建数据透视表",
    action: "createPivotTable",
    category: "pivot",
  },
  
  // ===== 命名范围 =====
  create_named_range: {
    description: "创建命名范围",
    action: "createNamedRange",
    category: "named-range",
  },
  delete_named_range: {
    description: "删除命名范围",
    action: "deleteNamedRange",
    category: "named-range",
  },
  
  // ===== 数据验证 =====
  add_data_validation: {
    description: "添加数据验证规则",
    action: "addDataValidation",
    category: "validation",
  },
  remove_data_validation: {
    description: "移除数据验证规则",
    action: "removeDataValidation",
    category: "validation",
  },
  
  // ===== 多步骤操作 =====
  multi_step: {
    description: "执行多个步骤的复合操作",
    action: "multiStep",
    category: "compound",
  },
  
  unknown: {
    description: "未知操作",
    action: "unknown",
    category: "unknown",
  },
};

// ----------------------
// Helpers: Table parameters
// ----------------------
const DEFAULT_SAMPLE_ROWS = 10;

function extractRowCount(text) {
  if (!text) return null;
  const str = String(text);
  let match = str.match(/(\d+)\s*(?:条|行|条数据|条记录)/);
  if (match && match[1]) return parseInt(match[1], 10);
  match = str.match(/(\d+)/);
  if (match && match[1]) return parseInt(match[1], 10);

  const cnMatch = str.match(/([一二三四五六七八九十])\s*(?:条|行)/);
  if (cnMatch && cnMatch[1]) {
    const map = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
    return map[cnMatch[1]] || null;
  }
  return null;
}

function extractColumnCount(text) {
  if (!text) return null;
  const str = String(text);
  const match = str.match(/(\d+)\s*列/);
  if (match && match[1]) return parseInt(match[1], 10);
  return null;
}

function extractTableName(text) {
  if (!text) return null;
  const match = String(text).match(/(?:名为|名称为|命名为|叫)([^，。,\s]+)/);
  if (match && match[1]) return match[1].trim();
  return null;
}

function splitHeaderTokens(segment) {
  if (!segment) return [];
  let cleaned = String(segment)
    .replace(/等.*$/, "")
    .replace(/(列|字段|项目).*$/, "")
    .replace(/\s+/g, "");
  const primary = cleaned.split(/[、，,;；\n]/);
  const tokens = [];
  for (let i = 0; i < primary.length; i++) {
    const parts = String(primary[i]).split(/以及|和|及/);
    for (let j = 0; j < parts.length; j++) {
      const token = String(parts[j]).trim();
      if (token) tokens.push(token);
    }
  }
  return tokens;
}

function extractHeadersFromText(text) {
  if (!text) return null;
  const str = String(text);
  let match = str.match(/(?:包含|包括|含有|含)([^。\n]+)/);
  if (match && match[1]) {
    const headers = splitHeaderTokens(match[1]);
    if (headers.length) return headers;
  }

  match = str.match(/(?:字段|列|表头)[:：]\s*([^。\n]+)/);
  if (match && match[1]) {
    const headers = splitHeaderTokens(match[1]);
    if (headers.length) return headers;
  }

  return null;
}

function coerceCellValue(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  return JSON.stringify(value);
}

function inferSampleValue(header, index) {
  const label = String(header || "");
  const lower = label.toLowerCase();
  const rowIndex = index + 1;

  if (label.indexOf("日期") !== -1 || label.indexOf("时间") !== -1 || lower.indexOf("date") !== -1) {
    const base = new Date();
    base.setDate(base.getDate() + index);
    const y = base.getFullYear();
    const m = String(base.getMonth() + 1).padStart(2, "0");
    const d = String(base.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + d;
  }
  if (label.indexOf("数量") !== -1 || lower.indexOf("qty") !== -1) {
    return rowIndex * 3;
  }
  if (
    label.indexOf("金额") !== -1 ||
    label.indexOf("价格") !== -1 ||
    label.indexOf("单价") !== -1 ||
    label.indexOf("总价") !== -1 ||
    label.indexOf("成本") !== -1 ||
    label.indexOf("收入") !== -1 ||
    lower.indexOf("price") !== -1 ||
    lower.indexOf("amount") !== -1
  ) {
    return rowIndex * 25;
  }
  if (label.indexOf("率") !== -1 || lower.indexOf("rate") !== -1) {
    return Number((0.05 * rowIndex).toFixed(2));
  }
  if (
    label.indexOf("人员") !== -1 ||
    label.indexOf("员工") !== -1 ||
    label.indexOf("姓名") !== -1 ||
    label.indexOf("客户") !== -1
  ) {
    return "人员" + rowIndex;
  }
  if (label.indexOf("类别") !== -1 || label.indexOf("类型") !== -1) {
    return "类型" + ((rowIndex % 3) + 1);
  }
  if (label.indexOf("商品") !== -1 || label.indexOf("产品") !== -1 || label.indexOf("名称") !== -1) {
    return "商品" + rowIndex;
  }
  if (label.indexOf("支付") !== -1 || label.indexOf("方式") !== -1 || label.indexOf("渠道") !== -1) {
    return "方式" + ((rowIndex % 3) + 1);
  }
  if (label.indexOf("地区") !== -1 || label.indexOf("城市") !== -1 || label.indexOf("区域") !== -1) {
    return "区域" + ((rowIndex % 3) + 1);
  }
  if (label.indexOf("编号") !== -1 || lower.indexOf("id") !== -1) {
    return "NO-" + String(rowIndex).padStart(3, "0");
  }

  return "样例" + rowIndex;
}

function normalizeHeaders(parameters) {
  if (!parameters) return null;
  const list = parameters.headers || parameters.columns || parameters.fields;
  if (!Array.isArray(list)) return null;
  const headers = list
    .map((item) => {
      if (typeof item === "string" || typeof item === "number" || typeof item === "boolean") {
        return String(item);
      }
      if (item && typeof item === "object" && item.name) {
        return String(item.name);
      }
      return null;
    })
    .filter(Boolean);
  return headers.length ? headers : null;
}

function normalizeTableData(raw, headers) {
  if (!raw) return null;
  if (Array.isArray(raw)) {
    if (raw.length === 0) return [];
    if (Array.isArray(raw[0])) {
      return raw.map((row) => row.map(coerceCellValue));
    }
    if (typeof raw[0] === "object") {
      const keys = headers || Object.keys(raw[0] || {});
      if (!keys.length) return null;
      return raw.map((row) => keys.map((key) => coerceCellValue(row[key])));
    }
    return [raw.map(coerceCellValue)];
  }
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if ((trimmed.startsWith("[") && trimmed.endsWith("]")) || (trimmed.startsWith("{") && trimmed.endsWith("}"))) {
      try {
        const parsed = JSON.parse(trimmed);
        return normalizeTableData(parsed, headers);
      } catch (e) {
        return [[raw]];
      }
    }
    return [[raw]];
  }
  if (typeof raw === "object") {
    return normalizeTableData(raw.values || raw.data || raw.rows || raw.table, headers);
  }
  return null;
}

function buildSampleRows(headers, count) {
  const rows = [];
  const total = Math.max(1, count || DEFAULT_SAMPLE_ROWS);
  for (let i = 0; i < total; i++) {
    const row = [];
    for (let j = 0; j < headers.length; j++) {
      row.push(inferSampleValue(headers[j], i));
    }
    rows.push(row);
  }
  return rows;
}

function buildDefaultHeaders(count) {
  const total = Math.max(1, count || 5);
  const headers = [];
  for (let i = 0; i < total; i++) {
    headers.push("列" + String(i + 1));
  }
  return headers;
}

function inferOperationFromText(text) {
  if (!text) return null;
  const str = String(text);
  if (/(表格|登记表|报表|清单|台账|清册|明细)/.test(str)) return "create_table";
  if (/(写入|插入|填充|追加|更新|导入)/.test(str)) return "insert_data";
  if (/(公式)/.test(str)) return "set_formula";
  if (/(图表|图形)/.test(str)) return "create_chart";
  if (/(筛选|过滤)/.test(str)) return "filter_range";
  if (/(排序)/.test(str)) return "sort_range";
  if (/(格式|高亮|着色|颜色|加粗)/.test(str)) return "format_range";
  return null;
}

function buildUserPrompt(userMessage, context) {
  let prompt = userMessage || "";
  
  // ===== 工作簿全局上下文 (类似IDE的工作区感知) =====
  if (context && context.workbookContext) {
    const wb = context.workbookContext;
    prompt += "\n\n【工作簿全局上下文】";
    prompt += "\n文件名: " + (wb.fileName || "工作簿");
    prompt += "\n工作表数: " + (wb.sheets ? wb.sheets.length : 0);
    
    // 工作表列表
    if (wb.sheets && wb.sheets.length > 0) {
      prompt += "\n\n工作表列表:";
      for (var i = 0; i < wb.sheets.length && i < 10; i++) {
        var sheet = wb.sheets[i];
        prompt += "\n  " + (i + 1) + ". \"" + sheet.name + "\"";
        if (sheet.hasData) {
          prompt += " - " + sheet.rowCount + "行×" + sheet.columnCount + "列";
        } else {
          prompt += " - 空";
        }
        if (sheet.hasTables) prompt += " [有表格]";
        if (sheet.hasCharts) prompt += " [有图表]";
        if (sheet.hasPivotTables) prompt += " [有透视表]";
      }
      if (wb.sheets.length > 10) {
        prompt += "\n  ... 还有" + (wb.sheets.length - 10) + "个工作表";
      }
    }
    
    // 表格列表
    if (wb.tables && wb.tables.length > 0) {
      prompt += "\n\n表格列表:";
      for (var j = 0; j < wb.tables.length && j < 5; j++) {
        var table = wb.tables[j];
        prompt += "\n  - \"" + table.name + "\" 在 " + table.sheetName;
        prompt += ": " + table.rowCount + "行";
        if (table.columns && table.columns.length > 0) {
          prompt += ", 列: [" + table.columns.slice(0, 5).join(", ") + "]";
          if (table.columns.length > 5) {
            prompt += "...";
          }
        }
      }
    }
    
    // 命名范围
    if (wb.namedRanges && wb.namedRanges.length > 0) {
      prompt += "\n\n命名范围:";
      for (var k = 0; k < wb.namedRanges.length && k < 5; k++) {
        var nr = wb.namedRanges[k];
        prompt += "\n  - " + nr.name + " = " + nr.address;
      }
      if (wb.namedRanges.length > 5) {
        prompt += "\n  ... 还有" + (wb.namedRanges.length - 5) + "个";
      }
    }
    
    // 图表
    if (wb.charts && wb.charts.length > 0) {
      prompt += "\n\n图表: " + wb.charts.map(function(c) { return c.name + "(" + c.chartType + ")"; }).join(", ");
    }
    
    // 统计信息
    prompt += "\n\n统计: " + (wb.totalCellsWithData || 0) + "个数据单元格, " + (wb.totalFormulas || 0) + "个公式";
    prompt += "\n质量评分: " + (wb.overallQualityScore || 0) + "/100";
    
    // 问题提示
    if (wb.issues && wb.issues.length > 0) {
      prompt += "\n\n注意事项:";
      for (var m = 0; m < wb.issues.length && m < 3; m++) {
        var issue = wb.issues[m];
        prompt += "\n  [" + issue.type + "] " + issue.message;
      }
    }
  }
  
  // ===== 当前选区上下文 =====
  if (context && context.selection) {
    const selection = context.selection;
    const sample =
      selection && Array.isArray(selection.values)
        ? selection.values.slice(0, 5).map((row) => (Array.isArray(row) ? row.slice(0, 6) : row))
        : [];
    prompt +=
      "\n\n[当前选区]\n" +
      "地址: " +
      selection.address +
      "\n行数: " +
      selection.rowCount +
      ", 列数: " +
      selection.columnCount;
    if (sample.length > 0) {
      prompt += "\n数据示例: " + JSON.stringify(sample);
    }
  }
  
  // ===== 用户偏好上下文 =====
  if (context && context.userPreferences) {
    var prefs = context.userPreferences;
    prompt += "\n\n[用户偏好]";
    if (prefs.defaultChartType) {
      prompt += "\n默认图表类型: " + prefs.defaultChartType;
    }
    if (prefs.favoriteOperations && prefs.favoriteOperations.length > 0) {
      prompt += "\n常用操作: " + prefs.favoriteOperations.slice(0, 5).join(", ");
    }
    if (prefs.lastUsedOperations && prefs.lastUsedOperations.length > 0) {
      var recentOps = prefs.lastUsedOperations.slice(0, 3).map(function(op) { return op.operation; });
      prompt += "\n最近操作: " + recentOps.join(", ");
    }
  }
  
  return prompt;
}

function ensureTabularParameters(aiResponse, userMessage) {
  if (!aiResponse || !aiResponse.operation) return aiResponse;
  let op = aiResponse.operation;

  if (!Object.prototype.hasOwnProperty.call(EXCEL_OPERATIONS, op) || op === "unknown") {
    const inferred =
      inferOperationFromText(userMessage) ||
      inferOperationFromText(aiResponse.description) ||
      inferOperationFromText(aiResponse.explanation);
    if (inferred) {
      op = inferred;
      aiResponse.operation = inferred;
    }
  }

  if (op !== "create_table" && op !== "insert_data" && op !== "write_range") return aiResponse;

  const parameters = aiResponse.parameters || {};
  let headers = normalizeHeaders(parameters);
  if (!headers) {
    headers =
      extractHeadersFromText(aiResponse.description) ||
      extractHeadersFromText(aiResponse.explanation) ||
      extractHeadersFromText(userMessage);
  }

  const rowCount =
    extractRowCount(userMessage) ||
    extractRowCount(aiResponse.explanation) ||
    extractRowCount(aiResponse.description) ||
    DEFAULT_SAMPLE_ROWS;

  const columnCount =
    extractColumnCount(userMessage) ||
    extractColumnCount(aiResponse.explanation) ||
    extractColumnCount(aiResponse.description) ||
    null;

  const rawData =
    parameters.values ||
    parameters.data ||
    parameters.rows ||
    parameters.table ||
    parameters.sampleData ||
    parameters.samples ||
    null;

  let data = normalizeTableData(rawData, headers);
  if ((!headers || headers.length === 0) && (!data || data.length === 0)) {
    headers = buildDefaultHeaders(columnCount || 5);
  }
  if ((!data || data.length === 0) && headers && headers.length) {
    data = buildSampleRows(headers, rowCount);
  }

  if (headers && !parameters.headers) {
    parameters.headers = headers;
  }
  if (data && !parameters.data && !parameters.values && !parameters.rows) {
    parameters.data = data;
  }
  if (!parameters.sampleCount) {
    parameters.sampleCount = rowCount;
  }

  const tableName =
    parameters.tableName ||
    parameters.name ||
    extractTableName(aiResponse.description) ||
    extractTableName(aiResponse.explanation) ||
    extractTableName(userMessage);
  if (tableName && !parameters.tableName && !parameters.name) {
    parameters.tableName = tableName;
  }

  aiResponse.parameters = parameters;
  return aiResponse;
}

// ----------------------
// AI Integration
// ----------------------

/**
 * 构建增强的系统提示词
 * 包含：操作类型、上下文感知、意图推断、智能建议
 */
function buildEnhancedSystemPrompt() {
  return `你是Excel智能助手，将用户自然语言指令转换为精确的Excel操作JSON。

## 🚨 核心规则（必须遵守）

### 1. 模糊请求必须先澄清
当用户请求不明确时，**绝对不能直接执行**，必须返回 clarify 操作询问：
- "清理表格" / "清里表格" → 问：清理什么？删除空行？格式化？删除重复？
- "删除没用的" / "del没用的" → 问：什么算"没用"？哪些列/行？判断标准是什么？
- "优化一下" / "简单点" → 问：优化什么方面？格式？结构？公式？
- "处理一下" / "搞一下" / "搞好看" → 问：具体想要什么效果？
- "把重复的删掉" → 问：哪列判断重复？保留第一个还是最后一个？
- "整理整理" → 问：需要怎样整理？排序？格式化？清理？
- "那几列删了" / "这些都删了" / "把这些删掉" → 问：具体是哪几列/哪些内容？
- "delete掉/del掉column/col" → 问：删除哪些列？判断标准是什么？

**触发澄清的关键词和短语**：
- 中文：清理、优化、整理、处理、搞、删没用的、重复、乱、丑、好看、简单、专业
- 英文混合：delete、del、column、col、那些、这些、那几、把...删
- 模糊指代：那个、这个、它们、这些、那些、上面的
- 错别字：清里（清理）、删出（删除）

### 2. 特殊上下文必须澄清
- **有筛选状态时的删除操作**：用户说"删除空行"但表格有筛选时，必须询问是删可见行还是全部行
- **批量填充空值**：空值可能有语义（代表缺失/不适用），必须问用户是否确定
- **跨表操作**：涉及多个工作表时，必须先确认是哪个Sheet/工作表
- **公式依赖**：删除包含公式的列/行时，必须警告可能影响其他单元格的公式依赖

### 3. 结构识别必须澄清
- **多Sheet场景**：当工作簿有多个工作表时，"汇总数据"等操作必须先问"哪个工作表/Sheet"
- **跨Sheet引用**：引用其他表数据时，必须问"哪些数据"、"放在哪个位置"、"引用方式"
- **隐藏列/行**：有隐藏内容时，必须提醒用户存在隐藏的数据
- **合并单元格**：有合并单元格时，必须警告可能影响操作结果
- **公式列删除**：删除包含公式的列时，必须提醒"公式"、"依赖"、"影响"

### 4. 多步任务处理
当用户请求包含多个步骤（如"先...再..."）：
- 第一步必须是澄清，不能直接执行
- 拆解任务并逐步确认

## 澄清响应格式
当需要澄清时，返回：
{"operation":"clarify","parameters":{"questions":["问题1","问题2"],"options":["选项1","选项2"]},"explanation":"我需要先确认一些细节...","confidence":0.9}

## 执行操作格式（重要：你只负责返回操作，不负责判断是否需要确认）
当用户请求明确时，直接返回操作：
{"operation":"delete_rows","parameters":{"condition":"空行","scope":"全表","estimatedRows":100},"explanation":"将删除表格中所有空行","confidence":0.9}

**注意**：
- 你只负责解析用户意图，生成操作参数
- 是否需要用户确认由 Agent 层决定，不是你的职责
- 即使是高风险操作，你也直接返回操作内容，Agent 会处理确认流程

## 重要规则
1. **只输出纯JSON**，不要任何额外文字、代码块或Markdown
2. 确保JSON可被JavaScript的JSON.parse()直接解析
3. 复杂任务使用 multi_step 操作，包含 steps 数组
4. **理解对话上下文**：当用户说"这个"、"它"、"上面的"时，参考对话历史理解指代
5. **宁可多问，不可误删** - 对于任何模糊的删除/修改请求，优先澄清
6. **注意环境信息**：仔细阅读工作簿环境和数据特征，识别筛选状态、多Sheet、隐藏列、公式依赖等特殊情况

## 对话上下文理解
- 用户说"给它排序" → 理解"它"指的是之前提到的数据/表格
- 用户说"再加一列" → 在之前创建的表格基础上添加
- 用户说"格式和刚才一样" → 复用之前的格式设置
- 用户说"撤销" → 执行撤销操作（前端处理）
- 如果用户的指令不明确，**优先澄清**而不是猜测执行

## 支持的操作类型

### 澄清和确认（优先级最高）
- **clarify**: { "questions": ["问题1"], "options": ["选项1", "选项2"], "context": "需要澄清的原因" }
- **confirm**: { "action": "操作名", "impactScope": "影响范围描述", "requireConfirmation": true }

### 数据写入
- **insert_data / write_range**: { "headers": ["列1"], "data": [[值]], "address": "A1" }
- **create_table**: { "headers": ["列1"], "data": [[值]], "tableName": "表名" }
- **batch_write**: { "address": "A1", "data": [[大量数据]], "batchSize": 1000 }

### 公式
- **set_formula**: { "address": "B1", "formula": "=SUM(A:A)" }
- **fill_formula**: { "startAddress": "B1", "endAddress": "B100", "formula": "=A1*2" }
- **batch_formula**: { "formulas": [{"address":"B1","formula":"=SUM(A:A)"},{"address":"C1","formula":"=AVERAGE(A:A)"}] }

### 格式化
- **format_range**: { "address": "A1:D10", "color": "#4472C4", "bold": true, "fontColor": "#FFFFFF" }
- **conditional_format**: { "address": "A1:A100", "rule": "greaterThan", "value": 100, "color": "#00FF00" }
- **auto_fit**: { "columns": "A:D" } 或 { "rows": "1:10" }

### 数据操作
- **sort_range**: { "address": "A1:D10", "column": 0, "ascending": true }
- **filter_range**: { "address": "A1:D10", "column": 0, "criteria": ">100" }
- **remove_duplicates**: { "address": "A1:D100", "columns": [0, 1] }
- **find_replace**: { "find": "旧值", "replace": "新值", "address": "A1:D100" }

### 图表
- **create_chart**: { "chartType": "column|bar|line|pie|scatter|area", "dataRange": "A1:B10", "chartName": "图表" }

### 工作表操作
- **create_sheet**: { "sheetName": "新工作表" }
- **rename_sheet**: { "oldName": "Sheet1", "newName": "数据表" }
- **copy_sheet**: { "sourceSheet": "Sheet1", "newName": "Sheet1副本" }
- **delete_sheet**: { "sheetName": "要删除的表" }
- **switch_sheet**: { "sheetName": "目标表" }

### 跨表操作
- **copy_to_sheet**: { "sourceAddress": "A1:D10", "targetSheet": "Sheet2", "targetAddress": "A1" }
- **merge_sheets**: { "sourceSheets": ["Sheet1","Sheet2"], "targetSheet": "汇总", "startRow": 1 }

### 清理操作
- **clear_range**: { "address": "A1:D100", "clearType": "all|contents|formats" }
- **clear_format**: { "address": "A1:D100" }

### 数据透视表
- **create_pivot_table**: { "sourceAddress": "A1:E100", "pivotTableName": "透视表1", "rowField": "产品", "columnField": "区域", "valueField": "销量", "aggregation": "sum|count|average" }

### 命名范围
- **create_named_range**: { "name": "销售数据", "address": "A1:D100", "comment": "可选注释" }
- **delete_named_range**: { "name": "销售数据" }

### 数据验证
- **add_data_validation**: { "address": "A1:A100", "type": "list|number|date|textLength", "values": ["选项1","选项2"] } 或 { "type": "number", "min": 0, "max": 100 }
- **remove_data_validation**: { "address": "A1:A100" }

### 多步骤复合操作
- **multi_step**: { "steps": [ {操作1}, {操作2}, ... ], "description": "任务描述" }

## 🚨 澄清示例（必须学习）

用户: "这个表太乱了，帮我清理一下"
输出:
{"operation":"clarify","parameters":{"questions":["您希望如何清理表格？","是要删除空行、调整格式、还是删除重复数据？"],"options":["删除空行","删除重复数据","调整格式和对齐","删除无用列"]},"explanation":"这个表格可以从多个方面进行清理，请告诉我您最想改善的是哪个方面？","confidence":0.95}

用户: "删除没用的列"
输出:
{"operation":"clarify","parameters":{"questions":["什么样的列算'没用'？","您是指空列、数据很少的列、还是特定的某些列？"],"options":["删除完全空白的列","删除数据少于50%的列","让我指定要删除的列"]},"explanation":"我需要先了解您判断'没用'的标准，才能安全地执行删除操作。","confidence":0.95}

用户: "把重复的删掉"
输出:
{"operation":"clarify","parameters":{"questions":["根据哪一列或哪几列来判断重复？","保留第一条还是最后一条重复记录？"],"options":["根据所有列判断","根据第一列判断","让我指定列"]},"explanation":"删除重复数据需要明确判断标准，请告诉我如何识别重复。","confidence":0.95}

用户: "delete掉那些没用的column"
输出:
{"operation":"clarify","parameters":{"questions":["您说的'没用的column'具体指哪些列？","判断标准是什么？空列？还是数据不完整的列？"],"options":["删除空白列","删除数据不足50%的列","让我指定具体列名"]},"explanation":"我需要确认您想删除哪些列，以及判断'没用'的标准。","confidence":0.95}

用户: "那几列删了"
输出:
{"operation":"clarify","parameters":{"questions":["您指的是哪几列？","请告诉我具体的列名或列号（如A、B、C列）"],"options":["让我查看表格后指定","删除空白列","删除最后几列"]},"explanation":"我不确定您想删除哪几列，请具体说明列名或位置。","confidence":0.95}

用户: "把这些都删了"
输出:
{"operation":"clarify","parameters":{"questions":["您说的'这些'具体指什么？","是某些行、某些列、还是某个区域的数据？"],"options":["删除选中的区域","删除特定的行","删除特定的列","让我具体说明"]},"explanation":"请告诉我您想删除的具体内容，我需要明确范围才能安全操作。","confidence":0.95}

用户: "del掉没用的col"
输出:
{"operation":"clarify","parameters":{"questions":["哪些列算'没用的'？","请告诉我判断标准或具体列名"],"options":["删除完全空白的列","删除数据稀疏的列","让我指定列"]},"explanation":"为了安全删除，我需要先了解您对'没用'的定义。","confidence":0.95}

用户: "帮我看看这个表有没有问题"
输出:
{"operation":"clarify","parameters":{"questions":["您想检查哪方面的问题？","是数据格式、公式错误、还是数据完整性？"],"options":["检查数据格式问题","检查公式错误","检查空值和缺失","检查重复数据"]},"explanation":"表格可能存在多种问题，请告诉我您主要关注哪个方面？","confidence":0.95}

用户: "删除所有空行" (上下文: 表格当前有筛选状态，只显示部分行)
输出:
{"operation":"clarify","parameters":{"questions":["表格当前有筛选状态，您想删除哪些空行？","是删除当前可见的空行，还是删除全部空行（包括被筛选隐藏的）？"],"options":["只删除当前可见的空行","删除所有空行（需要先清除筛选）","让我先查看筛选条件"]},"explanation":"您的表格当前有筛选状态，有些行可能被隐藏了。请确认您想删除的范围，以免误删隐藏的数据。","confidence":0.95}

用户: "把所有空值填成0"
输出:
{"operation":"clarify","parameters":{"questions":["这会将工作表中所有空白单元格都填充为数字0，可能会改变数据的含义。","空值有时表示'未知'或'不适用'，填成0可能导致统计结果错误。您确定要这样做吗？"],"options":["是，将所有空值填成0","只填充数值列的空值","只填充指定范围的空值","让我重新考虑"]},"explanation":"空值和0在数据分析中含义不同。空值可能代表数据缺失或不适用，而0是一个实际的数值。批量填充前请确认这是您想要的效果。","confidence":0.95}

## 🚨 结构识别示例（必须学习）

用户: "把所有数据汇总一下" (上下文: 工作簿有多个工作表 Sheet1, Sheet2, 汇总)
输出:
{"operation":"clarify","parameters":{"questions":["当前工作簿有多个工作表，您想汇总哪个Sheet的数据？","还是要把所有工作表的数据合并到一起？"],"options":["汇总当前工作表(Sheet1)的数据","汇总所有工作表的数据","让我指定要汇总的工作表"]},"explanation":"您的工作簿有多个Sheet，请告诉我您想汇总哪个工作表的数据。","confidence":0.95}

用户: "把Sheet2的数据引用过来" (上下文: 有多个工作表)
输出:
{"operation":"clarify","parameters":{"questions":["您想引用Sheet2的哪些数据？","引用后放在当前表的哪个位置？","是使用公式引用还是直接复制数据？"],"options":["引用全部数据","让我指定数据范围","使用公式动态引用","直接复制粘贴"]},"explanation":"跨表引用需要明确数据范围、目标位置和引用方式，请提供更多细节。","confidence":0.95}

用户: "删除C列" (上下文: C列包含公式，被其他单元格引用)
输出:
{"operation":"clarify","parameters":{"questions":["C列包含公式，删除后会影响依赖这些公式的其他单元格。","依赖C列的单元格可能会显示错误(#REF!)。您确定要删除吗？"],"options":["确认删除，我了解公式依赖的影响","让我先查看哪些单元格依赖C列","取消删除"]},"explanation":"⚠️ C列的公式被其他单元格引用，删除后可能导致公式错误。请确认您了解这个影响。","confidence":0.95}

用户: "删除多余的列" (上下文: 表格有隐藏列)
输出:
{"operation":"clarify","parameters":{"questions":["当前表格有隐藏列，您指的'多余'包括隐藏的列吗？","请告诉我具体要删除哪些列"],"options":["只删除可见的空白列","包括隐藏列一起处理","让我先取消隐藏查看所有列"]},"explanation":"表格中有隐藏列，在删除前请确认是否包括这些隐藏的内容。","confidence":0.95}

## 执行操作示例（用户意图明确时直接返回操作）

用户: "把状态全改成完成"
输出:
{"operation":"batch_update","parameters":{"targetColumn":"状态","newValue":"完成","scope":"全列"},"explanation":"将状态列的所有单元格修改为'完成'","confidence":0.9}

用户: "删除所有2023年之前的数据"
输出:
{"operation":"delete_rows","parameters":{"condition":"日期 < 2023-01-01","scope":"符合条件的行"},"explanation":"删除日期早于2023年的所有行","confidence":0.9}

用户: "把所有金额乘以1.1"
输出:
{"operation":"batch_formula","parameters":{"targetColumn":"金额","formula":"原值 * 1.1","scope":"全列"},"explanation":"将金额列所有数值乘以1.1（增加10%）","confidence":0.9}

用户: "把C列删了"
输出:
{"operation":"delete_column","parameters":{"column":"C"},"explanation":"删除C列","confidence":0.9}

用户: "把利润公式填到整列"
输出:
{"operation":"fill_formula","parameters":{"formula":"利润公式","scope":"整列"},"explanation":"将利润公式填充到整列","confidence":0.9}

## 正常执行示例

用户: "创建一个销售报表，包含数据、汇总公式和图表"
输出:
{"operation":"multi_step","parameters":{"description":"创建完整销售报表","steps":[{"operation":"create_table","parameters":{"headers":["日期","产品","销量","金额"],"data":[["2024-01-01","产品A",100,1000],["2024-01-02","产品B",150,1500]],"tableName":"销售数据"}},{"operation":"set_formula","parameters":{"address":"E1","formula":"=SUM(D:D)"}},{"operation":"create_chart","parameters":{"chartType":"column","dataRange":"A1:D3","chartName":"销售图表"}}]},"explanation":"创建销售数据表、添加汇总公式并生成图表","confidence":0.95}

用户: "给大于1000的金额标红色"
输出:
{"operation":"conditional_format","parameters":{"address":"D2:D100","rule":"greaterThan","value":1000,"format":{"fontColor":"#FF0000","bold":true}},"explanation":"对金额列应用条件格式，大于1000的显示红色加粗","confidence":0.9}

## 输出格式
{
  "operation": "操作类型",
  "parameters": { /* 具体参数 */ },
  "explanation": "用户友好的中文解释",
  "confidence": 0.9
}`;
}

async function callDeepSeekAI(userMessage, context = {}) {
  // 验证API密钥
  const validation = await validateApiKey(DEEPSEEK_API_KEY);
  apiKeyStore.validationResult = validation;
  if (!validation.valid) {
    throw new Error("API密钥无效: " + validation.error);
  }
  apiKeyStore.isValid = true;

  const systemPrompt = buildEnhancedSystemPrompt();

  try {
    const history =
      context && Array.isArray(context.conversationHistory) ? context.conversationHistory : [];
    const trimmedHistory = history
      .filter((item) => item && item.role && item.content)
      .slice(-6)
      .map((item) => ({ role: item.role, content: item.content }));

    const userPrompt = buildUserPrompt(userMessage, context);

    const resp = await axios.post(
      DEEPSEEK_API_URL,
      {
        model: DEEPSEEK_MODEL,
        messages: [
          { role: "system", content: systemPrompt },
          ...trimmedHistory,
          { role: "user", content: userPrompt },
        ],
        temperature: 0.2,
        max_tokens: 800,
        stream: false,
      },
      {
        headers: {
          Authorization: "Bearer " + DEEPSEEK_API_KEY,
          "Content-Type": "application/json",
        },
        timeout: 30000,
      }
    );

    const content =
      resp && resp.data && resp.data.choices && resp.data.choices[0] && resp.data.choices[0].message
        ? resp.data.choices[0].message.content
        : "";

    const cleaned = String(content)
      .replace(/```json\s*/i, "")
      .replace(/```\s*$/i, "")
      .trim();

    try {
      const parsed = JSON.parse(cleaned);
      if (!parsed.operation || !parsed.explanation) {
        throw new Error("AI响应缺少必需字段");
      }
      return parsed;
    } catch (e) {
      log("WARN", "AI响应不是严格JSON，走降级", { err: e.message, content: cleaned });
      return {
        operation: "unknown",
        description: "AI回复",
        parameters: {},
        explanation: cleaned || content || "（空响应）",
        confidence: 0.3,
      };
    }
  } catch (error) {
    const resp = error && error.response ? error.response : null;
    const data = resp && resp.data ? resp.data : null;
    const msg =
      (data && data.error && data.error.message) ||
      (error && error.message) ||
      "unknown error";
    throw new Error("AI服务暂时不可用: " + msg);
  }
}

// ----------------------
// Routes: API Key
// ----------------------
app.post("/api/config/key", async (req, res) => {
  try {
    const apiKey = req.body && req.body.apiKey ? req.body.apiKey : "";
    if (!apiKey) {
      return res.status(400).json({ success: false, error: "API密钥不能为空" });
    }

    const validation = await validateApiKey(apiKey);
    if (!validation.valid) {
      return res.status(400).json({ success: false, error: validation.error, code: validation.code });
    }

    const updated = updateApiKey(apiKey);
    updated.isValid = true;
    updated.validationResult = validation;

    res.json({
      success: true,
      message: "API密钥已更新并验证成功",
      lastUpdated: updated.lastUpdated,
      model: validation.model,
      usage: validation.usage,
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.delete("/api/config/key", (req, res) => {
  try {
    const updated = updateApiKey("");
    updated.isValid = false;
    updated.validationResult = null;

    res.json({ success: true, message: "API密钥已清除" });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get("/api/config/status", (req, res) => {
  const key = DEEPSEEK_API_KEY || "";
  res.json({
    success: true,
    configured: key.trim() !== "",
    isValid: apiKeyStore.isValid,
    lastUpdated: apiKeyStore.lastUpdated,
    maskedKey: key
      ? key.substring(0, 8) + "..." + key.substring(Math.max(0, key.length - 4))
      : null,
    model: DEEPSEEK_MODEL,
  });
});

// ----------------------
// Routes: Agent Chat (ReAct Mode)
// ----------------------
app.post("/agent/chat", async (req, res) => {
  try {
    const message = req.body && req.body.message ? req.body.message : "";
    const systemPrompt = req.body && req.body.systemPrompt ? req.body.systemPrompt : "";
    const responseFormat = req.body && req.body.responseFormat ? req.body.responseFormat : "text";

    if (!message) {
      return res.status(400).json({ success: false, error: "消息内容不能为空" });
    }

    if (!systemPrompt) {
      return res.status(400).json({ success: false, error: "Agent 模式需要 systemPrompt" });
    }

    // v2.9.24: 打印更详细的日志，包括执行历史
    const hasHistory = message.includes("## 执行历史");
    const historyPreview = hasHistory 
      ? message.substring(message.indexOf("## 执行历史"), message.indexOf("## 执行历史") + 300)
      : "(无执行历史)";
    log("INFO", "Agent request", { 
      taskPreview: message.substring(0, 100),
      hasHistory,
      historyPreview: historyPreview.substring(0, 200),
      messageLength: message.length 
    });

    // 验证API密钥
    const validation = await validateApiKey(DEEPSEEK_API_KEY);
    if (!validation.valid) {
      throw new Error("API密钥无效: " + validation.error);
    }

    // v2.9.31: 必须开启 JSON 模式！DeepSeek 官方文档明确支持
    // response_format: { type: "json_object" } 会保证输出是有效 JSON
    // 但前提是 prompt 要明确要求输出 JSON
    // v2.9.70: deepseek-chat 模型最大支持 8192 output tokens
    const resp = await axios.post(
      DEEPSEEK_API_URL,
      {
        model: DEEPSEEK_MODEL,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: message },
        ],
        temperature: 0.2,
        max_tokens: 8192, // v2.9.70: DeepSeek Chat 最大输出限制
        stream: false,
        // v2.9.31: 恢复 JSON 模式，这是 DeepSeek 保证 JSON 完整性的关键
        ...(responseFormat === "json" ? { response_format: { type: "json_object" } } : {}),
      },
      {
        headers: {
          Authorization: "Bearer " + DEEPSEEK_API_KEY,
          "Content-Type": "application/json",
        },
        timeout: 120000, // v2.9.30: 增加超时到 120 秒
      }
    );

    // v2.9.71: DeepSeek 模型说明
    // - deepseek-chat (V3): 默认不开启思考模式，无 reasoning_content
    // - deepseek-reasoner: 默认开启思考模式，会有 reasoning_content
    // - 若要在 V3 上开启思考: extra_body: { thinking: { type: "enabled" } }
    // 目前我们使用 deepseek-chat，不开启思考模式
    const messageObj = resp?.data?.choices?.[0]?.message || {};
    const content = messageObj.content || "";
    const reasoningContent = messageObj.reasoning_content || "";
    
    // 记录推理内容长度（只有 deepseek-reasoner 或开启 thinking 才会有）
    if (reasoningContent) {
      log("INFO", "DeepSeek reasoning (思考模式)", { 
        reasoningLength: reasoningContent.length,
        reasoningPreview: reasoningContent.substring(0, 300) + "..."
      });
    }
    
    // v2.9.25: 检查 finish_reason，如果是 length 说明被截断了
    const finishReason = resp?.data?.choices?.[0]?.finish_reason;
    const usage = resp?.data?.usage || {};
    
    if (finishReason === "length") {
      log("WARN", "Agent response truncated!", { 
        finishReason,
        contentLength: content.length,
        reasoningLength: reasoningContent.length,
        totalTokens: usage.total_tokens,
        completionTokens: usage.completion_tokens,
        contentPreview: content.substring(content.length - 100)
      });
    }

    log("INFO", "Agent response", { 
      content: content.substring(0, 200),
      finishReason,
      totalTokens: usage.total_tokens,
      completionTokens: usage.completion_tokens,
      reasoningTokens: usage.completion_tokens_details?.reasoning_tokens || 0
    });

    // v2.9.26: 将 finishReason 返回给前端，让前端知道是否截断
    res.json({
      success: true,
      message: content,
      finishReason: finishReason, // "stop" = 正常完成, "length" = 被截断
      truncated: finishReason === "length",
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    log("ERROR", "Agent chat error", { error: error.message });
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ----------------------
// v2.9.34: 使用 Function Calling + strict 模式生成结构化数据
// DeepSeek 官方支持 Tool Calls，strict 模式保证输出严格符合 schema
// ----------------------
app.post("/agent/generate-data", async (req, res) => {
  try {
    const { headers, count, description, existingItems } = req.body;
    
    if (!headers || !Array.isArray(headers) || headers.length === 0) {
      return res.status(400).json({ success: false, error: "headers 是必需的数组" });
    }
    
    const itemCount = count || 1;
    const existing = existingItems || [];
    
    log("INFO", "Function Calling (strict mode) data generation", { 
      headers, 
      count: itemCount,
      existingCount: existing.length 
    });

    // v2.9.34: 定义工具 - 使用 strict 模式
    // 所有属性必须在 required 中，且 additionalProperties 必须为 false
    const tools = [
      {
        type: "function",
        function: {
          name: "add_data_row",
          strict: true, // v2.9.34: 开启 strict 模式
          description: `添加一行数据到表格。表头是: ${headers.join(", ")}`,
          parameters: {
            type: "object",
            properties: headers.reduce((acc, h, i) => {
              acc[`col${i + 1}`] = { 
                type: "string", 
                description: `第${i + 1}列: ${h}` 
              };
              return acc;
            }, {}),
            required: headers.map((_, i) => `col${i + 1}`),
            additionalProperties: false // v2.9.34: strict 模式必需
          }
        }
      }
    ];

    const existingStr = existing.length > 0 
      ? `\n已有数据（不要重复）: ${existing.join("、")}` 
      : "";

    // v2.9.34: 使用 beta URL 以启用 strict 模式
    const BETA_API_URL = "https://api.deepseek.com/beta/chat/completions";
    
    const resp = await axios.post(
      BETA_API_URL,
      {
        model: DEEPSEEK_MODEL,
        messages: [
          { 
            role: "system", 
            content: `你是数据生成助手。用户需要生成 ${itemCount} 行表格数据。
请调用 add_data_row 函数来添加每一行数据。
每次调用 add_data_row 添加一行。
数据要真实、合理。`
          },
          { 
            role: "user", 
            content: `${description}

表头: ${headers.join(", ")}${existingStr}

请调用 add_data_row 函数生成 ${itemCount} 行数据。`
          },
        ],
        tools: tools,
        tool_choice: "auto",
        temperature: 0.3,
        max_tokens: 4000,
        stream: false,
      },
      {
        headers: {
          Authorization: "Bearer " + DEEPSEEK_API_KEY,
          "Content-Type": "application/json",
        },
        timeout: 60000,
      }
    );

    const messageObj = resp?.data?.choices?.[0]?.message || {};
    const toolCalls = messageObj.tool_calls || [];
    const finishReason = resp?.data?.choices?.[0]?.finish_reason;
    
    log("INFO", "Function Calling response", { 
      toolCallsCount: toolCalls.length,
      finishReason,
      hasContent: !!messageObj.content
    });

    // 解析 tool_calls 中的数据
    const rows = [];
    for (const toolCall of toolCalls) {
      if (toolCall.function && toolCall.function.name === "add_data_row") {
        try {
          const args = JSON.parse(toolCall.function.arguments);
          // 按列顺序提取值
          const row = headers.map((_, i) => args[`col${i + 1}`] || "");
          rows.push(row);
          log("INFO", "Parsed row", { row: row.slice(0, 3) });
        } catch (e) {
          log("WARN", "Failed to parse tool call", { error: e.message });
        }
      }
    }

    res.json({
      success: true,
      rows: rows,
      count: rows.length,
      finishReason: finishReason,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    log("ERROR", "Function Calling error", { error: error.message });
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

// ----------------------
// Routes: Chat
// ----------------------
app.post("/chat", async (req, res) => {
  try {
    const message = req.body && req.body.message ? req.body.message : "";
    const context = req.body && req.body.context ? req.body.context : {};

    if (!message) {
      return res.status(400).json({ success: false, error: "消息内容不能为空" });
    }

    log("INFO", "AI request", { message: message });

    const aiResponse = await callDeepSeekAI(message, context);
    const normalized = ensureTabularParameters(aiResponse, message);
    const excelCommand = generateExcelCommand(normalized);

    res.json({
      success: true,
      message: normalized.explanation || "已处理您的请求",
      operation: normalized.operation,
      parameters: normalized.parameters || {},
      excelCommand: excelCommand,
      confidence:
        typeof normalized.confidence === "number" ? normalized.confidence : 0.5,
      timestamp: new Date().toISOString(),
    });
  } catch (error) {
    log("ERROR", "chat error", { error: error.message });
    res.status(500).json({
      success: false,
      error: error.message,
      fallback: "已收到您的请求，但AI服务暂时不可用。请稍后重试。",
    });
  }
});

// ----------------------
// Routes: Stream Chat (SSE)
// ----------------------
app.post("/chat/stream", async (req, res) => {
  const message = req.body && req.body.message ? req.body.message : "";
  const context = req.body && req.body.context ? req.body.context : {};

  if (!message) {
    return res.status(400).json({ success: false, error: "消息内容不能为空" });
  }

  // 设置 SSE 响应头
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("X-Accel-Buffering", "no");

  // 发送 SSE 事件的辅助函数
  function sendEvent(eventType, data) {
    res.write("event: " + eventType + "\n");
    res.write("data: " + JSON.stringify(data) + "\n\n");
  }

  try {
    log("INFO", "Stream AI request", { message: message });

    // 验证 API 密钥
    const validation = await validateApiKey(DEEPSEEK_API_KEY);
    if (!validation.valid) {
      sendEvent("error", { error: "API密钥无效: " + validation.error });
      res.end();
      return;
    }

    // 发送开始事件
    sendEvent("start", { status: "processing", timestamp: new Date().toISOString() });

    // 使用增强的系统提示
    const systemPrompt = buildEnhancedSystemPrompt();

    const history =
      context && Array.isArray(context.conversationHistory) ? context.conversationHistory : [];
    const trimmedHistory = history
      .filter((item) => item && item.role && item.content)
      .slice(-6)
      .map((item) => ({ role: item.role, content: item.content }));

    const userPrompt = buildUserPrompt(message, context);

    // 使用流式API
    const streamResp = await axios.post(
      DEEPSEEK_API_URL,
      {
        model: DEEPSEEK_MODEL,
        messages: [
          { role: "system", content: systemPrompt },
          ...trimmedHistory,
          { role: "user", content: userPrompt },
        ],
        temperature: 0.2,
        max_tokens: 800,
        stream: true,
      },
      {
        headers: {
          Authorization: "Bearer " + DEEPSEEK_API_KEY,
          "Content-Type": "application/json",
        },
        timeout: 60000,
        responseType: "stream",
      }
    );

    let fullContent = "";
    let buffer = "";

    streamResp.data.on("data", function(chunk) {
      buffer += chunk.toString();
      const lines = buffer.split("\n");
      buffer = lines.pop() || "";

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line || line === "data: [DONE]") continue;
        if (!line.startsWith("data: ")) continue;

        try {
          const jsonStr = line.substring(6);
          const parsed = JSON.parse(jsonStr);
          const delta = parsed.choices && parsed.choices[0] && parsed.choices[0].delta;
          if (delta && delta.content) {
            fullContent += delta.content;
            sendEvent("chunk", { content: delta.content, accumulated: fullContent.length });
          }
        } catch (e) {
          // 忽略解析错误
        }
      }
    });

    streamResp.data.on("end", function() {
      try {
        // 清理并解析完整响应
        const cleaned = String(fullContent)
          .replace(/```json\s*/i, "")
          .replace(/```\s*$/i, "")
          .trim();

        let parsed;
        try {
          parsed = JSON.parse(cleaned);
        } catch (e) {
          parsed = {
            operation: "unknown",
            description: "AI回复",
            parameters: {},
            explanation: cleaned || "（空响应）",
            confidence: 0.3,
          };
        }

        const normalized = ensureTabularParameters(parsed, message);
        const excelCommand = generateExcelCommand(normalized);

        sendEvent("complete", {
          success: true,
          message: normalized.explanation || "已处理您的请求",
          operation: normalized.operation,
          parameters: normalized.parameters || {},
          excelCommand: excelCommand,
          confidence: typeof normalized.confidence === "number" ? normalized.confidence : 0.5,
          timestamp: new Date().toISOString(),
        });
      } catch (e) {
        sendEvent("error", { error: "解析响应失败: " + e.message });
      }
      res.end();
    });

    streamResp.data.on("error", function(err) {
      sendEvent("error", { error: "流式响应错误: " + err.message });
      res.end();
    });

  } catch (error) {
    log("ERROR", "stream chat error", { error: error.message });
    sendEvent("error", { error: error.message });
    res.end();
  }
});

app.post("/batch", async (req, res) => {
  try {
    const messages = req.body && req.body.messages ? req.body.messages : null;
    if (!Array.isArray(messages)) {
      return res.status(400).json({ success: false, error: "messages必须是数组" });
    }

    const results = [];
    for (let i = 0; i < messages.length; i++) {
      const m = messages[i];
      try {
        const aiResponse = await callDeepSeekAI(m);
        const normalized = ensureTabularParameters(aiResponse, m);
        results.push({
          message: m,
          response: normalized,
          excelCommand: generateExcelCommand(normalized),
        });
      } catch (error) {
        results.push({ message: m, success: false, error: error.message });
      }
    }

    res.json({ success: true, results: results, count: results.length, timestamp: new Date().toISOString() });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ----------------------
// Excel command generation
// ----------------------
function safeTemplateValue(value) {
  if (value === null || value === undefined) return '""';
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

function replaceAllCompat(str, search, replacement) {
  return String(str).split(search).join(replacement);
}

function generateExcelCommand(aiResponse) {
  const operation = aiResponse && aiResponse.operation ? aiResponse.operation : "";
  const parameters = aiResponse && aiResponse.parameters ? aiResponse.parameters : {};

  if (!Object.prototype.hasOwnProperty.call(EXCEL_OPERATIONS, operation)) {
    return { type: "unknown", command: "无法识别的操作", executable: false };
  }

  const opConfig = EXCEL_OPERATIONS[operation];
  let command = opConfig.template;

  Object.keys(parameters).forEach((key) => {
    const replacement = safeTemplateValue(parameters[key]);
    command = replaceAllCompat(command, "{{" + key + "}}", replacement);
  });

  return {
    type: operation,
    action: opConfig.action,
    command: command,
    parameters: parameters,
    executable: true,
  };
}

// ----------------------
// Global Error Handler
// ----------------------
app.use(function(err, req, res, _next) {
  const requestId = req.requestId || "unknown";
  
  log("ERROR", "Unhandled error", {
    requestId: requestId,
    error: err.message,
    stack: err.stack,
    path: req.path,
    method: req.method,
  });

  // 根据错误类型返回适当的状态码
  let statusCode = 500;
  let errorMessage = "服务器内部错误";

  if (err.name === "SyntaxError" && err.status === 400) {
    statusCode = 400;
    errorMessage = "请求格式错误，请检查JSON格式";
  } else if (err.code === "ECONNREFUSED") {
    statusCode = 502;
    errorMessage = "无法连接到AI服务";
  } else if (err.code === "ETIMEDOUT" || err.code === "ESOCKETTIMEDOUT") {
    statusCode = 504;
    errorMessage = "AI服务响应超时";
  }

  if (!res.headersSent) {
    res.status(statusCode).json({
      success: false,
      error: errorMessage,
      requestId: requestId,
      timestamp: new Date().toISOString(),
    });
  }
});

// ----------------------
// 404 Handler
// ----------------------
app.use(function(req, res) {
  res.status(404).json({
    success: false,
    error: "接口不存在: " + req.method + " " + req.path,
    availableEndpoints: [
      "GET  /api/health",
      "POST /chat",
      "POST /chat/stream",
      "POST /batch",
      "GET  /api/config/status",
      "POST /api/config/key",
      "DELETE /api/config/key",
    ],
  });
});

// ----------------------
// Start
// ----------------------
app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log("=========================================");
  console.log("Excel 智能助手 AI 后端服务已启动");
  console.log("地址: http://localhost:" + port);
  console.log("健康检查: GET  http://localhost:" + port + "/api/health");
  console.log("聊天接口: POST http://localhost:" + port + "/chat");
  console.log("流式聊天: POST http://localhost:" + port + "/chat/stream");
  console.log("配置密钥: POST http://localhost:" + port + "/api/config/key");
  console.log("清除密钥: DELETE http://localhost:" + port + "/api/config/key");
  console.log("密钥状态: GET  http://localhost:" + port + "/api/config/status");
  console.log("=========================================");
});

// ----------------------
// Graceful Shutdown
// ----------------------
let isShuttingDown = false;

function gracefulShutdown(signal) {
  if (isShuttingDown) return;
  isShuttingDown = true;
  
  log("INFO", "Received " + signal + ", shutting down gracefully...");
  
  // 给正在处理的请求一些时间完成
  setTimeout(function() {
    log("INFO", "Server shutdown complete");
    process.exit(0);
  }, 5000);
}

process.on("SIGTERM", function() { gracefulShutdown("SIGTERM"); });
process.on("SIGINT", function() { gracefulShutdown("SIGINT"); });

// 捕获未处理的Promise rejection
process.on("unhandledRejection", function(reason, promise) {
  log("ERROR", "Unhandled Promise Rejection", {
    reason: String(reason),
    promise: String(promise),
  });
});

// 捕获未捕获的异常
process.on("uncaughtException", function(error) {
  log("ERROR", "Uncaught Exception", {
    error: error.message,
    stack: error.stack,
  });
  // 给日志一些时间写入
  setTimeout(function() {
    process.exit(1);
  }, 1000);
});

module.exports = app;


