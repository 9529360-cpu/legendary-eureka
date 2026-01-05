/**
 * Backend config
 * IMPORTANT:
 * - In Office Add-in (https://localhost:3000), calling http://localhost:3001 often triggers mixed content / CORS issues.
 * - We use same-origin paths and rely on webpack devServer proxy.
 */

import { TIMEOUTS, RETRY, EXCEL_LIMITS as LIMITS, LOGGING } from "./config/constants";

// API Configuration
export const AI_BACKEND_URL = "http://localhost:3001";
export const API_BASE = ""; // same-origin
export const API_TIMEOUT = TIMEOUTS.API_DEFAULT;
export const MAX_RETRIES = RETRY.MAX_ATTEMPTS;
export const RETRY_DELAY = TIMEOUTS.RETRY_BASE_DELAY;

// Feature Flags
export const FEATURES = {
  AI_INTEGRATION: true,
  EXCEL_OPERATIONS: true,
  REAL_TIME_UPDATES: true,
  ERROR_LOGGING: true,
  DEBUG_MODE: process.env.NODE_ENV === "development",
};

// Error Configuration
export const ERROR_CONFIG = {
  SHOW_USER_FRIENDLY_MESSAGES: true,
  LOG_TO_CONSOLE: true,
  MAX_ERROR_LENGTH: LOGGING.DATA_TRUNCATE_LENGTH,
  SUPPRESSED_ERRORS: ["CORS", "NetworkError", "AbortError"],
};

// Excel Operation Limits
export const EXCEL_LIMITS = {
  MAX_ROWS_PER_OPERATION: LIMITS.MAX_ROWS_PER_OPERATION,
  MAX_COLUMNS_PER_OPERATION: LIMITS.MAX_COLUMNS_PER_OPERATION,
  MAX_CELLS_PER_BATCH: LIMITS.MAX_CELLS_PER_BATCH,
  TIMEOUT_MS: TIMEOUTS.EXCEL_OPERATION,
};

// 统一配置导出
export const CONFIG = {
  debug: FEATURES.DEBUG_MODE,
  apiTimeout: API_TIMEOUT,
  maxRetries: MAX_RETRIES,
  retryDelay: RETRY_DELAY,
};

// Use these in fetch():
//   fetch(`${API_BASE}/chat`)
//   fetch(`${API_BASE}/api/config/status`)

export default {
  AI_BACKEND_URL,
  API_BASE,
  API_TIMEOUT,
  MAX_RETRIES,
  RETRY_DELAY,
  FEATURES,
  ERROR_CONFIG,
  EXCEL_LIMITS,
  CONFIG,
};
