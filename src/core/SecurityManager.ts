/**
 * SecurityManager - 安全与兼容性管理
 *
 * 功能：
 * - 输入验证与清理
 * - Office.js 兼容性检测
 * - 权限控制
 * - 敏感数据处理
 * - 速率限制
 *
 * @version 1.0.0
 */

import { Logger } from "../utils/Logger";

// ============ 类型定义 ============

/**
 * Office.js API 版本要求
 */
export interface ApiRequirement {
  name: string;
  minVersion: string;
  fallback?: string;
}

/**
 * 兼容性检测结果
 */
export interface CompatibilityResult {
  supported: boolean;
  version?: string;
  missingApis: string[];
  warnings: string[];
  capabilities: Record<string, boolean>;
}

/**
 * 权限级别
 */
export enum PermissionLevel {
  READ = "read",
  WRITE = "write",
  FORMAT = "format",
  CALCULATE = "calculate",
  ADMIN = "admin",
}

/**
 * 操作权限定义
 */
export interface OperationPermission {
  operation: string;
  requiredLevel: PermissionLevel;
  description: string;
  riskLevel: "low" | "medium" | "high";
}

/**
 * 输入验证规则
 */
export interface ValidationRule {
  type: "string" | "number" | "range" | "formula" | "color" | "custom";
  maxLength?: number;
  minValue?: number;
  maxValue?: number;
  pattern?: RegExp;
  sanitize?: boolean;
  validator?: (value: unknown) => boolean;
}

/**
 * 验证结果
 */
export interface ValidationResult {
  valid: boolean;
  sanitized?: unknown;
  errors: string[];
  warnings: string[];
}

/**
 * 速率限制配置
 */
export interface RateLimitConfig {
  maxRequests: number;
  windowMs: number;
  blockDurationMs: number;
}

/**
 * 速率限制状态
 */
interface RateLimitState {
  requests: number[];
  blocked: boolean;
  blockedUntil?: number;
}

// ============ 常量 ============

/**
 * Office.js API 要求集
 */
const REQUIRED_APIS: ApiRequirement[] = [
  { name: "ExcelApi", minVersion: "1.1", fallback: undefined },
  { name: "ExcelApi", minVersion: "1.4", fallback: "basic-formatting" },
  { name: "ExcelApi", minVersion: "1.9", fallback: "legacy-charts" },
];

/**
 * 操作权限映射
 */
const OPERATION_PERMISSIONS: OperationPermission[] = [
  {
    operation: "excel_read_range",
    requiredLevel: PermissionLevel.READ,
    description: "读取单元格数据",
    riskLevel: "low",
  },
  {
    operation: "excel_write_range",
    requiredLevel: PermissionLevel.WRITE,
    description: "写入单元格数据",
    riskLevel: "medium",
  },
  {
    operation: "excel_format_range",
    requiredLevel: PermissionLevel.FORMAT,
    description: "格式化单元格",
    riskLevel: "low",
  },
  {
    operation: "excel_delete_range",
    requiredLevel: PermissionLevel.WRITE,
    description: "删除单元格",
    riskLevel: "high",
  },
  {
    operation: "excel_add_formula",
    requiredLevel: PermissionLevel.CALCULATE,
    description: "添加公式",
    riskLevel: "medium",
  },
  {
    operation: "excel_create_chart",
    requiredLevel: PermissionLevel.FORMAT,
    description: "创建图表",
    riskLevel: "low",
  },
  {
    operation: "excel_clear_worksheet",
    requiredLevel: PermissionLevel.ADMIN,
    description: "清空工作表",
    riskLevel: "high",
  },
];

/**
 * 危险模式列表
 */
const DANGEROUS_PATTERNS = [
  // 脚本注入
  /<script[\s\S]*?>[\s\S]*?<\/script>/gi,
  /javascript:/gi,
  /on\w+\s*=/gi,
  // 公式注入
  /^=\s*cmd\s*\|/i,
  /^=\s*system\s*\(/i,
  // SQL 注入模式
  /('|")\s*(or|and)\s*('|")/gi,
  /;\s*(drop|delete|truncate|update)\s+/gi,
  // 路径遍历
  /\.\.[/\\]/g,
];

// ============ SecurityManager 类 ============

class SecurityManagerImpl {
  private rateLimitStates: Map<string, RateLimitState> = new Map();
  private rateLimitConfig: RateLimitConfig = {
    maxRequests: 100,
    windowMs: 60000, // 1分钟
    blockDurationMs: 300000, // 5分钟
  };
  private userPermissions: Set<PermissionLevel> = new Set([
    PermissionLevel.READ,
    PermissionLevel.WRITE,
    PermissionLevel.FORMAT,
    PermissionLevel.CALCULATE,
  ]);

  // ============ 兼容性检测 ============

  /**
   * 检测 Office.js 兼容性
   */
  checkCompatibility(): CompatibilityResult {
    const result: CompatibilityResult = {
      supported: true,
      missingApis: [],
      warnings: [],
      capabilities: {},
    };

    // 检测 Office 是否可用
    if (typeof Office === "undefined") {
      result.supported = false;
      result.missingApis.push("Office");
      Logger.error("[SecurityManager] Office.js 不可用");
      return result;
    }

    // 检测 Excel 是否可用
    if (typeof Excel === "undefined") {
      result.supported = false;
      result.missingApis.push("Excel");
      Logger.error("[SecurityManager] Excel API 不可用");
      return result;
    }

    // 检测各 API 版本
    for (const req of REQUIRED_APIS) {
      const isSupported = this.checkApiVersion(req.name, req.minVersion);
      const capabilityKey = `${req.name}_${req.minVersion.replace(".", "_")}`;
      result.capabilities[capabilityKey] = isSupported;

      if (!isSupported) {
        if (req.fallback) {
          result.warnings.push(`${req.name} ${req.minVersion} 不可用，将使用 ${req.fallback} 模式`);
        } else {
          result.missingApis.push(`${req.name} ${req.minVersion}`);
          result.supported = false;
        }
      }
    }

    // 检测特定功能
    result.capabilities["conditional_formatting"] = this.checkApiVersion("ExcelApi", "1.6");
    result.capabilities["charts"] = this.checkApiVersion("ExcelApi", "1.1");
    result.capabilities["tables"] = this.checkApiVersion("ExcelApi", "1.1");
    result.capabilities["pivot_tables"] = this.checkApiVersion("ExcelApi", "1.8");
    result.capabilities["custom_functions"] = this.checkApiVersion("CustomFunctionsRuntime", "1.1");

    Logger.info("[SecurityManager] 兼容性检测完成", result);
    return result;
  }

  /**
   * 检测 API 版本
   */
  private checkApiVersion(setName: string, version: string): boolean {
    try {
      return Office.context.requirements.isSetSupported(setName, version);
    } catch {
      return false;
    }
  }

  /**
   * 获取推荐的 API 使用方式
   */
  getRecommendedApi(feature: string): string {
    const compatibility = this.checkCompatibility();

    const fallbacks: Record<string, Record<string, string>> = {
      formatting: {
        default: "Range.format",
        fallback: "Range.numberFormat + Range.values",
      },
      charts: {
        default: "Sheet.charts.add",
        fallback: "手动数据导出",
      },
      conditionalFormatting: {
        default: "Range.conditionalFormats.add",
        fallback: "Range.format（静态样式）",
      },
    };

    const featureFallbacks = fallbacks[feature];
    if (!featureFallbacks) {
      return "default";
    }

    const isAdvancedSupported = compatibility.capabilities[`ExcelApi_1_6`];
    return isAdvancedSupported ? featureFallbacks.default : featureFallbacks.fallback;
  }

  // ============ 输入验证 ============

  /**
   * 验证输入
   */
  validateInput(value: unknown, rules: ValidationRule | ValidationRule[]): ValidationResult {
    const ruleArray = Array.isArray(rules) ? rules : [rules];
    const result: ValidationResult = {
      valid: true,
      errors: [],
      warnings: [],
    };

    let sanitized = value;

    for (const rule of ruleArray) {
      const ruleResult = this.applyValidationRule(sanitized, rule);

      if (!ruleResult.valid) {
        result.valid = false;
        result.errors.push(...ruleResult.errors);
      }

      result.warnings.push(...ruleResult.warnings);

      if (ruleResult.sanitized !== undefined) {
        sanitized = ruleResult.sanitized;
      }
    }

    result.sanitized = sanitized;
    return result;
  }

  /**
   * 应用单个验证规则
   */
  private applyValidationRule(value: unknown, rule: ValidationRule): ValidationResult {
    const result: ValidationResult = {
      valid: true,
      errors: [],
      warnings: [],
    };

    switch (rule.type) {
      case "string":
        result.valid = this.validateString(value, rule, result);
        break;
      case "number":
        result.valid = this.validateNumber(value, rule, result);
        break;
      case "range":
        result.valid = this.validateRange(value, result);
        break;
      case "formula":
        result.valid = this.validateFormula(value, result);
        break;
      case "color":
        result.valid = this.validateColor(value, result);
        break;
      case "custom":
        if (rule.validator) {
          result.valid = rule.validator(value);
          if (!result.valid) {
            result.errors.push("自定义验证失败");
          }
        }
        break;
    }

    // 清理危险内容
    if (rule.sanitize && typeof value === "string") {
      result.sanitized = this.sanitizeInput(value);
    }

    return result;
  }

  /**
   * 验证字符串
   */
  private validateString(value: unknown, rule: ValidationRule, result: ValidationResult): boolean {
    if (typeof value !== "string") {
      result.errors.push("期望字符串类型");
      return false;
    }

    if (rule.maxLength && value.length > rule.maxLength) {
      result.errors.push(`字符串长度超过限制 (${rule.maxLength})`);
      return false;
    }

    if (rule.pattern && !rule.pattern.test(value)) {
      result.errors.push("字符串格式不匹配");
      return false;
    }

    return true;
  }

  /**
   * 验证数字
   */
  private validateNumber(value: unknown, rule: ValidationRule, result: ValidationResult): boolean {
    const num = typeof value === "number" ? value : parseFloat(String(value));

    if (isNaN(num)) {
      result.errors.push("无效的数字");
      return false;
    }

    if (rule.minValue !== undefined && num < rule.minValue) {
      result.errors.push(`数值小于最小值 (${rule.minValue})`);
      return false;
    }

    if (rule.maxValue !== undefined && num > rule.maxValue) {
      result.errors.push(`数值大于最大值 (${rule.maxValue})`);
      return false;
    }

    return true;
  }

  /**
   * 验证单元格范围
   */
  private validateRange(value: unknown, result: ValidationResult): boolean {
    if (typeof value !== "string") {
      result.errors.push("范围必须是字符串");
      return false;
    }

    // 支持的格式: A1, A1:B10, Sheet1!A1:B10
    const rangePattern = /^(?:'?[^'!]+['!])?[A-Z]{1,3}\d{1,7}(?::[A-Z]{1,3}\d{1,7})?$/i;

    if (!rangePattern.test(value)) {
      result.errors.push("无效的单元格范围格式");
      return false;
    }

    // 检查范围大小
    const match = value.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
    if (match) {
      const [, startCol, startRow, endCol, endRow] = match;
      const colDiff = this.columnToNumber(endCol) - this.columnToNumber(startCol) + 1;
      const rowDiff = parseInt(endRow) - parseInt(startRow) + 1;
      const cellCount = colDiff * rowDiff;

      if (cellCount > 100000) {
        result.warnings.push(`范围包含 ${cellCount} 个单元格，可能影响性能`);
      }
    }

    return true;
  }

  /**
   * 列名转数字
   */
  private columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * 验证公式
   */
  private validateFormula(value: unknown, result: ValidationResult): boolean {
    if (typeof value !== "string") {
      result.errors.push("公式必须是字符串");
      return false;
    }

    // 必须以 = 开头
    if (!value.startsWith("=")) {
      result.errors.push("公式必须以 = 开头");
      return false;
    }

    // 检查危险公式
    const dangerousFormulas = [
      /^=\s*WEBSERVICE/i,
      /^=\s*FILTERXML/i,
      /^=\s*DDEAUTO/i,
      /^=\s*cmd\s*\|/i,
    ];

    for (const pattern of dangerousFormulas) {
      if (pattern.test(value)) {
        result.errors.push("检测到潜在危险公式");
        return false;
      }
    }

    // 检查括号匹配
    let depth = 0;
    for (const char of value) {
      if (char === "(") depth++;
      if (char === ")") depth--;
      if (depth < 0) {
        result.errors.push("公式括号不匹配");
        return false;
      }
    }

    if (depth !== 0) {
      result.errors.push("公式括号不匹配");
      return false;
    }

    return true;
  }

  /**
   * 验证颜色
   */
  private validateColor(value: unknown, result: ValidationResult): boolean {
    if (typeof value !== "string") {
      result.errors.push("颜色必须是字符串");
      return false;
    }

    // 支持格式: #RGB, #RRGGBB, rgb(), rgba(), 颜色名称
    const colorPatterns = [
      /^#[0-9A-Fa-f]{3}$/,
      /^#[0-9A-Fa-f]{6}$/,
      /^rgb\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*\)$/i,
      /^rgba\(\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\s*,\s*[\d.]+\s*\)$/i,
    ];

    const namedColors = [
      "red",
      "green",
      "blue",
      "yellow",
      "orange",
      "purple",
      "black",
      "white",
      "gray",
      "grey",
      "pink",
      "brown",
    ];

    const isValid =
      colorPatterns.some((p) => p.test(value)) || namedColors.includes(value.toLowerCase());

    if (!isValid) {
      result.errors.push("无效的颜色格式");
      return false;
    }

    return true;
  }

  /**
   * 清理输入
   */
  sanitizeInput(input: string): string {
    let sanitized = input;

    // 移除危险模式
    for (const pattern of DANGEROUS_PATTERNS) {
      sanitized = sanitized.replace(pattern, "");
    }

    // HTML 实体编码
    sanitized = sanitized
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");

    // 移除控制字符
    // eslint-disable-next-line no-control-regex
    sanitized = sanitized.replace(/[\x00-\x1F\x7F]/g, "");

    // 限制长度
    const maxLength = 10000;
    if (sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength);
    }

    return sanitized;
  }

  /**
   * 检测敏感数据
   */
  detectSensitiveData(data: string): Array<{ type: string; masked: string }> {
    const detections: Array<{ type: string; masked: string }> = [];

    const patterns: Array<{ type: string; pattern: RegExp; maskFn: (s: string) => string }> = [
      // 邮箱
      {
        type: "email",
        pattern: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g,
        maskFn: (s) => s.replace(/^(.{2}).*(@.*)$/, "$1***$2"),
      },
      // 手机号
      {
        type: "phone",
        pattern: /1[3-9]\d{9}/g,
        maskFn: (s) => s.replace(/^(.{3}).*(.{4})$/, "$1****$2"),
      },
      // 身份证号
      {
        type: "id_card",
        pattern: /\d{17}[\dXx]/g,
        maskFn: (s) => s.replace(/^(.{4}).*(.{4})$/, "$1**********$2"),
      },
      // 银行卡号
      {
        type: "bank_card",
        pattern: /\d{16,19}/g,
        maskFn: (s) => s.replace(/^(.{4}).*(.{4})$/, "$1********$2"),
      },
      // API 密钥（常见格式）
      {
        type: "api_key",
        pattern: /(?:sk|pk|api[_-]?key)[_-]?[a-zA-Z0-9]{20,}/gi,
        maskFn: (s) => s.substring(0, 8) + "********",
      },
    ];

    for (const { type, pattern, maskFn } of patterns) {
      const matches = data.match(pattern);
      if (matches) {
        for (const match of matches) {
          detections.push({
            type,
            masked: maskFn(match),
          });
        }
      }
    }

    return detections;
  }

  /**
   * 掩码敏感数据
   */
  maskSensitiveData(data: string): string {
    let masked = data;

    // 掩码邮箱
    masked = masked.replace(
      /([a-zA-Z0-9._%+-]{2})[a-zA-Z0-9._%+-]*(@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g,
      "$1***$2"
    );

    // 掩码手机号
    masked = masked.replace(/(1[3-9]\d{2})\d{4}(\d{4})/g, "$1****$2");

    // 掩码身份证
    masked = masked.replace(/(\d{4})\d{10}(\d{3}[\dXx])/g, "$1**********$2");

    return masked;
  }

  // ============ 权限控制 ============

  /**
   * 检查操作权限
   */
  checkPermission(operation: string): { allowed: boolean; reason?: string } {
    const opPermission = OPERATION_PERMISSIONS.find((p) => p.operation === operation);

    if (!opPermission) {
      // 未定义的操作，默认允许（但记录警告）
      Logger.warn("[SecurityManager] 未定义权限的操作", { operation });
      return { allowed: true };
    }

    if (!this.userPermissions.has(opPermission.requiredLevel)) {
      return {
        allowed: false,
        reason: `需要 ${opPermission.requiredLevel} 权限执行 "${opPermission.description}"`,
      };
    }

    // 高风险操作需要额外确认
    if (opPermission.riskLevel === "high") {
      Logger.warn("[SecurityManager] 高风险操作", {
        operation,
        description: opPermission.description,
      });
    }

    return { allowed: true };
  }

  /**
   * 设置用户权限
   */
  setPermissions(permissions: PermissionLevel[]): void {
    this.userPermissions = new Set(permissions);
    Logger.info("[SecurityManager] 更新用户权限", { permissions });
  }

  /**
   * 获取用户权限
   */
  getPermissions(): PermissionLevel[] {
    return Array.from(this.userPermissions);
  }

  /**
   * 添加权限
   */
  addPermission(permission: PermissionLevel): void {
    this.userPermissions.add(permission);
  }

  /**
   * 移除权限
   */
  removePermission(permission: PermissionLevel): void {
    this.userPermissions.delete(permission);
  }

  // ============ 速率限制 ============

  /**
   * 检查速率限制
   */
  checkRateLimit(key: string = "global"): { allowed: boolean; retryAfter?: number } {
    const now = Date.now();
    let state = this.rateLimitStates.get(key);

    if (!state) {
      state = { requests: [], blocked: false };
      this.rateLimitStates.set(key, state);
    }

    // 检查是否被阻塞
    if (state.blocked && state.blockedUntil) {
      if (now < state.blockedUntil) {
        return {
          allowed: false,
          retryAfter: Math.ceil((state.blockedUntil - now) / 1000),
        };
      }
      // 阻塞期结束，重置状态
      state.blocked = false;
      state.requests = [];
    }

    // 清理过期的请求记录
    state.requests = state.requests.filter((t) => now - t < this.rateLimitConfig.windowMs);

    // 检查是否超过限制
    if (state.requests.length >= this.rateLimitConfig.maxRequests) {
      state.blocked = true;
      state.blockedUntil = now + this.rateLimitConfig.blockDurationMs;
      Logger.warn("[SecurityManager] 触发速率限制", { key });
      return {
        allowed: false,
        retryAfter: Math.ceil(this.rateLimitConfig.blockDurationMs / 1000),
      };
    }

    // 记录请求
    state.requests.push(now);
    return { allowed: true };
  }

  /**
   * 设置速率限制配置
   */
  setRateLimitConfig(config: Partial<RateLimitConfig>): void {
    this.rateLimitConfig = { ...this.rateLimitConfig, ...config };
  }

  /**
   * 重置速率限制
   */
  resetRateLimit(key?: string): void {
    if (key) {
      this.rateLimitStates.delete(key);
    } else {
      this.rateLimitStates.clear();
    }
  }

  // ============ 审计日志 ============

  /**
   * 记录安全事件
   */
  logSecurityEvent(
    eventType: string,
    details: Record<string, unknown>,
    severity: "info" | "warning" | "error" = "info"
  ): void {
    const event = {
      type: eventType,
      timestamp: new Date().toISOString(),
      severity,
      details: this.maskSensitiveData(JSON.stringify(details)),
    };

    switch (severity) {
      case "error":
        Logger.error("[SecurityAudit]", event);
        break;
      case "warning":
        Logger.warn("[SecurityAudit]", event);
        break;
      default:
        Logger.info("[SecurityAudit]", event);
    }

    // 可选：发送到审计服务
    this.sendToAuditService(event);
  }

  /**
   * 发送到审计服务（占位）
   */
  private sendToAuditService(_event: unknown): void {
    // 在生产环境中，这里可以发送到集中的审计日志服务
  }

  // ============ 工具方法 ============

  /**
   * 生成安全的随机 ID
   */
  generateSecureId(length: number = 16): string {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "";

    // 使用 crypto API 如果可用
    if (typeof crypto !== "undefined" && crypto.getRandomValues) {
      const array = new Uint8Array(length);
      crypto.getRandomValues(array);
      for (let i = 0; i < length; i++) {
        result += chars[array[i] % chars.length];
      }
    } else {
      for (let i = 0; i < length; i++) {
        result += chars[Math.floor(Math.random() * chars.length)];
      }
    }

    return result;
  }

  /**
   * 验证 URL
   */
  validateUrl(url: string, allowedHosts?: string[]): boolean {
    try {
      const parsed = new URL(url);

      // 只允许 HTTP/HTTPS
      if (!["http:", "https:"].includes(parsed.protocol)) {
        return false;
      }

      // 检查允许的主机
      if (allowedHosts && allowedHosts.length > 0) {
        return allowedHosts.some(
          (host) => parsed.hostname === host || parsed.hostname.endsWith(`.${host}`)
        );
      }

      return true;
    } catch {
      return false;
    }
  }

  /**
   * 重置所有安全状态
   */
  reset(): void {
    this.rateLimitStates.clear();
    this.userPermissions = new Set([
      PermissionLevel.READ,
      PermissionLevel.WRITE,
      PermissionLevel.FORMAT,
      PermissionLevel.CALCULATE,
    ]);
  }
}

// ============ 单例导出 ============

export const SecurityManager = new SecurityManagerImpl();

// 便捷方法导出
export const security = {
  checkCompatibility: () => SecurityManager.checkCompatibility(),
  validateInput: (value: unknown, rules: ValidationRule | ValidationRule[]) =>
    SecurityManager.validateInput(value, rules),
  sanitizeInput: (input: string) => SecurityManager.sanitizeInput(input),
  checkPermission: (operation: string) => SecurityManager.checkPermission(operation),
  checkRateLimit: (key?: string) => SecurityManager.checkRateLimit(key),
  detectSensitiveData: (data: string) => SecurityManager.detectSensitiveData(data),
  maskSensitiveData: (data: string) => SecurityManager.maskSensitiveData(data),
  generateSecureId: (length?: number) => SecurityManager.generateSecureId(length),
};
