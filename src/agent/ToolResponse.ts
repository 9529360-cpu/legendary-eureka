/**
 * ToolResponse - 统一工具响应格式 v1.0
 *
 * 借鉴自 sv-excel-agent errors.py
 *
 * 提供统一的成功/错误响应格式，便于 LLM 理解工具执行结果。
 *
 * @see https://github.com/Sylvian/sv-excel-agent
 */

// ==================== 类型定义 ====================

/**
 * 错误代码枚举
 */
export const ErrorCodes = {
  // 通用错误
  UNKNOWN_ERROR: "UNKNOWN_ERROR",
  TIMEOUT: "TIMEOUT",
  VALIDATION_ERROR: "VALIDATION_ERROR",

  // Excel 相关
  SHEET_NOT_FOUND: "SHEET_NOT_FOUND",
  RANGE_INVALID: "RANGE_INVALID",
  CELL_NOT_FOUND: "CELL_NOT_FOUND",
  FORMULA_ERROR: "FORMULA_ERROR",
  PERMISSION_DENIED: "PERMISSION_DENIED",

  // 工具相关
  TOOL_NOT_FOUND: "TOOL_NOT_FOUND",
  TOOL_EXECUTION_FAILED: "TOOL_EXECUTION_FAILED",
  PARAMETER_MISSING: "PARAMETER_MISSING",
  PARAMETER_INVALID: "PARAMETER_INVALID",

  // 审批相关
  APPROVAL_REQUIRED: "APPROVAL_REQUIRED",
  APPROVAL_REJECTED: "APPROVAL_REJECTED",
  APPROVAL_TIMEOUT: "APPROVAL_TIMEOUT",
} as const;

export type ErrorCode = (typeof ErrorCodes)[keyof typeof ErrorCodes];

/**
 * 响应状态
 */
export type ResponseStatus = "success" | "error";

/**
 * 基础响应接口
 */
export interface BaseResponse {
  status: ResponseStatus;
}

/**
 * 成功响应
 */
export interface SuccessResponse<T = unknown> extends BaseResponse {
  status: "success";
  data?: T;
  message?: string;
  cells?: CellResult[];
  metadata?: Record<string, unknown>;
}

/**
 * 错误响应
 */
export interface ErrorResponse extends BaseResponse {
  status: "error";
  error: string;
  code?: ErrorCode;
  details?: Record<string, unknown>;
  cells?: CellResult[];
}

/**
 * 单元格结果
 */
export interface CellResult {
  cell: string; // 单元格地址，如 "A1"
  value: unknown; // 设置的值
  calculated?: unknown; // 公式计算结果
  error?: string; // 错误信息
}

// ==================== ToolSuccess 类 ====================

/**
 * 工具成功响应构建器
 *
 * @example
 * ```typescript
 * // 简单成功
 * return new ToolSuccess().toJSON();
 * // {"status": "success"}
 *
 * // 带消息
 * return new ToolSuccess({ message: "已创建工作表" }).toJSON();
 * // {"status": "success", "message": "已创建工作表"}
 *
 * // 带数据
 * return new ToolSuccess({ data: { sheets: ["Sheet1", "Sheet2"] } }).toJSON();
 * // {"status": "success", "data": {"sheets": ["Sheet1", "Sheet2"]}}
 *
 * // 带单元格结果
 * return new ToolSuccess({ cells: [{cell: "A1", value: 100, calculated: 100}] }).toJSON();
 * ```
 */
export class ToolSuccess<T = unknown> {
  private response: SuccessResponse<T>;

  constructor(options: Omit<SuccessResponse<T>, "status"> = {}) {
    this.response = {
      status: "success",
      ...options,
    };
  }

  /**
   * 添加数据
   */
  withData(data: T): ToolSuccess<T> {
    this.response.data = data;
    return this;
  }

  /**
   * 添加消息
   */
  withMessage(message: string): ToolSuccess<T> {
    this.response.message = message;
    return this;
  }

  /**
   * 添加单元格结果
   */
  withCells(cells: CellResult[]): ToolSuccess<T> {
    this.response.cells = cells;
    return this;
  }

  /**
   * 添加元数据
   */
  withMetadata(metadata: Record<string, unknown>): ToolSuccess<T> {
    this.response.metadata = metadata;
    return this;
  }

  /**
   * 转换为 JSON 字符串
   */
  toJSON(): string {
    return JSON.stringify(this.response);
  }

  /**
   * 获取响应对象
   */
  toObject(): SuccessResponse<T> {
    return { ...this.response };
  }
}

// ==================== ToolError 类 ====================

/**
 * 工具错误响应构建器
 *
 * @example
 * ```typescript
 * // 简单错误
 * return new ToolError("工作表不存在").toJSON();
 * // {"status": "error", "error": "工作表不存在"}
 *
 * // 带错误码
 * return new ToolError("Sheet1 不存在", ErrorCodes.SHEET_NOT_FOUND).toJSON();
 * // {"status": "error", "error": "Sheet1 不存在", "code": "SHEET_NOT_FOUND"}
 *
 * // 带详情
 * return new ToolError("公式错误", ErrorCodes.FORMULA_ERROR)
 *   .withDetails({ formula: "=A1+", position: "B2" })
 *   .toJSON();
 *
 * // 带失败单元格
 * return new ToolError("公式计算失败")
 *   .withCells([{cell: "A1", value: "=1/0", error: "#DIV/0!"}])
 *   .toJSON();
 * ```
 */
export class ToolError extends Error {
  private response: ErrorResponse;

  constructor(message: string, code?: ErrorCode) {
    super(message);
    this.name = "ToolError";
    this.response = {
      status: "error",
      error: message,
    };
    if (code) {
      this.response.code = code;
    }
  }

  /**
   * 添加错误码
   */
  withCode(code: ErrorCode): ToolError {
    this.response.code = code;
    return this;
  }

  /**
   * 添加详细信息
   */
  withDetails(details: Record<string, unknown>): ToolError {
    this.response.details = details;
    return this;
  }

  /**
   * 添加失败的单元格
   */
  withCells(cells: CellResult[]): ToolError {
    this.response.cells = cells;
    return this;
  }

  /**
   * 转换为 JSON 字符串
   */
  toJSON(): string {
    return JSON.stringify(this.response);
  }

  /**
   * 获取响应对象
   */
  toObject(): ErrorResponse {
    return { ...this.response };
  }

  /**
   * 从 Error 创建 ToolError
   */
  static fromError(err: Error, code?: ErrorCode): ToolError {
    return new ToolError(err.message, code ?? ErrorCodes.UNKNOWN_ERROR);
  }

  /**
   * 从未知值创建 ToolError
   */
  static fromUnknown(err: unknown, code?: ErrorCode): ToolError {
    if (err instanceof ToolError) {
      return err;
    }
    if (err instanceof Error) {
      return ToolError.fromError(err, code);
    }
    return new ToolError(String(err), code ?? ErrorCodes.UNKNOWN_ERROR);
  }
}

// ==================== 辅助函数 ====================

/**
 * 创建成功响应
 */
export function success<T = unknown>(options?: Omit<SuccessResponse<T>, "status">): ToolSuccess<T> {
  return new ToolSuccess(options);
}

/**
 * 创建错误响应
 */
export function error(message: string, code?: ErrorCode): ToolError {
  return new ToolError(message, code);
}

/**
 * 检查响应是否成功
 */
export function isSuccess(response: BaseResponse): response is SuccessResponse {
  return response.status === "success";
}

/**
 * 检查响应是否失败
 */
export function isError(response: BaseResponse): response is ErrorResponse {
  return response.status === "error";
}

/**
 * 解析 JSON 响应
 */
export function parseResponse(json: string): SuccessResponse | ErrorResponse {
  try {
    const parsed = JSON.parse(json);
    if (parsed.status === "success" || parsed.status === "error") {
      return parsed;
    }
    // 没有 status 字段，包装为成功
    return { status: "success", data: parsed };
  } catch {
    return { status: "error", error: "Invalid JSON response", code: ErrorCodes.UNKNOWN_ERROR };
  }
}

/**
 * 检查响应中是否有公式错误
 */
export function hasFormulaErrors(cells?: CellResult[]): boolean {
  if (!cells) return false;
  return cells.some((cell) => {
    if (cell.error) return true;
    const calc = String(cell.calculated ?? "");
    return calc.startsWith("#") && calc.endsWith("!");
  });
}

/**
 * 创建单元格结果
 */
export function cellResult(
  cell: string,
  value: unknown,
  calculated?: unknown,
  error?: string
): CellResult {
  const result: CellResult = { cell, value };
  if (calculated !== undefined) result.calculated = calculated;
  if (error) result.error = error;
  return result;
}

// ==================== 导出 ====================

export default {
  ToolSuccess,
  ToolError,
  ErrorCodes,
  success,
  error,
  isSuccess,
  isError,
  parseResponse,
  hasFormulaErrors,
  cellResult,
};
