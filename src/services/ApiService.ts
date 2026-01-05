/**
 * API服务模块 - 统一管理所有API调用
 */

import { API_BASE, API_TIMEOUT, MAX_RETRIES, RETRY_DELAY } from "../config";

const MAX_ERROR_SNIPPET = 500;

export interface ChatRequest {
  message: string;
  systemPrompt?: string; // 自定义系统提示词 (NEW - v2.5.0 ReAct Agent)
  responseFormat?: "text" | "json"; // 响应格式 (NEW - v2.5.0 ReAct Agent)
  context?: {
    selection?: {
      address: string;
      values: (string | number | boolean | null)[][];
      rowCount: number;
      columnCount: number;
    };
    conversationHistory?: Array<{
      role: "user" | "assistant";
      content: string;
    }>;
    // 工作簿全局上下文 (NEW - v2.3.0)
    workbookContext?: {
      fileName: string;
      sheets: Array<{
        name: string;
        rowCount: number;
        columnCount: number;
        hasData: boolean;
        hasTables: boolean;
        hasCharts: boolean;
        hasPivotTables: boolean;
      }>;
      tables: Array<{
        name: string;
        sheetName: string;
        rowCount: number;
        columns: string[];
      }>;
      namedRanges: Array<{
        name: string;
        address: string;
      }>;
      charts: Array<{
        name: string;
        chartType: string;
      }>;
      totalCellsWithData: number;
      totalFormulas: number;
      overallQualityScore: number;
      issues: Array<{
        type: string;
        message: string;
        location?: string;
      }>;
    };
    // 用户偏好 (NEW - v2.3.0)
    userPreferences?: {
      defaultChartType: string;
      favoriteOperations: string[];
      lastUsedOperations: Array<{ operation: string; timestamp: number }>;
    };
  };
}

export interface ChatResponse {
  success: boolean;
  message: string;
  operation: string;
  parameters: Record<string, any>;
  excelCommand?: {
    type: string;
    action: string;
    command: string;
    parameters: Record<string, any>;
    executable: boolean;
  };
  confidence: number;
  timestamp: string;
  error?: string;
  fallback?: string;
}

export interface ApiKeyStatus {
  success: boolean;
  configured: boolean;
  isValid: boolean;
  lastUpdated: string;
  maskedKey: string | null;
  model: string;
}

export interface ApiKeyValidation {
  success: boolean;
  message?: string;
  lastUpdated?: string;
  model?: string;
  usage?: any;
  error?: string;
  code?: string | number;
}

export interface HealthStatus {
  success: boolean;
  status: string;
  service: string;
  version: string;
  environment: string;
  port: number;
  allowedOrigins: string[];
  model: string;
  configured: boolean;
}

class ApiService {
  private static instance: ApiService;
  private abortController: AbortController | null = null;

  private constructor() {}

  public static getInstance(): ApiService {
    if (!ApiService.instance) {
      ApiService.instance = new ApiService();
    }
    return ApiService.instance;
  }

  private async readJsonResponse<T>(response: Response): Promise<{
    data: T | null;
    rawText: string;
    parsed: boolean;
  }> {
    const rawText = await response.text();
    const trimmed = rawText.trim();
    if (!trimmed) {
      return { data: null, rawText: "", parsed: false };
    }

    try {
      return { data: JSON.parse(trimmed) as T, rawText: trimmed, parsed: true };
    } catch {
      return { data: null, rawText: trimmed, parsed: false };
    }
  }

  private extractErrorMessage(data: unknown): string | null {
    if (!data || typeof data !== "object") {
      return null;
    }

    const maybe =
      (data as { error?: unknown; message?: unknown }).error ??
      (data as { message?: unknown }).message;

    if (!maybe) {
      return null;
    }

    return typeof maybe === "string" ? maybe : JSON.stringify(maybe);
  }

  private formatNonJsonError(response: Response, rawText: string): string {
    const status = `${response.status} ${response.statusText}`.trim();
    if (!rawText) {
      return `Empty response from API (${status})`;
    }

    const snippet = rawText.replace(/\s+/g, " ").slice(0, MAX_ERROR_SNIPPET);
    return `Non-JSON response from API (${status}): ${snippet}`;
  }

  private formatGatewayError(response: Response, rawText: string): string | null {
    if (![502, 503, 504].includes(response.status)) {
      return null;
    }

    const status = `${response.status} ${response.statusText}`.trim();
    const snippet = rawText ? rawText.replace(/\s+/g, " ").slice(0, MAX_ERROR_SNIPPET) : "";
    const detail = snippet ? ` 详情: ${snippet}` : "";

    return `后端服务不可用或代理超时（${status}），请确认后端已启动并可访问。${detail}`;
  }

  private async parseJsonOrThrow<T>(response: Response): Promise<T> {
    const { data, rawText, parsed } = await this.readJsonResponse<T>(response);

    if (!response.ok) {
      const message =
        this.formatGatewayError(response, rawText) ??
        this.extractErrorMessage(data) ??
        this.formatNonJsonError(response, rawText);
      throw new Error(message || `HTTP error! status: ${response.status}`);
    }

    if (!parsed || data === null) {
      throw new Error(this.formatNonJsonError(response, rawText));
    }

    return data;
  }

  /**
   * 发送聊天请求到AI后端
   */
  public async sendChatRequest(request: ChatRequest): Promise<ChatResponse> {
    this.abortController = new AbortController();
    const timeoutId = setTimeout(() => {
      this.abortController?.abort();
    }, API_TIMEOUT);

    try {
      const response = await fetch(`${API_BASE}/chat`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(request),
        signal: this.abortController.signal,
      });

      clearTimeout(timeoutId);

      return await this.parseJsonOrThrow<ChatResponse>(response);
    } catch (error) {
      clearTimeout(timeoutId);

      // 如果是中止错误，重新抛出
      if (error instanceof DOMException && error.name === "AbortError") {
        throw new Error("请求超时，请稍后重试");
      }

      // 网络错误或其他错误
      throw error;
    }
  }

  /**
   * Agent 专用请求 - 支持自定义 systemPrompt 和 ReAct 模式
   */
  public async sendAgentRequest(request: {
    message: string;
    systemPrompt: string;
    responseFormat?: "text" | "json";
  }): Promise<{ success: boolean; message: string; error?: string; truncated?: boolean }> {
    this.abortController = new AbortController();
    const timeoutId = setTimeout(() => {
      this.abortController?.abort();
    }, 60000); // Agent 请求可能需要更长时间

    try {
      console.log("[ApiService] Agent request:", request.message.substring(0, 100));

      const response = await fetch(`${API_BASE}/agent/chat`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(request),
        signal: this.abortController.signal,
      });

      clearTimeout(timeoutId);

      const result = await this.parseJsonOrThrow<{
        success: boolean;
        message: string;
        error?: string;
        truncated?: boolean;
        finishReason?: string;
      }>(response);

      // v2.9.26: 检查是否截断
      if (result.truncated) {
        console.warn("[ApiService] Agent response was truncated!");
      }

      console.log("[ApiService] Agent response:", result.message?.substring(0, 200));
      return result;
    } catch (error) {
      clearTimeout(timeoutId);

      if (error instanceof DOMException && error.name === "AbortError") {
        throw new Error("Agent 请求超时，请稍后重试");
      }

      throw error;
    }
  }

  /**
   * v2.9.33: 使用 Function Calling 生成结构化数据
   * 比纯 JSON 模式更可靠，DeepSeek 官方支持
   */
  public async generateDataWithFunctionCalling(request: {
    headers: string[];
    count: number;
    description: string;
    existingItems?: string[];
  }): Promise<{ success: boolean; rows: string[][]; count: number; error?: string }> {
    this.abortController = new AbortController();
    const timeoutId = setTimeout(() => {
      this.abortController?.abort();
    }, 90000); // Function Calling 可能需要更长时间

    try {
      console.log("[ApiService] Function Calling request:", request.description.substring(0, 50));

      const response = await fetch(`${API_BASE}/agent/generate-data`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(request),
        signal: this.abortController.signal,
      });

      clearTimeout(timeoutId);

      const result = await this.parseJsonOrThrow<{
        success: boolean;
        rows: string[][];
        count: number;
        error?: string;
        finishReason?: string;
      }>(response);

      console.log("[ApiService] Function Calling result:", result.count, "rows");
      return result;
    } catch (error) {
      clearTimeout(timeoutId);

      if (error instanceof DOMException && error.name === "AbortError") {
        return { success: false, rows: [], count: 0, error: "请求超时" };
      }

      return { success: false, rows: [], count: 0, error: String(error) };
    }
  }

  /**
   * 获取API密钥状态
   */
  public async getApiKeyStatus(): Promise<ApiKeyStatus> {
    try {
      const response = await fetch(`${API_BASE}/api/config/status`, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      });

      return await this.parseJsonOrThrow<ApiKeyStatus>(response);
    } catch (error) {
      console.error("获取API密钥状态失败:", error);
      throw error;
    }
  }

  /**
   * 验证并设置API密钥
   */
  public async setApiKey(apiKey: string): Promise<ApiKeyValidation> {
    try {
      const response = await fetch(`${API_BASE}/api/config/key`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ apiKey }),
      });

      return await this.parseJsonOrThrow<ApiKeyValidation>(response);
    } catch (error) {
      console.error("设置API密钥失败:", error);
      throw error;
    }
  }

  /**
   * 清除API密钥
   */
  public async clearApiKey(): Promise<{ success: boolean; message: string }> {
    try {
      const response = await fetch(`${API_BASE}/api/config/key`, {
        method: "DELETE",
        headers: {
          "Content-Type": "application/json",
        },
      });

      return await this.parseJsonOrThrow<{ success: boolean; message: string }>(response);
    } catch (error) {
      console.error("清除API密钥失败:", error);
      throw error;
    }
  }

  /**
   * 检查后端健康状态
   */
  public async checkHealth(): Promise<HealthStatus> {
    try {
      const response = await fetch(`${API_BASE}/api/health`, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      });

      return await this.parseJsonOrThrow<HealthStatus>(response);
    } catch (error) {
      console.error("健康检查失败:", error);
      throw error;
    }
  }

  /**
   * 取消当前请求
   */
  public cancelRequest(): void {
    if (this.abortController) {
      this.abortController.abort();
      this.abortController = null;
    }
  }

  /**
   * 流式聊天事件类型
   */
  public static StreamEventTypes = {
    START: "start",
    CHUNK: "chunk",
    COMPLETE: "complete",
    ERROR: "error",
  } as const;

  /**
   * 流式聊天接口 - SSE 实现
   * @param request 聊天请求
   * @param onChunk 收到内容片段时的回调
   * @param onComplete 完成时的回调（包含完整解析结果）
   * @param onError 错误回调
   * @param onStart 开始回调
   */
  public async sendStreamingChatRequest(
    request: ChatRequest,
    callbacks: {
      onStart?: () => void;
      onChunk?: (content: string, accumulated: number) => void;
      onComplete?: (response: ChatResponse) => void;
      onError?: (error: string) => void;
    }
  ): Promise<void> {
    this.abortController = new AbortController();

    try {
      const response = await fetch(`${API_BASE}/chat/stream`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "text/event-stream",
        },
        body: JSON.stringify(request),
        signal: this.abortController.signal,
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const reader = response.body?.getReader();
      if (!reader) {
        throw new Error("无法获取响应流");
      }

      // eslint-disable-next-line no-undef
      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const line of lines) {
          if (line.startsWith("event: ")) {
            // 事件类型会在 data 行中处理
            continue;
          }

          if (line.startsWith("data: ")) {
            try {
              const dataStr = line.substring(6);
              const data = JSON.parse(dataStr);

              // 根据事件内容判断类型
              if (data.status === "processing") {
                callbacks.onStart?.();
              } else if (data.content !== undefined) {
                callbacks.onChunk?.(data.content, data.accumulated || 0);
              } else if (data.success !== undefined) {
                callbacks.onComplete?.(data as ChatResponse);
              } else if (data.error) {
                callbacks.onError?.(data.error);
              }
            } catch {
              // 忽略解析错误
            }
          }
        }
      }
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") {
        callbacks.onError?.("请求已取消");
        return;
      }
      callbacks.onError?.((error as Error).message || "流式请求失败");
    }
  }

  /**
   * 重试机制包装器
   */
  public async withRetry<T>(
    operation: () => Promise<T>,
    maxRetries: number = MAX_RETRIES,
    retryDelay: number = RETRY_DELAY
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await operation();
      } catch (error) {
        lastError = error as Error;

        if (attempt === maxRetries) {
          break;
        }

        // 等待一段时间后重试
        await new Promise((resolve) => setTimeout(resolve, retryDelay * attempt));
      }
    }

    throw lastError || new Error("操作失败");
  }
}

export default ApiService.getInstance();
