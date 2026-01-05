/**
 * useApiSettings Hook
 * v2.9.8: 封装后端连接状态和 API 密钥管理
 *
 * 边界原则：
 * - App.tsx 不应直接调用 ApiService 的设置方法
 * - 所有后端健康检查、API密钥操作通过此 hook 进行
 */

import { useState, useCallback, useRef } from "react";
import ApiService, { ApiKeyStatus } from "../../services/ApiService";

export interface ApiSettingsState {
  /** 后端是否健康 */
  backendHealthy: boolean | null;
  /** 后端检查中 */
  backendChecking: boolean;
  /** 后端错误信息 */
  backendError: string | null;
  /** API 密钥状态 */
  apiKeyStatus: ApiKeyStatus | null;
  /** API 密钥操作中 */
  apiKeyBusy: boolean;
}

export interface UseApiSettingsReturn extends ApiSettingsState {
  /** 检查后端健康状态 */
  refreshBackendStatus: (showError?: boolean) => Promise<boolean>;
  /** 检查 API 密钥状态 */
  checkApiKeyStatus: () => Promise<ApiKeyStatus | null>;
  /** 设置 API 密钥 */
  setApiKey: (key: string) => Promise<{ success: boolean; message?: string }>;
  /** 清除 API 密钥 */
  clearApiKey: () => Promise<{ success: boolean; message?: string }>;
  /** 初始化（检查后端和密钥状态） */
  bootstrap: () => Promise<void>;
}

export function useApiSettings(): UseApiSettingsReturn {
  const [backendHealthy, setBackendHealthy] = useState<boolean | null>(null);
  const [backendChecking, setBackendChecking] = useState(false);
  const [backendError, setBackendError] = useState<string | null>(null);
  const [apiKeyStatus, setApiKeyStatus] = useState<ApiKeyStatus | null>(null);
  const [apiKeyBusy, setApiKeyBusy] = useState(false);

  // 防止并发检查
  const checkingRef = useRef(false);

  const refreshBackendStatus = useCallback(
    async (showError = false): Promise<boolean> => {
      if (checkingRef.current) {
        return backendHealthy ?? false;
      }

      checkingRef.current = true;
      setBackendChecking(true);

      try {
        await ApiService.checkHealth();
        setBackendHealthy(true);
        setBackendError(null);
        return true;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        setBackendHealthy(false);
        setBackendError(message);
        setApiKeyStatus(null);
        if (showError) {
          console.warn("后端服务不可用:", message);
        }
        return false;
      } finally {
        setBackendChecking(false);
        checkingRef.current = false;
      }
    },
    [backendHealthy]
  );

  const checkApiKeyStatus = useCallback(async (): Promise<ApiKeyStatus | null> => {
    if (backendHealthy === false) {
      return null;
    }

    try {
      const status = await ApiService.getApiKeyStatus();
      setApiKeyStatus(status);
      return status;
    } catch (error) {
      console.error("检查API密钥状态失败:", error);
      setApiKeyStatus(null);
      return null;
    }
  }, [backendHealthy]);

  const setApiKey = useCallback(
    async (key: string): Promise<{ success: boolean; message?: string }> => {
      if (!key.trim() || apiKeyBusy) {
        return { success: false, message: "密钥为空或正在处理中" };
      }

      // 先确保后端可用
      const backendOk = backendHealthy === true ? true : await refreshBackendStatus(true);
      if (!backendOk) {
        return { success: false, message: "后端服务不可用" };
      }

      setApiKeyBusy(true);
      try {
        const result = await ApiService.setApiKey(key.trim());
        if (result.success) {
          await checkApiKeyStatus();
        }
        return result;
      } catch (error) {
        const message = error instanceof Error ? error.message : "设置API密钥时发生错误";
        return { success: false, message };
      } finally {
        setApiKeyBusy(false);
      }
    },
    [apiKeyBusy, backendHealthy, refreshBackendStatus, checkApiKeyStatus]
  );

  const clearApiKey = useCallback(async (): Promise<{ success: boolean; message?: string }> => {
    const backendOk = backendHealthy === true ? true : await refreshBackendStatus(true);
    if (!backendOk) {
      return { success: false, message: "后端服务不可用" };
    }

    setApiKeyBusy(true);
    try {
      const result = await ApiService.clearApiKey();
      if (result.success) {
        await checkApiKeyStatus();
      }
      return result;
    } catch (error) {
      const message = error instanceof Error ? error.message : "清除API密钥时发生错误";
      return { success: false, message };
    } finally {
      setApiKeyBusy(false);
    }
  }, [backendHealthy, refreshBackendStatus, checkApiKeyStatus]);

  const bootstrap = useCallback(async (): Promise<void> => {
    const ok = await refreshBackendStatus();
    if (ok) {
      await checkApiKeyStatus();
    } else {
      setApiKeyStatus(null);
    }
  }, [refreshBackendStatus, checkApiKeyStatus]);

  return {
    // State
    backendHealthy,
    backendChecking,
    backendError,
    apiKeyStatus,
    apiKeyBusy,
    // Actions
    refreshBackendStatus,
    checkApiKeyStatus,
    setApiKey,
    clearApiKey,
    bootstrap,
  };
}

export default useApiSettings;
