// 声明 window.ExcelAssistant 类型，解决App.tsx 类型错误
export type BackendHealth = {
  success?: boolean;
  status?: string;
  port?: number;
  model?: string;
  configured?: boolean;
  allowedOrigins?: string[];
};

export type AssistantStatus = {
  backend: {
    ok: boolean;
    lastCheckedAt?: string;
    error?: string;
    health?: BackendHealth;
  };
};

export interface AuditLogEntry {
  time: string;
  action: string;
  payload: any;
}

declare global {
  interface Window {
    ExcelAssistant?: {
      status: AssistantStatus;
      checkBackend: () => Promise<AssistantStatus["backend"]>;
      chat: (message: string, context?: any) => Promise<any>;
      perceiveExcel: () => Promise<any>;
      audit: {
        log: (action: string, payload: any) => void;
        list: () => AuditLogEntry[];
        clear: () => void;
      };
    };
  }
}
