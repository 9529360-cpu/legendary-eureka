/**
 * ConfigManager 测试
 *
 * 覆盖：
 * - 配置读取与设置
 * - 环境切换
 * - 配置验证
 * - 持久化
 * - 变更通知
 *
 * @version 2.0.0 - 匹配新的 ConfigManager API
 */

import {
  ConfigManager,
  type Environment,
  type ApiConfig,
  type ExcelConfig,
  type AgentConfig,
  type UiConfig,
  type LoggingConfig,
  type SecurityConfig,
  type FeatureFlags,
  type ConfigChangeEvent,
} from "../core/ConfigManager";

// 模拟 localStorage
const mockLocalStorage = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: (key: string) => store[key] || null,
    setItem: (key: string, value: string) => {
      store[key] = value;
    },
    removeItem: (key: string) => {
      delete store[key];
    },
    clear: () => {
      store = {};
    },
  };
})();

Object.defineProperty(global, "localStorage", { value: mockLocalStorage });

// 忽略未使用的类型导入（用于类型检查）
const _typeCheck:
  | Environment
  | ApiConfig
  | ExcelConfig
  | AgentConfig
  | UiConfig
  | LoggingConfig
  | SecurityConfig
  | FeatureFlags
  | null = null;
void _typeCheck;

describe("ConfigManager", () => {
  beforeEach(() => {
    ConfigManager.reset();
    mockLocalStorage.clear();
  });

  describe("环境管理", () => {
    it("应该返回当前环境", () => {
      const env = ConfigManager.environment;

      expect(["development", "staging", "production", "test"]).toContain(env);
    });

    it("应该检测是否为开发环境", () => {
      expect(ConfigManager.isDevelopment).toBe(true);
    });

    it("应该检测是否为生产环境", () => {
      expect(ConfigManager.isProduction).toBe(false);
    });
  });

  describe("API 配置", () => {
    it("应该返回默认 API 配置", () => {
      const apiConfig = ConfigManager.api;

      expect(apiConfig).toBeDefined();
      expect(apiConfig.backendUrl).toBeDefined();
      expect(apiConfig.timeout).toBeGreaterThan(0);
    });

    it("应该通过 set 更新 API 配置", () => {
      ConfigManager.set("api.backendUrl", "https://custom-api.example.com");
      ConfigManager.set("api.timeout", 60000);

      const apiConfig = ConfigManager.api;
      expect(apiConfig.backendUrl).toBe("https://custom-api.example.com");
      expect(apiConfig.timeout).toBe(60000);
    });

    it("应该通过 get 获取特定配置值", () => {
      const timeout = ConfigManager.get<number>("api.timeout");
      expect(timeout).toBeGreaterThan(0);
    });
  });

  describe("Excel 配置", () => {
    it("应该返回默认 Excel 配置", () => {
      const excelConfig = ConfigManager.excel;

      expect(excelConfig).toBeDefined();
      expect(excelConfig.maxRowsPerOperation).toBeGreaterThan(0);
      expect(excelConfig.operationTimeout).toBeGreaterThan(0);
    });

    it("应该更新 Excel 配置", () => {
      ConfigManager.set("excel.maxRowsPerOperation", 50000);
      ConfigManager.set("excel.enableUndo", true);

      const excelConfig = ConfigManager.excel;
      expect(excelConfig.maxRowsPerOperation).toBe(50000);
      expect(excelConfig.enableUndo).toBe(true);
    });
  });

  describe("Agent 配置", () => {
    it("应该返回默认 Agent 配置", () => {
      const agentConfig = ConfigManager.agent;

      expect(agentConfig).toBeDefined();
      expect(agentConfig.maxReactIterations).toBeGreaterThan(0);
      expect(agentConfig.maxPlanSteps).toBeGreaterThan(0);
    });

    it("应该更新 Agent 配置", () => {
      ConfigManager.set("agent.maxReactIterations", 20);
      ConfigManager.set("agent.verboseLogging", true);

      const agentConfig = ConfigManager.agent;
      expect(agentConfig.maxReactIterations).toBe(20);
      expect(agentConfig.verboseLogging).toBe(true);
    });
  });

  describe("UI 配置", () => {
    it("应该返回默认 UI 配置", () => {
      const uiConfig = ConfigManager.ui;

      expect(uiConfig).toBeDefined();
      expect(uiConfig.theme).toBeDefined();
    });

    it("应该更新主题", () => {
      ConfigManager.set("ui.theme", "dark");

      const uiConfig = ConfigManager.ui;
      expect(uiConfig.theme).toBe("dark");
    });
  });

  describe("日志配置", () => {
    it("应该返回默认日志配置", () => {
      const loggingConfig = ConfigManager.logging;

      expect(loggingConfig).toBeDefined();
      expect(loggingConfig.level).toBeDefined();
    });

    it("应该更新日志级别", () => {
      ConfigManager.set("logging.level", "debug");
      ConfigManager.set("logging.enableConsole", true);

      const loggingConfig = ConfigManager.logging;
      expect(loggingConfig.level).toBe("debug");
      expect(loggingConfig.enableConsole).toBe(true);
    });
  });

  describe("安全配置", () => {
    it("应该返回默认安全配置", () => {
      const securityConfig = ConfigManager.security;

      expect(securityConfig).toBeDefined();
      expect(securityConfig.enablePermissionCheck).toBeDefined();
    });

    it("应该更新安全配置", () => {
      ConfigManager.set("security.enablePermissionCheck", true);
      ConfigManager.set("security.confirmationThreshold", "low");

      const securityConfig = ConfigManager.security;
      expect(securityConfig.enablePermissionCheck).toBe(true);
      expect(securityConfig.confirmationThreshold).toBe("low");
    });
  });

  describe("功能开关", () => {
    it("应该返回默认功能开关", () => {
      const flags = ConfigManager.features;

      expect(flags).toBeDefined();
      expect(flags.aiIntegration).toBeDefined();
    });

    it("应该更新功能开关", () => {
      ConfigManager.set("features.debugMode", true);

      const flags = ConfigManager.features;
      expect(flags.debugMode).toBe(true);
    });

    it("应该检查调试模式", () => {
      ConfigManager.set("features.debugMode", true);
      expect(ConfigManager.isDebug).toBe(true);

      ConfigManager.set("features.debugMode", false);
      expect(ConfigManager.isDebug).toBe(false);
    });
  });

  describe("批量更新", () => {
    it("应该批量更新配置", () => {
      ConfigManager.update({
        api: {
          timeout: 45000,
        },
        ui: {
          theme: "dark",
        },
      });

      expect(ConfigManager.api.timeout).toBe(45000);
      expect(ConfigManager.ui.theme).toBe("dark");
    });
  });

  describe("配置验证", () => {
    it("应该验证有效配置", () => {
      const result = ConfigManager.validate();

      expect(result.valid).toBe(true);
      expect(result.errors).toHaveLength(0);
    });

    it("应该检测无效配置", () => {
      ConfigManager.set("api.timeout", -1);

      const result = ConfigManager.validate();

      expect(result.valid).toBe(false);
      expect(result.errors.length).toBeGreaterThan(0);
    });
  });

  describe("变更通知", () => {
    it("应该在配置变更时触发监听器", () => {
      const changes: ConfigChangeEvent[] = [];
      ConfigManager.addChangeListener((event) => {
        changes.push(event);
      });

      ConfigManager.set("api.backendUrl", "https://new-api.example.com");

      expect(changes.length).toBe(1);
      expect(changes[0].path).toBe("api.backendUrl");
      expect(changes[0].newValue).toBe("https://new-api.example.com");
    });

    it("应该移除监听器", () => {
      const changes: ConfigChangeEvent[] = [];
      const unsubscribe = ConfigManager.addChangeListener((event) => {
        changes.push(event);
      });

      ConfigManager.set("api.backendUrl", "https://1.example.com");
      unsubscribe();
      ConfigManager.set("api.backendUrl", "https://2.example.com");

      expect(changes.length).toBe(1);
    });
  });

  describe("配置导出/导入", () => {
    it("应该导出所有配置为 JSON", () => {
      const exported = ConfigManager.export();

      expect(exported).toBeDefined();
      expect(typeof exported).toBe("string");

      const parsed = JSON.parse(exported);
      expect(parsed.api).toBeDefined();
      expect(parsed.excel).toBeDefined();
      expect(parsed.agent).toBeDefined();
    });

    it("应该导入配置", () => {
      const configToImport = JSON.stringify({
        api: { backendUrl: "https://imported-api.example.com", timeout: 55000 },
        ui: { theme: "dark" },
      });

      const result = ConfigManager.import(configToImport);

      expect(result).toBe(true);
      expect(ConfigManager.api.backendUrl).toBe("https://imported-api.example.com");
      expect(ConfigManager.ui.theme).toBe("dark");
    });

    it("应该在导入无效 JSON 时返回 false", () => {
      const result = ConfigManager.import("invalid json");

      expect(result).toBe(false);
    });
  });

  describe("配置重置", () => {
    it("应该重置到默认值", () => {
      ConfigManager.set("api.timeout", 99999);

      ConfigManager.reset();

      const apiConfig = ConfigManager.api;
      expect(apiConfig.timeout).not.toBe(99999);
    });
  });

  describe("类型安全", () => {
    it("应该提供类型安全的配置访问", () => {
      const apiConfig = ConfigManager.api;

      const _url: string = apiConfig.backendUrl;
      const _timeout: number = apiConfig.timeout;

      expect(typeof _url).toBe("string");
      expect(typeof _timeout).toBe("number");
    });

    it("应该通过泛型访问配置", () => {
      const timeout = ConfigManager.get<number>("api.timeout");

      expect(timeout).toBeDefined();
      expect(typeof timeout).toBe("number");
    });

    it("应该返回 undefined 对于不存在的路径", () => {
      const nonexistent = ConfigManager.get("nonexistent.path");

      expect(nonexistent).toBeUndefined();
    });
  });

  describe("getConfig 方法", () => {
    it("应该返回完整的只读配置", () => {
      const config = ConfigManager.getConfig();

      expect(config).toBeDefined();
      expect(config.api).toBeDefined();
      expect(config.excel).toBeDefined();
      expect(config.agent).toBeDefined();
      expect(config.ui).toBeDefined();
      expect(config.logging).toBeDefined();
      expect(config.security).toBeDefined();
      expect(config.features).toBeDefined();
      expect(config.environment).toBeDefined();
      expect(config.version).toBeDefined();
    });
  });
});
