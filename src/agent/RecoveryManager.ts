/**
 * RecoveryManager - 错误恢复管理器 v4.1
 *
 * 提供智能错误恢复能力：
 * 1. 自动重试（带退避）
 * 2. 替代方案执行
 * 3. 优雅降级
 * 4. 跳过非关键步骤
 *
 * @module agent/RecoveryManager
 */

// ========== 简化的步骤类型 ==========

/**
 * 恢复所需的步骤信息（PlanStep 的子集）
 */
export interface RecoverableStep {
  id: string;
  action: string;
  description?: string;
  parameters?: Record<string, unknown>;
  isWriteOperation?: boolean;
  dependsOn?: string[];
}

// ========== 恢复动作类型 ==========

/**
 * 恢复动作类型
 */
export type RecoveryActionType = "retry" | "skip" | "substitute" | "abort";

/**
 * 恢复动作
 */
export interface RecoveryAction {
  /** 动作类型 */
  type: RecoveryActionType;

  /** 动作描述 */
  description: string;

  /** 重试延迟 (ms)，仅 retry 类型 */
  delay?: number;

  /** 跳过原因，仅 skip 类型 */
  reason?: string;

  /** 替代步骤，仅 substitute 类型 */
  alternativeStep?: RecoverableStep;

  /** 用户消息，仅 abort 类型 */
  userMessage?: string;
}

/**
 * 恢复策略
 */
export interface RecoveryStrategy {
  /** 策略名称 */
  name: string;

  /** 错误模式（正则） */
  errorPattern: RegExp;

  /** 适用的工具（可选，不指定则全部适用） */
  applicableTools?: string[];

  /** 恢复函数 */
  recover: (error: Error, step: RecoverableStep) => Promise<RecoveryAction | null>;

  /** 优先级（数字越小优先级越高） */
  priority: number;
}

// ========== 内置恢复策略 ==========

/**
 * 范围不存在 - 降级为读取选区
 */
const RANGE_NOT_FOUND_STRATEGY: RecoveryStrategy = {
  name: "range_not_found",
  errorPattern: /range.*not\s*found|invalid\s*range|范围.*不存在|无效.*范围/i,
  priority: 10,
  recover: async (error, step) => {
    // 如果是读取操作，降级为读取当前选区
    if (step.action.includes("read") || step.action.includes("get")) {
      return {
        type: "substitute",
        description: "范围不存在，降级为读取当前选区",
        alternativeStep: {
          id: `${step.id}_fallback`,
          action: "excel_read_selection",
          description: "读取当前选区",
          parameters: {},
          isWriteOperation: false,
          dependsOn: [],
        },
      };
    }
    return null;
  },
};

/**
 * 工作表不存在 - 先创建
 */
const SHEET_NOT_EXIST_STRATEGY: RecoveryStrategy = {
  name: "sheet_not_exist",
  errorPattern: /sheet.*not\s*exist|worksheet.*not\s*found|工作表.*不存在/i,
  priority: 10,
  recover: async (error, step) => {
    // 提取工作表名
    const sheetNameMatch = error.message.match(/["']([^"']+)["']/);
    const sheetName = sheetNameMatch?.[1] || "Sheet1";

    return {
      type: "substitute",
      description: `工作表 "${sheetName}" 不存在，自动创建`,
      alternativeStep: {
        id: `${step.id}_create_sheet`,
        action: "excel_create_sheet",
        description: `创建工作表 ${sheetName}`,
        parameters: { name: sheetName },
        isWriteOperation: true,
        dependsOn: [],
      },
    };
  },
};

/**
 * 网络/API 错误 - 重试
 */
const NETWORK_ERROR_STRATEGY: RecoveryStrategy = {
  name: "network_error",
  errorPattern: /network|timeout|ECONNREFUSED|fetch\s*failed|网络.*错误|超时/i,
  priority: 5,
  recover: async () => {
    return {
      type: "retry",
      description: "网络错误，2秒后重试",
      delay: 2000,
    };
  },
};

/**
 * 临时错误 - 短暂重试
 */
const TRANSIENT_ERROR_STRATEGY: RecoveryStrategy = {
  name: "transient_error",
  errorPattern: /busy|locked|temporarily|暂时|繁忙|锁定/i,
  priority: 5,
  recover: async () => {
    return {
      type: "retry",
      description: "资源暂时不可用，1秒后重试",
      delay: 1000,
    };
  },
};

/**
 * 权限错误 - 跳过
 */
const PERMISSION_ERROR_STRATEGY: RecoveryStrategy = {
  name: "permission_error",
  errorPattern: /permission|access\s*denied|权限|拒绝访问/i,
  priority: 20,
  recover: async (error, step) => {
    // 非关键操作可以跳过
    if (!step.isWriteOperation) {
      return {
        type: "skip",
        description: "权限不足，跳过此步骤",
        reason: "权限不足",
      };
    }
    // 关键操作需要中止
    return {
      type: "abort",
      description: "权限不足，无法执行操作",
      userMessage: `无法执行 "${step.description}"：权限不足`,
    };
  },
};

/**
 * 数据格式错误 - 跳过非关键步骤
 */
const DATA_FORMAT_ERROR_STRATEGY: RecoveryStrategy = {
  name: "data_format_error",
  errorPattern: /invalid\s*data|format\s*error|数据格式|格式错误/i,
  priority: 15,
  recover: async (error, step) => {
    if (!step.isWriteOperation) {
      return {
        type: "skip",
        description: "数据格式不匹配，跳过此步骤",
        reason: "数据格式错误",
      };
    }
    return null;
  },
};

/**
 * 公式错误 - 尝试修复
 */
const FORMULA_ERROR_STRATEGY: RecoveryStrategy = {
  name: "formula_error",
  errorPattern: /formula.*error|#REF|#NAME|#VALUE|公式.*错误/i,
  applicableTools: ["excel_set_formula", "excel_batch_formula", "excel_fill_formula"],
  priority: 10,
  recover: async (error, step) => {
    // 尝试简化公式
    const params = step.parameters as Record<string, unknown>;
    const formula = params.formula as string;

    if (formula && formula.includes("(")) {
      // 尝试移除可能有问题的函数嵌套
      return {
        type: "skip",
        description: "公式错误，跳过此步骤",
        reason: `公式可能有语法问题: ${formula.substring(0, 50)}...`,
      };
    }

    return null;
  },
};

/**
 * 默认策略 - 非关键步骤跳过
 */
const DEFAULT_STRATEGY: RecoveryStrategy = {
  name: "default",
  errorPattern: /.*/,
  priority: 100,
  recover: async (error, step) => {
    if (!step.isWriteOperation) {
      return {
        type: "skip",
        description: "遇到错误，跳过非关键步骤",
        reason: error.message,
      };
    }
    // 关键操作返回 null，让调用者决定
    return null;
  },
};

// ========== RecoveryManager 类 ==========

/**
 * 错误恢复管理器
 */
export class RecoveryManager {
  private strategies: RecoveryStrategy[] = [];
  private retryCount: Map<string, number> = new Map();
  private maxRetries: number;

  constructor(config: { maxRetries?: number } = {}) {
    this.maxRetries = config.maxRetries ?? 3;
    this.registerBuiltinStrategies();
  }

  /**
   * 注册内置策略
   */
  private registerBuiltinStrategies(): void {
    this.strategies = [
      NETWORK_ERROR_STRATEGY,
      TRANSIENT_ERROR_STRATEGY,
      RANGE_NOT_FOUND_STRATEGY,
      SHEET_NOT_EXIST_STRATEGY,
      FORMULA_ERROR_STRATEGY,
      PERMISSION_ERROR_STRATEGY,
      DATA_FORMAT_ERROR_STRATEGY,
      DEFAULT_STRATEGY,
    ].sort((a, b) => a.priority - b.priority);
  }

  /**
   * 注册自定义策略
   */
  registerStrategy(strategy: RecoveryStrategy): void {
    this.strategies.push(strategy);
    this.strategies.sort((a, b) => a.priority - b.priority);
  }

  /**
   * 尝试恢复
   *
   * @param step 失败的步骤
   * @param error 错误对象
   * @returns 恢复动作，或 null 表示无法恢复
   */
  async recover(step: RecoverableStep, error: Error): Promise<RecoveryAction | null> {
    const stepKey = step.id;
    const currentRetry = this.retryCount.get(stepKey) || 0;

    console.log(`[RecoveryManager] 尝试恢复步骤 ${step.id}，错误: ${error.message}`);

    // 找到匹配的策略
    for (const strategy of this.strategies) {
      // 检查错误模式是否匹配
      if (!strategy.errorPattern.test(error.message)) {
        continue;
      }

      // 检查工具是否适用
      if (strategy.applicableTools && !strategy.applicableTools.includes(step.action)) {
        continue;
      }

      console.log(`[RecoveryManager] 匹配策略: ${strategy.name}`);

      // 执行恢复
      const action = await strategy.recover(error, step);

      if (action) {
        // 检查重试次数
        if (action.type === "retry") {
          if (currentRetry >= this.maxRetries) {
            console.log(`[RecoveryManager] 已达最大重试次数 (${this.maxRetries})`);
            continue; // 跳过此策略，尝试下一个
          }
          this.retryCount.set(stepKey, currentRetry + 1);
        }

        console.log(`[RecoveryManager] 恢复动作: ${action.type} - ${action.description}`);
        return action;
      }
    }

    console.log(`[RecoveryManager] 无法恢复步骤 ${step.id}`);
    return null;
  }

  /**
   * 重置重试计数
   */
  resetRetryCount(stepId?: string): void {
    if (stepId) {
      this.retryCount.delete(stepId);
    } else {
      this.retryCount.clear();
    }
  }

  /**
   * 获取已注册的策略列表
   */
  getStrategies(): Array<{ name: string; priority: number }> {
    return this.strategies.map((s) => ({
      name: s.name,
      priority: s.priority,
    }));
  }
}

// ========== 工厂函数 ==========

/**
 * 创建恢复管理器
 */
export function createRecoveryManager(config?: { maxRetries?: number }): RecoveryManager {
  return new RecoveryManager(config);
}

export default RecoveryManager;
