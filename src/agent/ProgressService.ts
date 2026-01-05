/**
 * ProgressService - 进度更新服务 v1.0
 *
 * 借鉴自 Activepieces progress.service.ts
 *
 * 功能特性：
 * 1. 防抖机制避免频繁更新
 * 2. 互斥锁防止并发冲突
 * 3. 最大阈值强制推送
 * 4. 回调通知
 *
 * @see https://github.com/activepieces/activepieces
 */

import { Logger } from "../utils/Logger";

const MODULE_NAME = "ProgressService";

// ==================== Mutex 实现 ====================

/**
 * 互斥锁
 *
 * 用于防止并发进度更新
 */
export class Mutex {
  private locked = false;
  private waiting: Array<() => void> = [];

  /**
   * 获取锁
   */
  async acquire(): Promise<void> {
    if (!this.locked) {
      this.locked = true;
      return;
    }

    // 等待锁释放
    return new Promise<void>((resolve) => {
      this.waiting.push(resolve);
    });
  }

  /**
   * 释放锁
   */
  release(): void {
    if (this.waiting.length > 0) {
      const next = this.waiting.shift();
      next?.();
    } else {
      this.locked = false;
    }
  }

  /**
   * 在锁内执行
   */
  async runExclusive<T>(fn: () => Promise<T>): Promise<T> {
    await this.acquire();
    try {
      return await fn();
    } finally {
      this.release();
    }
  }

  /**
   * 是否已锁定
   */
  isLocked(): boolean {
    return this.locked;
  }
}

// ==================== 类型定义 ====================

/**
 * 进度信息
 */
export interface ProgressInfo {
  /** 当前步骤 */
  currentStep: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 进度百分比 (0-100) */
  percentage: number;
  /** 当前状态描述 */
  status: string;
  /** 额外数据 */
  metadata?: Record<string, unknown>;
  /** 时间戳 */
  timestamp: Date;
}

/**
 * 进度监听器
 */
export type ProgressListener = (progress: ProgressInfo) => void | Promise<void>;

/**
 * 进度服务配置
 */
export interface ProgressServiceConfig {
  /** 防抖延迟（毫秒） */
  debounceMs: number;
  /** 最大间隔阈值（毫秒） - 超过此时间强制推送 */
  maxThresholdMs: number;
  /** 是否启用日志 */
  enableLogging: boolean;
}

/**
 * 默认配置
 */
export const DEFAULT_PROGRESS_CONFIG: ProgressServiceConfig = {
  debounceMs: 500, // 0.5秒防抖
  maxThresholdMs: 5000, // 5秒强制推送
  enableLogging: false,
};

// ==================== ProgressService 类 ====================

/**
 * 进度服务
 *
 * 管理进度更新，带防抖和互斥锁保护
 */
export class ProgressService {
  private mutex = new Mutex();
  private config: ProgressServiceConfig;
  private listeners: Set<ProgressListener> = new Set();
  private lastProgress: ProgressInfo | null = null;
  private lastPushTime: number = 0;
  private pendingProgress: ProgressInfo | null = null;
  private debounceTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(config: Partial<ProgressServiceConfig> = {}) {
    this.config = { ...DEFAULT_PROGRESS_CONFIG, ...config };
  }

  // ==================== 监听器管理 ====================

  /**
   * 添加进度监听器
   */
  addListener(listener: ProgressListener): () => void {
    this.listeners.add(listener);

    // 返回取消订阅函数
    return () => {
      this.listeners.delete(listener);
    };
  }

  /**
   * 移除所有监听器
   */
  removeAllListeners(): void {
    this.listeners.clear();
  }

  /**
   * 获取监听器数量
   */
  getListenerCount(): number {
    return this.listeners.size;
  }

  // ==================== 进度更新 ====================

  /**
   * 更新进度（带防抖）
   */
  async updateProgress(progress: Omit<ProgressInfo, "timestamp">): Promise<void> {
    const fullProgress: ProgressInfo = {
      ...progress,
      timestamp: new Date(),
    };

    this.pendingProgress = fullProgress;

    // 检查是否需要立即推送（超过最大阈值）
    const now = Date.now();
    const timeSinceLastPush = now - this.lastPushTime;

    if (timeSinceLastPush >= this.config.maxThresholdMs) {
      await this.flushProgress();
      return;
    }

    // 使用防抖
    this.scheduleFlush();
  }

  /**
   * 立即推送进度（跳过防抖）
   */
  async pushProgressNow(progress: Omit<ProgressInfo, "timestamp">): Promise<void> {
    const fullProgress: ProgressInfo = {
      ...progress,
      timestamp: new Date(),
    };

    this.pendingProgress = fullProgress;
    await this.flushProgress();
  }

  /**
   * 设置防抖定时器
   */
  private scheduleFlush(): void {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
    }

    this.debounceTimer = setTimeout(async () => {
      await this.flushProgress();
    }, this.config.debounceMs);
  }

  /**
   * 推送待处理的进度更新
   */
  private async flushProgress(): Promise<void> {
    if (!this.pendingProgress) return;

    await this.mutex.runExclusive(async () => {
      if (!this.pendingProgress) return;

      const progress = this.pendingProgress;
      this.pendingProgress = null;

      if (this.config.enableLogging) {
        Logger.debug(MODULE_NAME, "推送进度", {
          step: `${progress.currentStep}/${progress.totalSteps}`,
          percentage: `${progress.percentage}%`,
          status: progress.status,
        });
      }

      // 通知所有监听器
      await this.notifyListeners(progress);

      this.lastProgress = progress;
      this.lastPushTime = Date.now();
    });
  }

  /**
   * 通知所有监听器
   */
  private async notifyListeners(progress: ProgressInfo): Promise<void> {
    const promises = Array.from(this.listeners).map(async (listener) => {
      try {
        await listener(progress);
      } catch (error) {
        Logger.error(MODULE_NAME, "进度监听器错误", { error });
      }
    });

    await Promise.all(promises);
  }

  // ==================== 便捷方法 ====================

  /**
   * 开始任务
   */
  async start(totalSteps: number, status: string = "开始执行"): Promise<void> {
    await this.pushProgressNow({
      currentStep: 0,
      totalSteps,
      percentage: 0,
      status,
    });
  }

  /**
   * 更新当前步骤
   */
  async step(currentStep: number, totalSteps: number, status: string): Promise<void> {
    const percentage = Math.round((currentStep / totalSteps) * 100);
    await this.updateProgress({
      currentStep,
      totalSteps,
      percentage,
      status,
    });
  }

  /**
   * 完成任务
   */
  async complete(totalSteps: number, status: string = "执行完成"): Promise<void> {
    await this.pushProgressNow({
      currentStep: totalSteps,
      totalSteps,
      percentage: 100,
      status,
    });
  }

  /**
   * 错误终止
   */
  async error(currentStep: number, totalSteps: number, errorMessage: string): Promise<void> {
    await this.pushProgressNow({
      currentStep,
      totalSteps,
      percentage: Math.round((currentStep / totalSteps) * 100),
      status: `错误: ${errorMessage}`,
      metadata: { error: true },
    });
  }

  // ==================== 状态查询 ====================

  /**
   * 获取最后的进度信息
   */
  getLastProgress(): ProgressInfo | null {
    return this.lastProgress;
  }

  /**
   * 获取当前百分比
   */
  getCurrentPercentage(): number {
    return this.lastProgress?.percentage ?? 0;
  }

  /**
   * 是否正在处理更新
   */
  isUpdating(): boolean {
    return this.mutex.isLocked() || this.pendingProgress !== null;
  }

  // ==================== 生命周期 ====================

  /**
   * 重置服务状态
   */
  reset(): void {
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
    this.pendingProgress = null;
    this.lastProgress = null;
    this.lastPushTime = 0;
  }

  /**
   * 销毁服务
   */
  dispose(): void {
    this.reset();
    this.removeAllListeners();
  }
}

// ==================== 全局实例 ====================

let globalProgressService: ProgressService | null = null;

/**
 * 获取全局进度服务实例
 */
export function getProgressService(config?: Partial<ProgressServiceConfig>): ProgressService {
  if (!globalProgressService) {
    globalProgressService = new ProgressService(config);
  }
  return globalProgressService;
}

/**
 * 重置全局进度服务
 */
export function resetProgressService(): void {
  if (globalProgressService) {
    globalProgressService.dispose();
    globalProgressService = null;
  }
}

// ==================== 辅助函数 ====================

/**
 * 创建进度追踪器
 *
 * @example
 * ```typescript
 * const tracker = createProgressTracker(10, (progress) => {
 *   console.log(`${progress.percentage}%`);
 * });
 *
 * await tracker.step('处理第1步');
 * await tracker.step('处理第2步');
 * await tracker.complete();
 * ```
 */
export function createProgressTracker(
  totalSteps: number,
  onProgress: ProgressListener
): {
  step: (status: string) => Promise<void>;
  complete: (status?: string) => Promise<void>;
  error: (message: string) => Promise<void>;
} {
  let currentStep = 0;
  const service = new ProgressService({ debounceMs: 100 });
  service.addListener(onProgress);

  return {
    step: async (status: string) => {
      currentStep++;
      await service.step(currentStep, totalSteps, status);
    },
    complete: async (status: string = "完成") => {
      await service.complete(totalSteps, status);
    },
    error: async (message: string) => {
      await service.error(currentStep, totalSteps, message);
    },
  };
}

// ==================== 导出 ====================

export default ProgressService;
