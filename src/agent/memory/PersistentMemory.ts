/**
 * PersistentMemory - 持久化内存层 v4.1
 *
 * 基于 IndexedDB 实现经验记忆持久化
 *
 * 核心特性：
 * 1. 对话历史持久化
 * 2. 经验记忆存储
 * 3. 工具统计持久化
 * 4. 自动过期清理
 *
 * @module agent/memory/PersistentMemory
 */

// ========== 类型定义 ==========

/**
 * 存储的对话消息
 */
export interface StoredMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  timestamp: number;
  sessionId: string;
  metadata?: Record<string, unknown>;
}

/**
 * 存储的经验
 */
export interface StoredEpisode {
  id: string;
  sessionId: string;
  intent: string;
  actions: string[];
  result: "success" | "failure" | "partial";
  feedback?: string;
  timestamp: number;
  duration: number;
  toolsUsed: string[];
  errorMessage?: string;
}

/**
 * 工具统计
 */
export interface ToolStats {
  name: string;
  totalCalls: number;
  successCalls: number;
  failureCalls: number;
  totalDuration: number;
  avgDuration: number;
  lastUsed: number;
}

/**
 * 会话摘要
 */
export interface SessionSummary {
  id: string;
  startTime: number;
  endTime: number;
  messageCount: number;
  successRate: number;
  title: string;
}

/**
 * 内存配置
 */
export interface PersistentMemoryConfig {
  /** 数据库名称 */
  dbName?: string;

  /** 数据库版本 */
  dbVersion?: number;

  /** 消息最大保留天数 */
  messageRetentionDays?: number;

  /** 经验最大保留天数 */
  episodeRetentionDays?: number;

  /** 自动清理间隔 (ms) */
  cleanupInterval?: number;
}

// ========== 默认配置 ==========

const DEFAULT_CONFIG: Required<PersistentMemoryConfig> = {
  dbName: "excel-copilot-memory",
  dbVersion: 1,
  messageRetentionDays: 30,
  episodeRetentionDays: 90,
  cleanupInterval: 24 * 60 * 60 * 1000, // 每天清理一次
};

// ========== IndexedDB 工具函数 ==========

/**
 * 打开数据库
 */
function openDatabase(config: Required<PersistentMemoryConfig>): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(config.dbName, config.dbVersion);

    request.onerror = () => {
      reject(new Error(`打开数据库失败: ${request.error?.message}`));
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;

      // 创建消息存储
      if (!db.objectStoreNames.contains("messages")) {
        const messageStore = db.createObjectStore("messages", { keyPath: "id" });
        messageStore.createIndex("sessionId", "sessionId", { unique: false });
        messageStore.createIndex("timestamp", "timestamp", { unique: false });
      }

      // 创建经验存储
      if (!db.objectStoreNames.contains("episodes")) {
        const episodeStore = db.createObjectStore("episodes", { keyPath: "id" });
        episodeStore.createIndex("sessionId", "sessionId", { unique: false });
        episodeStore.createIndex("timestamp", "timestamp", { unique: false });
        episodeStore.createIndex("intent", "intent", { unique: false });
        episodeStore.createIndex("result", "result", { unique: false });
      }

      // 创建工具统计存储
      if (!db.objectStoreNames.contains("toolStats")) {
        db.createObjectStore("toolStats", { keyPath: "name" });
      }

      // 创建会话摘要存储
      if (!db.objectStoreNames.contains("sessions")) {
        const sessionStore = db.createObjectStore("sessions", { keyPath: "id" });
        sessionStore.createIndex("startTime", "startTime", { unique: false });
      }

      // 创建设置存储
      if (!db.objectStoreNames.contains("settings")) {
        db.createObjectStore("settings", { keyPath: "key" });
      }
    };
  });
}

/**
 * 执行事务操作
 */
async function withTransaction<T>(
  db: IDBDatabase,
  storeNames: string | string[],
  mode: IDBTransactionMode,
  callback: (stores: Record<string, IDBObjectStore>) => Promise<T> | T
): Promise<T> {
  return new Promise((resolve, reject) => {
    const names = Array.isArray(storeNames) ? storeNames : [storeNames];
    const transaction = db.transaction(names, mode);
    const stores: Record<string, IDBObjectStore> = {};

    for (const name of names) {
      stores[name] = transaction.objectStore(name);
    }

    transaction.onerror = () => {
      reject(new Error(`事务失败: ${transaction.error?.message}`));
    };

    try {
      const result = callback(stores);
      if (result instanceof Promise) {
        result.then(resolve).catch(reject);
      } else {
        transaction.oncomplete = () => resolve(result);
      }
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * 添加记录
 */
function addRecord<T>(store: IDBObjectStore, record: T): Promise<void> {
  return new Promise((resolve, reject) => {
    const request = store.put(record);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(new Error(`添加记录失败: ${request.error?.message}`));
  });
}

/**
 * 获取记录
 */
function getRecord<T>(store: IDBObjectStore, key: IDBValidKey): Promise<T | undefined> {
  return new Promise((resolve, reject) => {
    const request = store.get(key);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(new Error(`获取记录失败: ${request.error?.message}`));
  });
}

/**
 * 获取所有记录
 */
function getAllRecords<T>(store: IDBObjectStore): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(new Error(`获取记录失败: ${request.error?.message}`));
  });
}

/**
 * 按索引获取记录
 */
function getRecordsByIndex<T>(
  store: IDBObjectStore,
  indexName: string,
  value: IDBValidKey
): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const index = store.index(indexName);
    const request = index.getAll(value);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(new Error(`获取记录失败: ${request.error?.message}`));
  });
}

/**
 * 删除记录
 */
function deleteRecord(store: IDBObjectStore, key: IDBValidKey): Promise<void> {
  return new Promise((resolve, reject) => {
    const request = store.delete(key);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(new Error(`删除记录失败: ${request.error?.message}`));
  });
}

// ========== PersistentMemory 类 ==========

/**
 * 持久化内存管理器
 */
export class PersistentMemory {
  private config: Required<PersistentMemoryConfig>;
  private db: IDBDatabase | null = null;
  private initialized: boolean = false;
  private cleanupTimer: ReturnType<typeof setInterval> | null = null;

  constructor(config: PersistentMemoryConfig = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * 初始化
   */
  async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      this.db = await openDatabase(this.config);
      this.initialized = true;

      // 启动自动清理
      this.startAutoCleanup();

      console.log("[PersistentMemory] 初始化完成");
    } catch (error) {
      console.error("[PersistentMemory] 初始化失败:", error);
      throw error;
    }
  }

  /**
   * 关闭
   */
  close(): void {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = null;
    }

    if (this.db) {
      this.db.close();
      this.db = null;
    }

    this.initialized = false;
  }

  /**
   * 确保已初始化
   */
  private ensureInitialized(): void {
    if (!this.initialized || !this.db) {
      throw new Error("PersistentMemory 未初始化，请先调用 initialize()");
    }
  }

  // ========== 消息相关 ==========

  /**
   * 保存消息
   */
  async saveMessage(message: Omit<StoredMessage, "id" | "timestamp">): Promise<string> {
    this.ensureInitialized();

    const id = `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const storedMessage: StoredMessage = {
      ...message,
      id,
      timestamp: Date.now(),
    };

    await withTransaction(this.db!, "messages", "readwrite", async (stores) => {
      await addRecord(stores.messages, storedMessage);
    });

    return id;
  }

  /**
   * 获取会话消息
   */
  async getSessionMessages(sessionId: string): Promise<StoredMessage[]> {
    this.ensureInitialized();

    return withTransaction(this.db!, "messages", "readonly", async (stores) => {
      const messages = await getRecordsByIndex<StoredMessage>(
        stores.messages,
        "sessionId",
        sessionId
      );
      return messages.sort((a, b) => a.timestamp - b.timestamp);
    });
  }

  /**
   * 获取最近的消息
   */
  async getRecentMessages(limit: number = 50): Promise<StoredMessage[]> {
    this.ensureInitialized();

    return withTransaction(this.db!, "messages", "readonly", async (stores) => {
      const messages = await getAllRecords<StoredMessage>(stores.messages);
      return messages.sort((a, b) => b.timestamp - a.timestamp).slice(0, limit);
    });
  }

  // ========== 经验相关 ==========

  /**
   * 保存经验
   */
  async saveEpisode(episode: Omit<StoredEpisode, "id" | "timestamp">): Promise<string> {
    this.ensureInitialized();

    const id = `ep_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const storedEpisode: StoredEpisode = {
      ...episode,
      id,
      timestamp: Date.now(),
    };

    await withTransaction(this.db!, "episodes", "readwrite", async (stores) => {
      await addRecord(stores.episodes, storedEpisode);
    });

    return id;
  }

  /**
   * 获取相似经验
   */
  async getSimilarEpisodes(intent: string, limit: number = 5): Promise<StoredEpisode[]> {
    this.ensureInitialized();

    return withTransaction(this.db!, "episodes", "readonly", async (stores) => {
      const episodes = await getAllRecords<StoredEpisode>(stores.episodes);

      // 简单的关键词匹配
      const keywords = intent.toLowerCase().split(/\s+/);

      const scored = episodes.map((ep) => {
        const epKeywords = ep.intent.toLowerCase().split(/\s+/);
        let score = 0;
        for (const kw of keywords) {
          if (epKeywords.some((ek) => ek.includes(kw) || kw.includes(ek))) {
            score++;
          }
        }
        return { episode: ep, score };
      });

      return scored
        .filter((s) => s.score > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, limit)
        .map((s) => s.episode);
    });
  }

  /**
   * 获取成功经验
   */
  async getSuccessfulEpisodes(limit: number = 10): Promise<StoredEpisode[]> {
    this.ensureInitialized();

    return withTransaction(this.db!, "episodes", "readonly", async (stores) => {
      const episodes = await getRecordsByIndex<StoredEpisode>(stores.episodes, "result", "success");
      return episodes.sort((a, b) => b.timestamp - a.timestamp).slice(0, limit);
    });
  }

  // ========== 工具统计相关 ==========

  /**
   * 更新工具统计
   */
  async updateToolStats(name: string, success: boolean, duration: number): Promise<void> {
    this.ensureInitialized();

    await withTransaction(this.db!, "toolStats", "readwrite", async (stores) => {
      const existing = await getRecord<ToolStats>(stores.toolStats, name);

      const stats: ToolStats = existing ?? {
        name,
        totalCalls: 0,
        successCalls: 0,
        failureCalls: 0,
        totalDuration: 0,
        avgDuration: 0,
        lastUsed: Date.now(),
      };

      stats.totalCalls++;
      if (success) {
        stats.successCalls++;
      } else {
        stats.failureCalls++;
      }
      stats.totalDuration += duration;
      stats.avgDuration = stats.totalDuration / stats.totalCalls;
      stats.lastUsed = Date.now();

      await addRecord(stores.toolStats, stats);
    });
  }

  /**
   * 获取工具统计
   */
  async getToolStats(name?: string): Promise<ToolStats | ToolStats[] | null> {
    this.ensureInitialized();

    return withTransaction(this.db!, "toolStats", "readonly", async (stores) => {
      if (name) {
        return (await getRecord<ToolStats>(stores.toolStats, name)) ?? null;
      }
      return getAllRecords<ToolStats>(stores.toolStats);
    });
  }

  // ========== 会话相关 ==========

  /**
   * 保存会话摘要
   */
  async saveSession(session: SessionSummary): Promise<void> {
    this.ensureInitialized();

    await withTransaction(this.db!, "sessions", "readwrite", async (stores) => {
      await addRecord(stores.sessions, session);
    });
  }

  /**
   * 获取最近会话
   */
  async getRecentSessions(limit: number = 10): Promise<SessionSummary[]> {
    this.ensureInitialized();

    return withTransaction(this.db!, "sessions", "readonly", async (stores) => {
      const sessions = await getAllRecords<SessionSummary>(stores.sessions);
      return sessions.sort((a, b) => b.startTime - a.startTime).slice(0, limit);
    });
  }

  // ========== 设置相关 ==========

  /**
   * 保存设置
   */
  async saveSetting(key: string, value: unknown): Promise<void> {
    this.ensureInitialized();

    await withTransaction(this.db!, "settings", "readwrite", async (stores) => {
      await addRecord(stores.settings, { key, value, updatedAt: Date.now() });
    });
  }

  /**
   * 获取设置
   */
  async getSetting<T>(key: string): Promise<T | undefined> {
    this.ensureInitialized();

    return withTransaction(this.db!, "settings", "readonly", async (stores) => {
      const record = await getRecord<{ key: string; value: T }>(stores.settings, key);
      return record?.value;
    });
  }

  // ========== 清理相关 ==========

  /**
   * 启动自动清理
   */
  private startAutoCleanup(): void {
    if (this.cleanupTimer) return;

    // 立即执行一次清理
    this.cleanup().catch(console.error);

    // 定时清理
    this.cleanupTimer = setInterval(() => {
      this.cleanup().catch(console.error);
    }, this.config.cleanupInterval);
  }

  /**
   * 清理过期数据
   */
  async cleanup(): Promise<{ messagesDeleted: number; episodesDeleted: number }> {
    this.ensureInitialized();

    const messageThreshold = Date.now() - this.config.messageRetentionDays * 24 * 60 * 60 * 1000;
    const episodeThreshold = Date.now() - this.config.episodeRetentionDays * 24 * 60 * 60 * 1000;

    let messagesDeleted = 0;
    let episodesDeleted = 0;

    // 清理消息
    await withTransaction(this.db!, "messages", "readwrite", async (stores) => {
      const messages = await getAllRecords<StoredMessage>(stores.messages);
      for (const msg of messages) {
        if (msg.timestamp < messageThreshold) {
          await deleteRecord(stores.messages, msg.id);
          messagesDeleted++;
        }
      }
    });

    // 清理经验
    await withTransaction(this.db!, "episodes", "readwrite", async (stores) => {
      const episodes = await getAllRecords<StoredEpisode>(stores.episodes);
      for (const ep of episodes) {
        if (ep.timestamp < episodeThreshold) {
          await deleteRecord(stores.episodes, ep.id);
          episodesDeleted++;
        }
      }
    });

    if (messagesDeleted > 0 || episodesDeleted > 0) {
      console.log(
        `[PersistentMemory] 清理完成: ${messagesDeleted} 消息, ${episodesDeleted} 经验`
      );
    }

    return { messagesDeleted, episodesDeleted };
  }

  /**
   * 清空所有数据
   */
  async clearAll(): Promise<void> {
    this.ensureInitialized();

    const storeNames = ["messages", "episodes", "toolStats", "sessions", "settings"];

    for (const name of storeNames) {
      await withTransaction(this.db!, name, "readwrite", async (stores) => {
        stores[name].clear();
      });
    }

    console.log("[PersistentMemory] 已清空所有数据");
  }

  /**
   * 获取存储统计
   */
  async getStorageStats(): Promise<{
    messages: number;
    episodes: number;
    toolStats: number;
    sessions: number;
  }> {
    this.ensureInitialized();

    const storeNames = ["messages", "episodes", "toolStats", "sessions"];
    const stats: Record<string, number> = {};

    for (const name of storeNames) {
      await withTransaction(this.db!, name, "readonly", async (stores) => {
        const records = await getAllRecords(stores[name]);
        stats[name] = records.length;
      });
    }

    return stats as ReturnType<typeof this.getStorageStats> extends Promise<infer T> ? T : never;
  }

  /**
   * 导出数据
   */
  async exportData(): Promise<{
    messages: StoredMessage[];
    episodes: StoredEpisode[];
    toolStats: ToolStats[];
    sessions: SessionSummary[];
  }> {
    this.ensureInitialized();

    const messages = await withTransaction(this.db!, "messages", "readonly", async (stores) => {
      return getAllRecords<StoredMessage>(stores.messages);
    });

    const episodes = await withTransaction(this.db!, "episodes", "readonly", async (stores) => {
      return getAllRecords<StoredEpisode>(stores.episodes);
    });

    const toolStats = await withTransaction(this.db!, "toolStats", "readonly", async (stores) => {
      return getAllRecords<ToolStats>(stores.toolStats);
    });

    const sessions = await withTransaction(this.db!, "sessions", "readonly", async (stores) => {
      return getAllRecords<SessionSummary>(stores.sessions);
    });

    return { messages, episodes, toolStats, sessions };
  }

  /**
   * 导入数据
   */
  async importData(data: {
    messages?: StoredMessage[];
    episodes?: StoredEpisode[];
    toolStats?: ToolStats[];
    sessions?: SessionSummary[];
  }): Promise<{ imported: number; errors: number }> {
    this.ensureInitialized();

    let imported = 0;
    let errors = 0;

    if (data.messages) {
      for (const msg of data.messages) {
        try {
          await withTransaction(this.db!, "messages", "readwrite", async (stores) => {
            await addRecord(stores.messages, msg);
          });
          imported++;
        } catch {
          errors++;
        }
      }
    }

    if (data.episodes) {
      for (const ep of data.episodes) {
        try {
          await withTransaction(this.db!, "episodes", "readwrite", async (stores) => {
            await addRecord(stores.episodes, ep);
          });
          imported++;
        } catch {
          errors++;
        }
      }
    }

    if (data.toolStats) {
      for (const stats of data.toolStats) {
        try {
          await withTransaction(this.db!, "toolStats", "readwrite", async (stores) => {
            await addRecord(stores.toolStats, stats);
          });
          imported++;
        } catch {
          errors++;
        }
      }
    }

    if (data.sessions) {
      for (const session of data.sessions) {
        try {
          await withTransaction(this.db!, "sessions", "readwrite", async (stores) => {
            await addRecord(stores.sessions, session);
          });
          imported++;
        } catch {
          errors++;
        }
      }
    }

    console.log(`[PersistentMemory] 导入完成: ${imported} 成功, ${errors} 失败`);
    return { imported, errors };
  }
}

// ========== 工厂函数 ==========

/**
 * 创建持久化内存实例
 */
export async function createPersistentMemory(
  config?: PersistentMemoryConfig
): Promise<PersistentMemory> {
  const memory = new PersistentMemory(config);
  await memory.initialize();
  return memory;
}

export default PersistentMemory;
