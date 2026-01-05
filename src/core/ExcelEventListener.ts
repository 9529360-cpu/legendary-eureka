/**
 * ExcelEventListener - Excel 事件监听模块
 *
 * 功能：
 * 1. 监听选区变化事件
 * 2. 监听数据修改事件
 * 3. 监听工作表切换事件
 * 4. 提供事件回调机制
 *
 * 设计目标：实时感知 Excel 操作，让 AI 保持最新的工作簿状态
 */

/* eslint-disable no-undef */

/**
 * 事件类型枚举
 */
export enum ExcelEventType {
  SELECTION_CHANGED = "selectionChanged",
  DATA_CHANGED = "dataChanged",
  WORKSHEET_ACTIVATED = "worksheetActivated",
  WORKSHEET_ADDED = "worksheetAdded",
  WORKSHEET_DELETED = "worksheetDeleted",
  TABLE_CHANGED = "tableChanged",
  CALCULATION_COMPLETED = "calculationCompleted",
}

/**
 * 选区变化事件数据
 */
export interface SelectionChangedEventData {
  type: ExcelEventType.SELECTION_CHANGED;
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  timestamp: Date;
}

/**
 * 数据变化事件数据
 */
export interface DataChangedEventData {
  type: ExcelEventType.DATA_CHANGED;
  address: string;
  sheetName: string;
  changeType: "values" | "formulas" | "format";
  timestamp: Date;
}

/**
 * 工作表激活事件数据
 */
export interface WorksheetActivatedEventData {
  type: ExcelEventType.WORKSHEET_ACTIVATED;
  worksheetName: string;
  worksheetId: string;
  timestamp: Date;
}

/**
 * 工作表添加事件数据
 */
export interface WorksheetAddedEventData {
  type: ExcelEventType.WORKSHEET_ADDED;
  worksheetName: string;
  worksheetId: string;
  timestamp: Date;
}

/**
 * 工作表删除事件数据
 */
export interface WorksheetDeletedEventData {
  type: ExcelEventType.WORKSHEET_DELETED;
  worksheetId: string;
  timestamp: Date;
}

/**
 * 表格变化事件数据
 */
export interface TableChangedEventData {
  type: ExcelEventType.TABLE_CHANGED;
  tableName: string;
  sheetName: string;
  changeType: "add" | "delete" | "update";
  timestamp: Date;
}

/**
 * 所有事件数据类型
 */
export type ExcelEventData =
  | SelectionChangedEventData
  | DataChangedEventData
  | WorksheetActivatedEventData
  | WorksheetAddedEventData
  | WorksheetDeletedEventData
  | TableChangedEventData;

/**
 * 事件处理器类型
 */
export type EventHandler<T extends ExcelEventData = ExcelEventData> = (event: T) => void;

/**
 * 事件订阅配置
 */
export interface EventListenerConfig {
  enableSelectionTracking: boolean;
  enableDataTracking: boolean;
  enableWorksheetTracking: boolean;
  enableTableTracking: boolean;
  debounceMs: number;
  maxEventsBuffer: number;
}

/**
 * Excel 事件监听器类
 */
export class ExcelEventListener {
  private config: EventListenerConfig;
  private handlers: Map<ExcelEventType, Set<EventHandler>> = new Map();
  private eventBuffer: ExcelEventData[] = [];
  private registeredHandlers: Map<string, OfficeExtension.EventHandlerResult<any>> = new Map();
  private debounceTimers: Map<string, NodeJS.Timeout> = new Map();
  private isInitialized: boolean = false;

  constructor(config?: Partial<EventListenerConfig>) {
    this.config = {
      enableSelectionTracking: true,
      enableDataTracking: true,
      enableWorksheetTracking: true,
      enableTableTracking: false, // 表格追踪默认关闭，因为开销较大
      debounceMs: 100,
      maxEventsBuffer: 100,
      ...config,
    };

    // 初始化处理器映射
    Object.values(ExcelEventType).forEach((type) => {
      this.handlers.set(type, new Set());
    });
  }

  /**
   * 初始化事件监听（需要在 Excel.run 内部调用）
   */
  async initialize(context: Excel.RequestContext): Promise<void> {
    if (this.isInitialized) {
      console.warn("ExcelEventListener 已经初始化");
      return;
    }

    try {
      const workbook = context.workbook;

      // 1. 监听选区变化
      if (this.config.enableSelectionTracking) {
        await this.registerSelectionHandler(workbook, context);
      }

      // 2. 监听工作表事件
      if (this.config.enableWorksheetTracking) {
        await this.registerWorksheetHandlers(workbook, context);
      }

      // 3. 监听数据变化（需要在每个工作表上注册）
      if (this.config.enableDataTracking) {
        await this.registerDataHandlers(workbook, context);
      }

      this.isInitialized = true;
      console.log("ExcelEventListener 初始化完成");
    } catch (error) {
      console.error("ExcelEventListener 初始化失败:", error);
      throw error;
    }
  }

  /**
   * 订阅事件
   */
  on<T extends ExcelEventData>(type: ExcelEventType, handler: EventHandler<T>): void {
    const handlers = this.handlers.get(type);
    if (handlers) {
      handlers.add(handler as EventHandler);
    }
  }

  /**
   * 取消订阅事件
   */
  off<T extends ExcelEventData>(type: ExcelEventType, handler: EventHandler<T>): void {
    const handlers = this.handlers.get(type);
    if (handlers) {
      handlers.delete(handler as EventHandler);
    }
  }

  /**
   * 一次性订阅事件
   */
  once<T extends ExcelEventData>(type: ExcelEventType, handler: EventHandler<T>): void {
    const onceHandler: EventHandler<T> = (event) => {
      handler(event);
      this.off(type, onceHandler);
    };
    this.on(type, onceHandler);
  }

  /**
   * 获取事件缓冲区
   */
  getEventBuffer(): ExcelEventData[] {
    return [...this.eventBuffer];
  }

  /**
   * 清除事件缓冲区
   */
  clearEventBuffer(): void {
    this.eventBuffer = [];
  }

  /**
   * 获取最近的事件
   */
  getRecentEvents(count: number = 10): ExcelEventData[] {
    return this.eventBuffer.slice(-count);
  }

  /**
   * 销毁监听器
   */
  async dispose(context: Excel.RequestContext): Promise<void> {
    // 移除所有已注册的 Office.js 事件处理器
    for (const [key, handler] of this.registeredHandlers) {
      try {
        handler.remove();
        console.log(`移除事件处理器: ${key}`);
      } catch (error) {
        console.error(`移除事件处理器失败 ${key}:`, error);
      }
    }
    this.registeredHandlers.clear();

    // 清除防抖定时器
    for (const timer of this.debounceTimers.values()) {
      clearTimeout(timer);
    }
    this.debounceTimers.clear();

    // 清除处理器
    this.handlers.forEach((set) => set.clear());

    await context.sync();
    this.isInitialized = false;
    console.log("ExcelEventListener 已销毁");
  }

  // ==================== 私有方法 ====================

  private async registerSelectionHandler(
    workbook: Excel.Workbook,
    _context: Excel.RequestContext
  ): Promise<void> {
    const handler = workbook.onSelectionChanged.add(async (_args) => {
      await Excel.run(async (ctx) => {
        try {
          const selection = ctx.workbook.getSelectedRange();
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();

          selection.load("address, rowCount, columnCount");
          sheet.load("name");

          await ctx.sync();

          /* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
          this.emitEventDebounced(ExcelEventType.SELECTION_CHANGED, {
            type: ExcelEventType.SELECTION_CHANGED,
            address: selection.address,
            sheetName: sheet.name,
            rowCount: selection.rowCount,
            columnCount: selection.columnCount,
            timestamp: new Date(),
          });
          /* eslint-enable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
        } catch (error) {
          console.error("处理选区变化事件失败:", error);
        }
      });
    });

    await _context.sync();
    this.registeredHandlers.set("selection", handler);
  }

  private async registerWorksheetHandlers(
    workbook: Excel.Workbook,
    context: Excel.RequestContext
  ): Promise<void> {
    // 工作表激活事件
    const activatedHandler = workbook.worksheets.onActivated.add(async (args) => {
      await Excel.run(async (ctx) => {
        try {
          const sheet = ctx.workbook.worksheets.getItem(args.worksheetId);
          sheet.load("name, id");
          await ctx.sync();

          /* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
          this.emitEvent({
            type: ExcelEventType.WORKSHEET_ACTIVATED,
            worksheetName: sheet.name,
            worksheetId: sheet.id,
            timestamp: new Date(),
          });
          /* eslint-enable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
        } catch (error) {
          console.error("处理工作表激活事件失败:", error);
        }
      });
    });

    // 工作表添加事件
    const addedHandler = workbook.worksheets.onAdded.add(async (args) => {
      await Excel.run(async (ctx) => {
        try {
          const sheet = ctx.workbook.worksheets.getItem(args.worksheetId);
          sheet.load("name, id");
          await ctx.sync();

          /* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
          this.emitEvent({
            type: ExcelEventType.WORKSHEET_ADDED,
            worksheetName: sheet.name,
            worksheetId: sheet.id,
            timestamp: new Date(),
          });
          /* eslint-enable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
        } catch (error) {
          console.error("处理工作表添加事件失败:", error);
        }
      });
    });

    // 工作表删除事件
    const deletedHandler = workbook.worksheets.onDeleted.add(async (_args) => {
      this.emitEvent({
        type: ExcelEventType.WORKSHEET_DELETED,
        worksheetId: _args.worksheetId,
        timestamp: new Date(),
      });
    });

    await context.sync();
    this.registeredHandlers.set("worksheetActivated", activatedHandler);
    this.registeredHandlers.set("worksheetAdded", addedHandler);
    this.registeredHandlers.set("worksheetDeleted", deletedHandler);
  }

  private async registerDataHandlers(
    workbook: Excel.Workbook,
    context: Excel.RequestContext
  ): Promise<void> {
    // 获取所有工作表
    const sheets = workbook.worksheets;
    sheets.load("items");
    await context.sync();

    // 为每个工作表注册数据变化事件
    /* eslint-disable office-addins/no-context-sync-in-loop */
    for (let i = 0; i < sheets.items.length; i++) {
      const sheet = sheets.items[i];
      sheet.load("name, id");
      await context.sync();

      const sheetName = sheet.name;
      const sheetId = sheet.id;

      const handler = sheet.onChanged.add(async (args) => {
        this.emitEventDebounced(ExcelEventType.DATA_CHANGED, {
          type: ExcelEventType.DATA_CHANGED,
          address: args.address,
          sheetName: sheetName,
          changeType: this.mapChangeType(args.changeType),
          timestamp: new Date(),
        });
      });

      await context.sync();
      this.registeredHandlers.set(`data_${sheetId}`, handler);
    }
    /* eslint-enable office-addins/no-context-sync-in-loop */
  }

  private mapChangeType(changeType: string): "values" | "formulas" | "format" {
    switch (changeType) {
      case "RangeEdited":
      case "RowInserted":
      case "RowDeleted":
      case "ColumnInserted":
      case "ColumnDeleted":
        return "values";
      case "CellInserted":
      case "CellDeleted":
        return "values";
      default:
        return "values";
    }
  }

  private emitEvent(event: ExcelEventData): void {
    // 添加到缓冲区
    this.eventBuffer.push(event);
    if (this.eventBuffer.length > this.config.maxEventsBuffer) {
      this.eventBuffer = this.eventBuffer.slice(-this.config.maxEventsBuffer);
    }

    // 触发处理器
    const handlers = this.handlers.get(event.type);
    if (handlers) {
      handlers.forEach((handler) => {
        try {
          handler(event);
        } catch (error) {
          console.error(`事件处理器错误 (${event.type}):`, error);
        }
      });
    }
  }

  private emitEventDebounced(type: ExcelEventType, event: ExcelEventData): void {
    const key = `${type}_${event.type === ExcelEventType.DATA_CHANGED ? (event as DataChangedEventData).address : ""}`;

    // 清除之前的定时器
    const existingTimer = this.debounceTimers.get(key);
    if (existingTimer) {
      clearTimeout(existingTimer);
    }

    // 设置新的定时器
    const timer = setTimeout(() => {
      this.emitEvent(event);
      this.debounceTimers.delete(key);
    }, this.config.debounceMs);

    this.debounceTimers.set(key, timer);
  }
}

/**
 * 创建事件监听器的工厂函数
 */
export function createExcelEventListener(
  config?: Partial<EventListenerConfig>
): ExcelEventListener {
  return new ExcelEventListener(config);
}
