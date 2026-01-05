/**
 * useUndoStack - 撤销栈管理 Hook
 * @file src/taskpane/hooks/useUndoStack.ts
 * @description v2.9.12 管理撤销操作，保存和恢复 Excel 状态
 */
import * as React from "react";
import type { UndoStackItem, CellValue, OperationHistoryItem } from "../types";
import { uid } from "../utils";

export interface UseUndoStackOptions {
  /** 最大撤销栈大小 */
  maxStackSize?: number;
  /** Toast 提示回调 */
  showToast?: (message: string, intent: "success" | "warning" | "error") => void;
  /** 添加到历史记录的回调 */
  addToHistory?: (item: OperationHistoryItem) => void;
}

export interface UseUndoStackReturn {
  /** 撤销栈 */
  undoStack: UndoStackItem[];
  /** 撤销栈长度 */
  undoCount: number;
  /** 保存当前状态用于撤销 */
  saveStateForUndo: (operation: string, rangeAddress?: string) => Promise<UndoStackItem | null>;
  /** 执行撤销 */
  performUndo: () => Promise<boolean>;
  /** 添加到撤销栈 */
  addToUndoStack: (item: UndoStackItem) => void;
  /** 清空撤销栈 */
  clearUndoStack: () => void;
}

/**
 * 撤销栈管理 Hook
 */
export function useUndoStack(options: UseUndoStackOptions = {}): UseUndoStackReturn {
  const { maxStackSize = 10, showToast, addToHistory } = options;

  const [undoStack, setUndoStack] = React.useState<UndoStackItem[]>([]);

  // 使用 ref 存储回调和状态，避免 performUndo 依赖变化
  const callbacksRef = React.useRef({ showToast, addToHistory });
  callbacksRef.current = { showToast, addToHistory };

  const undoStackRef = React.useRef(undoStack);
  undoStackRef.current = undoStack;

  /**
   * 保存操作前的状态，用于撤销
   */
  const saveStateForUndo = React.useCallback(
    async (operation: string, rangeAddress?: string): Promise<UndoStackItem | null> => {
      if (typeof Excel === "undefined") {
        return null;
      }

      try {
        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.load("name");

          // 如果没有指定范围，获取已使用范围
          let targetRange: Excel.Range;
          if (rangeAddress) {
            targetRange = sheet.getRange(rangeAddress);
          } else {
            targetRange = sheet.getUsedRange();
          }

          targetRange.load("address,values,formulas");
          await context.sync();

          const undoItem: UndoStackItem = {
            id: uid(),
            operation,
            timestamp: new Date(),
            sheetName: sheet.name,
            rangeAddress: targetRange.address,
            previousValues: targetRange.values as CellValue[][],
            previousFormulas: targetRange.formulas as string[][],
          };

          return undoItem;
        });
      } catch (error) {
        console.warn("无法保存撤销状态:", error);
        return null;
      }
    },
    []
  );

  /**
   * 执行撤销操作
   */
  const performUndo = React.useCallback(async (): Promise<boolean> => {
    const currentStack = undoStackRef.current;
    const { showToast: toast, addToHistory: history } = callbacksRef.current;

    if (currentStack.length === 0) {
      toast?.("没有可撤销的操作", "warning");
      return false;
    }

    const lastUndo = currentStack[currentStack.length - 1];

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(lastUndo.sheetName);
        const range = sheet.getRange(lastUndo.rangeAddress);

        // 优先恢复公式（若有），否则恢复值
        if (lastUndo.previousFormulas) {
          range.formulas = lastUndo.previousFormulas;
        } else {
          range.values = lastUndo.previousValues;
        }

        await context.sync();
      });

      // 从栈中移除已撤销的操作
      setUndoStack((prev) => prev.slice(0, -1));

      toast?.(`已撤销: ${lastUndo.operation}`, "success");

      history?.({
        id: uid(),
        operation: "撤销",
        timestamp: new Date(),
        success: true,
        details: `撤销了 ${lastUndo.operation}`,
      });

      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      toast?.(`撤销失败: ${message}`, "error");
      return false;
    }
  }, []); // 空依赖，函数稳定

  /**
   * 添加到撤销栈
   */
  const addToUndoStack = React.useCallback(
    (item: UndoStackItem) => {
      setUndoStack((prev) => {
        const newStack = [...prev, item];
        // 限制栈大小
        if (newStack.length > maxStackSize) {
          return newStack.slice(-maxStackSize);
        }
        return newStack;
      });
    },
    [maxStackSize]
  );

  /**
   * 清空撤销栈
   */
  const clearUndoStack = React.useCallback(() => {
    setUndoStack([]);
  }, []);

  return {
    undoStack,
    undoCount: undoStack.length,
    saveStateForUndo,
    performUndo,
    addToUndoStack,
    clearUndoStack,
  };
}

export default useUndoStack;
