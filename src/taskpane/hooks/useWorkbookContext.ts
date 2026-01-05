/**
 * useWorkbookContext - 工作簿上下文 Hook
 * @file src/taskpane/hooks/useWorkbookContext.ts
 * @description v2.9.12 订阅工作簿结构信息，扫描逻辑委托给 ExcelScanner service
 */
import * as React from "react";
import type { WorkbookContext } from "../types/ui.types";
import { scanWorkbook as scanWorkbookService } from "../../services/ExcelScanner";

export interface UseWorkbookContextReturn {
  /** 工作簿上下文 */
  workbookContext: WorkbookContext | null;
  /** 是否正在扫描 */
  isScanning: boolean;
  /** 扫描进度 (0-100) */
  scanProgress: number;
  /** 手动触发扫描 */
  scanWorkbook: () => Promise<WorkbookContext | null>;
  /** 获取工作簿摘要（用于 HeaderBar） */
  getWorkbookSummary: () =>
    | {
        sheetCount: number;
        tableCount: number;
        formulaCount: number;
        qualityScore: number;
      }
    | undefined;
}

/**
 * 工作簿上下文 Hook
 * 只负责订阅状态和触发 refresh，扫描逻辑在 ExcelScanner service
 */
export function useWorkbookContext(): UseWorkbookContextReturn {
  const [workbookContext, setWorkbookContext] = React.useState<WorkbookContext | null>(null);
  const [isScanning, setScanning] = React.useState(false);
  const [scanProgress, setScanProgress] = React.useState(0);
  const scanIntervalRef = React.useRef<ReturnType<typeof setInterval> | null>(null);

  // 使用 ref 存储状态，避免 scanWorkbook 依赖变化
  const stateRef = React.useRef({ isScanning, workbookContext });
  stateRef.current = { isScanning, workbookContext };

  const scanWorkbook = React.useCallback(async (): Promise<WorkbookContext | null> => {
    if (stateRef.current.isScanning) return stateRef.current.workbookContext;

    setScanning(true);
    setScanProgress(0);

    try {
      const context = await scanWorkbookService((progress) => {
        setScanProgress(progress.progress);
      });

      if (context) {
        setWorkbookContext(context);
      }
      return context;
    } finally {
      setScanning(false);
      setScanProgress(0);
    }
  }, []); // 空依赖，函数稳定

  // 初次加载时扫描，并定期刷新
  React.useEffect(() => {
    void scanWorkbook();

    // 定期刷新（每5分钟）
    scanIntervalRef.current = setInterval(
      () => {
        void scanWorkbook();
      },
      5 * 60 * 1000
    );

    return () => {
      if (scanIntervalRef.current) {
        clearInterval(scanIntervalRef.current);
      }
    };
  }, []); // 只在 mount 时执行

  const getWorkbookSummary = React.useCallback(() => {
    if (!workbookContext) return undefined;
    return {
      sheetCount: workbookContext.sheets.length,
      tableCount: workbookContext.tables.length,
      formulaCount: workbookContext.totalFormulas,
      qualityScore: workbookContext.overallQualityScore,
    };
  }, [workbookContext]);

  return {
    workbookContext,
    isScanning,
    scanProgress,
    scanWorkbook,
    getWorkbookSummary,
  };
}

export default useWorkbookContext;
