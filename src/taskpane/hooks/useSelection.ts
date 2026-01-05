/**
 * useSelection - Excel 选区状态管理 Hook
 * @file src/taskpane/hooks/useSelection.ts
 * @description v2.9.8 管理 Excel 选区状态和数据读取
 */
import * as React from "react";
import type { SelectionResult, DataSummary, CellValue } from "../types/ui.types";

export interface UseSelectionReturn {
  /** 最后选区信息 */
  lastSelection: SelectionResult | null;
  /** 数据摘要 */
  dataSummary: DataSummary | null;
  /** 是否正在分析 */
  isAnalyzing: boolean;
  /** 分析进度 (0-100) */
  analysisProgress: number;
  /** 更新选区 */
  setSelection: (selection: SelectionResult | null) => void;
  /** 更新数据摘要 */
  setDataSummary: (summary: DataSummary | null) => void;
  /** 设置分析状态 */
  setAnalyzing: (analyzing: boolean) => void;
  /** 设置分析进度 */
  setAnalysisProgress: (progress: number) => void;
  /** 读取当前选区（需要 Excel context） */
  readSelection: () => Promise<SelectionResult | null>;
}

/**
 * 选区状态管理 Hook
 */
export function useSelection(): UseSelectionReturn {
  const [lastSelection, setLastSelection] = React.useState<SelectionResult | null>(null);
  const [dataSummary, setDataSummary] = React.useState<DataSummary | null>(null);
  const [isAnalyzing, setAnalyzing] = React.useState(false);
  const [analysisProgress, setAnalysisProgress] = React.useState(0);

  const setSelection = React.useCallback((selection: SelectionResult | null) => {
    setLastSelection(selection);
  }, []);

  const readSelection = React.useCallback(async (): Promise<SelectionResult | null> => {
    if (typeof Excel === "undefined" || !Excel.run) {
      return null;
    }

    try {
      return await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(["address", "values", "numberFormat", "rowCount", "columnCount"]);
        await ctx.sync();

        const result: SelectionResult = {
          address: range.address,
          values: range.values as CellValue[][],
          numberFormat: range.numberFormat,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
        };

        setLastSelection(result);
        return result;
      });
    } catch (error) {
      console.error("[useSelection] Error reading selection:", error);
      return null;
    }
  }, []);

  return {
    lastSelection,
    dataSummary,
    isAnalyzing,
    analysisProgress,
    setSelection,
    setDataSummary,
    setAnalyzing,
    setAnalysisProgress,
    readSelection,
  };
}

export default useSelection;
