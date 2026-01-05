/**
 * useSelectionListener - Excel 选区监听 Hook
 * @file src/taskpane/hooks/useSelectionListener.ts
 * @description v2.9.12 监听 Excel 选区变化，触发主动式分析
 */
import * as React from "react";
import type { SelectionResult, DataSummary, ProactiveSuggestion } from "../types";
import { readSelection } from "../utils";
import { generateDataSummary, generateProactiveSuggestions } from "../utils/dataAnalysis";

export interface UseSelectionListenerOptions {
  /** 是否启用自动分析 */
  autoAnalyzeEnabled: boolean;
  /** 是否正在忙碌（发送消息中） */
  busy: boolean;
  /** 发送消息的回调 */
  onSend: (text: string) => Promise<void>;
  /** 防抖延迟（毫秒） */
  debounceMs?: number;
}

export interface UseSelectionListenerReturn {
  /** 最后一次选区 */
  lastSelection: SelectionResult | null;
  /** 数据摘要 */
  dataSummary: DataSummary | null;
  /** 主动建议列表 */
  proactiveSuggestions: ProactiveSuggestion[];
  /** 是否正在分析 */
  isAnalyzing: boolean;
  /** 分析进度 (0-100) */
  analysisProgress: number;
  /** 手动触发分析 */
  triggerAnalysis: () => Promise<void>;
}

/**
 * Excel 选区监听 Hook
 * 监听选区变化，自动触发数据分析和建议生成
 */
export function useSelectionListener(
  options: UseSelectionListenerOptions
): UseSelectionListenerReturn {
  const { autoAnalyzeEnabled, busy, onSend, debounceMs = 500 } = options;

  const [lastSelection, setLastSelection] = React.useState<SelectionResult | null>(null);
  const [dataSummary, setDataSummary] = React.useState<DataSummary | null>(null);
  const [proactiveSuggestions, setProactiveSuggestions] = React.useState<ProactiveSuggestion[]>([]);
  const [isAnalyzing, setIsAnalyzing] = React.useState(false);
  const [analysisProgress, setAnalysisProgress] = React.useState(0);

  const debounceRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const onSendRef = React.useRef(onSend);

  // 使用 ref 存储可变状态，避免 handleSelectionChanged 依赖变化导致重新注册事件
  const stateRef = React.useRef({ autoAnalyzeEnabled, busy, isAnalyzing });
  React.useEffect(() => {
    stateRef.current = { autoAnalyzeEnabled, busy, isAnalyzing };
  }, [autoAnalyzeEnabled, busy, isAnalyzing]);

  // 保持 onSend 引用最新
  React.useEffect(() => {
    onSendRef.current = onSend;
  }, [onSend]);

  /**
   * 执行主动式分析
   */
  const performProactiveAnalysis = React.useCallback(async (): Promise<void> => {
    if (typeof Excel === "undefined") {
      return;
    }

    setIsAnalyzing(true);
    setAnalysisProgress(10);

    try {
      const selection = await readSelection();
      setAnalysisProgress(40);

      // 如果选区太小（少于2行或1列），不进行分析
      if (selection.rowCount < 2 || selection.columnCount < 1) {
        setDataSummary(null);
        setProactiveSuggestions([]);
        setIsAnalyzing(false);
        setAnalysisProgress(0);
        return;
      }

      setLastSelection(selection);
      setAnalysisProgress(60);

      // 生成数据摘要
      const summary = generateDataSummary(selection);
      setDataSummary(summary);
      setAnalysisProgress(80);

      // 生成主动建议
      const suggestions = generateProactiveSuggestions(selection, summary, onSendRef.current);
      setProactiveSuggestions(suggestions);

      setAnalysisProgress(100);
      setTimeout(() => setAnalysisProgress(0), 1000);
    } catch (error) {
      console.warn("主动分析失败:", error);
    } finally {
      setIsAnalyzing(false);
    }
  }, []);

  /**
   * 处理选区变化事件
   */
  const handleSelectionChanged = React.useCallback(
    async (_event: Excel.SelectionChangedEventArgs): Promise<void> => {
      // 防抖处理，避免频繁触发
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }

      debounceRef.current = setTimeout(() => {
        const { autoAnalyzeEnabled: auto, busy: isBusy, isAnalyzing: analyzing } = stateRef.current;
        if (auto && !isBusy && !analyzing) {
          void performProactiveAnalysis();
        }
      }, debounceMs);
    },
    [debounceMs, performProactiveAnalysis]
  );

  // 设置选区监听
  React.useEffect(() => {
    if (typeof Excel === "undefined") {
      return;
    }

    // eslint-disable-next-line no-undef
    let eventHandler: OfficeExtension.EventHandlerResult<Excel.SelectionChangedEventArgs> | null =
      null;

    const setupSelectionListener = async () => {
      try {
        await Excel.run(async (context) => {
          const workbook = context.workbook;
          // workbook.onSelectionChanged 在测试环境中可能不存在
          if (workbook && workbook.onSelectionChanged && typeof workbook.onSelectionChanged.add === "function") {
            eventHandler = workbook.onSelectionChanged.add(handleSelectionChanged);
          }
          await context.sync();
        });
      } catch (error) {
        console.warn("无法设置选区监听:", error);
      }
    };

    void setupSelectionListener();

    return () => {
      // 清理事件监听
      if (eventHandler) {
        void Excel.run(async (context) => {
          try {
            eventHandler?.remove();
          } catch {}
          await context.sync();
        });
      }
      if (debounceRef.current) {
        clearTimeout(debounceRef.current);
      }
    };
  }, [handleSelectionChanged]);

  return {
    lastSelection,
    dataSummary,
    proactiveSuggestions,
    isAnalyzing,
    analysisProgress,
    triggerAnalysis: performProactiveAnalysis,
  };
}

export default useSelectionListener;
