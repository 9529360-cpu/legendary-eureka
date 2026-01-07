﻿import * as React from "react";
import {
  FluentProvider,
  webDarkTheme,
  webLightTheme,
  makeStyles,
  shorthands,
  tokens,
  Caption1,
  Toaster,
  useId,
  useToastController,
  Text,
  ProgressBar,
} from "@fluentui/react-components";
// v2.9.8: 图标已移至各子组件，此处不再需要
import { ErrorHandler } from "../../core/ErrorHandler";
// v2.9.8: ApiService 调用已移至 hooks，仅保留类型导入
import type { ApiKeyStatus } from "../../services/ApiService";
import { DataAnalyzer } from "../../core/DataAnalyzer";
// v2.9.8: Agent 初始化已移至 useAgent hook
// v2.9.8: 新增组件导入
import { ChatInputArea } from "./ChatInputArea";
import { ApiConfigDialog } from "./ApiConfigDialog";
import { PreviewConfirmDialog } from "./PreviewConfirmDialog";
import { ApprovalDialog } from "./ApprovalDialog";
import { HeaderBar } from "./HeaderBar";
import { MessageList } from "./MessageList";
import { WelcomeView } from "./WelcomeView";
import { InsightPanel } from "./InsightPanel";
import type { DataSummary as InsightDataSummary, ProactiveSuggestion as InsightSuggestion } from "./InsightPanel";

// v2.9.8: 模块化重构 - 导入抽取的类型、工具函数和 hooks
import { 
  loadUserPreferences, 
  saveUserPreferences,
  // Excel 辅助函数
  uid,
} from "../utils";
import { useApiSettings, useAgentV4, useWorkbookContext, useSelectionListener, useUndoStack, useProactiveAgent } from "../hooks";
import type {
  CopilotAction,
  ChatMessage,
  DataInsight,
  OperationHistoryItem,
  UserPreferences,
  AgentThought,
  AgentPlanUI as AgentPlan,
} from "../types";

// v2.9.8: 类型定义和工具函数已移至 ../types 和 ../utils



// ========== Styles ==========
// v2.9.8: 大部分样式已移至各子组件（HeaderBar, MessageList, WelcomeView, InsightPanel 等）

const useStyles = makeStyles({
  // ========== 根容器 ==========
  app: {
    width: "100%",
    height: "100vh",
    margin: 0,
    padding: 0,
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    fontFamily: "Segoe UI, -apple-system, BlinkMacSystemFont, sans-serif",
  },
  
  container: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    height: "100%",
    margin: 0,
    padding: 0,
    overflow: "hidden",
  },

  // ========== 对话区域 ==========
  chatContainer: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },

  // ========== 进度条 ==========
  progressWrapper: {
    ...shorthands.padding("8px", "16px"),
  },
  
  // v2.9.17: Agent 进度显示
  agentProgressWrapper: {
    ...shorthands.padding("8px", "16px"),
    display: "flex",
    alignItems: "center",
    ...shorthands.gap("8px"),
  },
  agentProgressBar: {
    flex: 1,
  },
  agentProgressText: {
    whiteSpace: "nowrap",
  },
  
  // v4.3: 主动洞察快速操作区域
  proactiveActionsWrapper: {
    ...shorthands.padding("8px", "16px"),
    display: "flex",
    flexWrap: "wrap",
    ...shorthands.gap("8px"),
    ...shorthands.borderBottom("1px", "solid", tokens.colorNeutralStroke2),
  },
  proactiveActionsTitle: {
    width: "100%",
    marginBottom: "4px",
  },
  proactiveActionButton: {
    ...shorthands.padding("4px", "12px"),
    fontSize: "12px",
    ...shorthands.borderRadius("16px"),
    ...shorthands.border("1px", "solid", tokens.colorBrandStroke1),
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorBrandForeground1,
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorBrandBackground2,
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
  proactiveExecuteAllButton: {
    ...shorthands.padding("4px", "12px"),
    fontSize: "12px",
    ...shorthands.borderRadius("16px"),
    ...shorthands.border("1px", "solid", tokens.colorPaletteGreenBorder1),
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorPaletteGreenBackground2,
    },
    "&:disabled": {
      opacity: 0.5,
      cursor: "not-allowed",
    },
  },
});

// v2.9.9: applyExcelCommand 和 applyAction 已删除
// 所有 Excel 操作现在通过 Agent 工具层执行 (src/agent/ExcelAdapter.ts)
// UI 只负责展示，不直接调用 Excel API

const App: React.FC = () => {
  const styles = useStyles();
  const errorHandler = ErrorHandler.getInstance();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  // v2.9.8: 使用 useApiSettings hook 管理后端连接和 API 密钥
  const apiSettings = useApiSettings();
  // v2.9.9: legacyChat 已删除，所有请求统一走 Agent
  // v4.0: 使用 useAgentV4 hook 管理 Agent 初始化和调用（新架构）
  const agent = useAgentV4({
    maxIterations: 30,
    enableMemory: true,
    verboseLogging: true,
  });

  // v4.3: 主动洞察型 Agent - 观察 → 判断 → 建议 → 确认
  const proactiveAgent = useProactiveAgent({
    autoAnalyzeOnStart: true,
    autoAnalyzeOnSheetChange: true,
    verboseLogging: true,
  });

  const [_dataAnalyzer] = React.useState(() => new DataAnalyzer());
  
  // ���ر������û�ƫ��
  const [userPreferences, setUserPreferences] = React.useState<UserPreferences>(() => loadUserPreferences());

  const [isDarkTheme, _setIsDarkTheme] = React.useState(userPreferences.theme === "dark");
  const [_showSidebar, _setShowSidebar] = React.useState(false);
  const [_showAnalysisPanel, _setShowAnalysisPanel] = React.useState(false);

  const [messages, setMessages] = React.useState<ChatMessage[]>([
    {
      id: uid(),
      role: "assistant",
      text: "你好！👋 欢迎使用 Excel 智能助手。\n\n选中 Excel 中的数据，然后告诉我你想做什么，比如：\n• 「分析这些数据」\n• 「求和」\n• 「创建图表」",
      timestamp: new Date(),
    },
  ]);
  const [input, setInput] = React.useState("");
  const [busy, setBusy] = React.useState(false);
  // v2.9.9: streaming 状态变量已删除（统一走 Agent 模式）

  const [_insights, _setInsights] = React.useState<DataInsight[]>([]);

  const [_history, setHistory] = React.useState<OperationHistoryItem[]>([]);
  const [_historyIndex, setHistoryIndex] = React.useState(-1);

  // ����Ԥ����ȷ��
  const [previewDialogOpen, setPreviewDialogOpen] = React.useState(false);
  const [_pendingActions, setPendingActions] = React.useState<CopilotAction[]>([]);
  const [previewMessage, setPreviewMessage] = React.useState("");
  const [_requireConfirmation, _setRequireConfirmation] = React.useState(true);

  // v2.9.8: API 状态现在由 useApiSettings hook 管理
  // 从 hook 解构出需要的状态（保持向后兼容的变量名）
  const { 
    backendHealthy, 
    backendChecking, 
    backendError: _backendError,
    apiKeyStatus,
    apiKeyBusy,
  } = apiSettings;
  
  // UI 专属状态（dialog 开关、输入框等）保持在组件内
  const [apiKeyDialogOpen, setApiKeyDialogOpen] = React.useState(false);
  const [apiKeyInput, setApiKeyInput] = React.useState("");

  // ===== ����ʽ Copilot ״̬ =====
  const [autoAnalyzeEnabled, _setAutoAnalyzeEnabled] = React.useState(userPreferences.autoAnalyze);

  // v2.9.12: 使用 useWorkbookContext hook 管理工作簿扫描
  const {
    workbookContext,
    isScanning,
    scanProgress,
    scanWorkbook,
    getWorkbookSummary: _getWorkbookSummary,
  } = useWorkbookContext();

  // 用于 useSelectionListener 的 onSend ref（避免循环依赖）
  const onSendRef = React.useRef<(text: string) => Promise<void>>(async () => {});

  // v2.9.12: 使用 useUndoStack hook 管理撤销
  const {
    undoStack,
    undoCount: _undoCount,
    saveStateForUndo: _saveStateForUndo,
    performUndo,
    addToUndoStack: _addToUndoStack,
  } = useUndoStack({
    maxStackSize: 10,
    showToast: (message, intent) => {
      dispatchToast(<Text>{message}</Text>, { intent });
    },
    addToHistory: (item) => {
      setHistory((prev) => [item, ...prev].slice(0, 50));
    },
  });

  // v2.9.12: 使用 useSelectionListener hook 管理选区监听
  const {
    lastSelection,
    dataSummary,
    proactiveSuggestions,
    isAnalyzing,
    analysisProgress,
  } = useSelectionListener({
    autoAnalyzeEnabled,
    busy,
    onSend: (text) => onSendRef.current(text),
  });

  // ===== Agent Loop 状态 (用于UI显示思维链) =====
  const [_agentPlan, _setAgentPlan] = React.useState<AgentPlan | null>(null);
  const [_agentThoughts, setAgentThoughts] = React.useState<AgentThought[]>([]);
  const [isAgentRunning, setIsAgentRunning] = React.useState(false);
  
  // ===== 稳定的回调函数（避免 HeaderBar 重渲染） =====
  const handleRefreshWorkbook = React.useCallback(() => {
    void scanWorkbook();
  }, []); // scanWorkbook 现在是稳定的
  
  const handleUndo = React.useCallback(() => {
    void performUndo();
  }, []); // performUndo 现在是稳定的
  
  const handleOpenSettings = React.useCallback(() => {
    setApiKeyDialogOpen(true);
  }, []);
  
  // 用于实时更新 Agent 思考过程的 Ref
  const currentAgentMsgIdRef = React.useRef<string | null>(null);
  const agentStepsRef = React.useRef<string[]>([]);

  // v2.9.8: Agent 初始化已移至 useAgent hook
  // 保留 updateAgentMessageRef 用于 UI 更新
  const updateAgentMessageRef = React.useRef<((step: string) => void) | null>(null);
  
  // 记录操作用于学习
  function _recordOperation(operation: string) {
    setUserPreferences((prev) => {
      const lastUsed = prev.lastUsedOperations.filter(
        (op) => op.operation !== operation
      );
      lastUsed.unshift({ operation, timestamp: Date.now() });
      // 只保留最近20条
      const trimmed = lastUsed.slice(0, 20);
      const newPrefs = { ...prev, lastUsedOperations: trimmed };
      saveUserPreferences(newPrefs);
      return newPrefs;
    });
  }

  React.useEffect(() => {
    void bootstrapBackendStatus();
  }, []);

  // v4.3: 主动洞察 Agent 消息同步
  // 当 proactiveAgent 产生新消息时，添加到聊天列表
  React.useEffect(() => {
    const latestMessage = proactiveAgent.latestMessage;
    if (latestMessage) {
      // 只添加 agent 类型的消息（洞察和建议）
      if (latestMessage.type === "insight" || latestMessage.type === "suggestion") {
        setMessages((prev) => {
          // 避免重复添加
          const alreadyExists = prev.some((m) => m.id === latestMessage.id);
          if (alreadyExists) return prev;
          
          return [
            ...prev,
            {
              id: latestMessage.id,
              role: "assistant" as const,
              text: latestMessage.content,
              timestamp: latestMessage.timestamp,
            },
          ];
        });
      }
    }
  }, [proactiveAgent.latestMessage]);

  // v4.3: 当有新的洞察报告时显示
  React.useEffect(() => {
    if (proactiveAgent.insights && proactiveAgent.insights.narrativeDescription) {
      // 检查是否已经显示过这个洞察
      const insightId = `insight-${Date.now()}`;
      setMessages((prev) => {
        // 如果最后一条消息已经是洞察消息，不重复添加
        const lastMsg = prev[prev.length - 1];
        if (lastMsg?.text === proactiveAgent.insights?.narrativeDescription) return prev;
        
        return [
          ...prev,
          {
            id: insightId,
            role: "assistant" as const,
            text: proactiveAgent.insights!.narrativeDescription,
            timestamp: new Date(),
          },
        ];
      });
    }
  }, [proactiveAgent.insights?.narrativeDescription]);

  // v2.9.12: 选区监听已移至 useSelectionListener hook
  // v2.9.12: handleSelectionChanged 和 performProactiveAnalysis 已移至 useSelectionListener hook

  // v2.9.12: Moved to separate modules
  // - parseFormulaReferences, analyzeFormulaComplexity: utils/dataAnalysis.ts
  // - scanWorkbook, verifyOperationResult: services/ExcelScanner.ts
  // - generateDataSummary, generateProactiveSuggestions: utils/dataAnalysis.ts
  // - Workbook scan useEffect: useWorkbookContext hook

  function getActionLabel(action: CopilotAction): string {
    switch (action.type) {
      case "executeCommand":
        return action.label;
      case "writeRange":
      case "setFormula":
      case "writeCell":
        return action.address;
      default:
        // ȷ��������������
        return "δ֪����";
    }
  }

  /**
   * �жϲ����Ƿ���Ҫȷ�ϣ��߷��ղ�����
   */
  function isHighRiskAction(action: CopilotAction): boolean {
    if (action.type === "executeCommand") {
      const cmd = action.command;
      // �߷��ղ�����������ɾ������ʽ������Χ��
      const highRiskActions = ["clear", "delete", "format"];
      if (highRiskActions.some((risk) => cmd.action?.toLowerCase().includes(risk))) {
        return true;
      }
      // �漰����Χ����д��
      if (cmd.type === "write" && cmd.parameters?.values) {
        const values = cmd.parameters.values;
        if (Array.isArray(values) && values.length > 50) {
          return true;
        }
      }
    }
    if (action.type === "writeRange") {
      if (action.values.length > 50) {
        return true;
      }
    }
    return false;
  }

  /**
   * 生成操作预览描述
   */
  function _generatePreviewDescription(actions: CopilotAction[]): string {
    const descriptions = actions.map((action) => {
      const label = getActionLabel(action);
      const risk = isHighRiskAction(action) ? " ??" : "";
      
      if (action.type === "writeRange") {
        return `?? д�����ݵ� ${action.address}��${action.values.length}�� �� ${action.values[0]?.length || 0}�У�${risk}`;
      }
      if (action.type === "setFormula") {
        return `?? ���ù�ʽ ${action.formula} �� ${action.address}`;
      }
      if (action.type === "writeCell") {
        return `?? д�� "${action.value}" �� ${action.address}`;
      }
      if (action.type === "executeCommand") {
        const cmd = action.command;
        return `?? ִ�� ${cmd.type}/${cmd.action}: ${label}${risk}`;
      }
      return `? ${label}${risk}`;
    });

    return descriptions.join("\n");
  }

  // v2.9.12: saveStateForUndo, performUndo, addToUndoStack 已移至 useUndoStack hook

  /**
   * 请求操作确认
   */
  async function _requestOperationConfirmation(
    actions: CopilotAction[],
    message: string
  ): Promise<boolean> {
    return new Promise((resolve) => {
      setPendingActions(actions);
      setPreviewMessage(message);
      setPreviewDialogOpen(true);
      
      // ʹ��һ�����صĻص�����
      const handler = (confirmed: boolean) => {
        setPreviewDialogOpen(false);
        setPendingActions([]);
        setPreviewMessage("");
        resolve(confirmed);
      };
      
      // �洢�ص��Թ��Ի���ʹ��
      (window as unknown as Record<string, unknown>)._confirmHandler = handler;
    });
  }

  // v2.9.9: applyActionsAutomatically 和 attemptErrorRecovery 已删除
  // 所有 Excel 操作现在通过 Agent 工具层执行 (src/agent/ExcelAdapter.ts)

  // ===== ReAct Agent 核心系统 =====
  // Agent-First 架构: Agent 核心已迁移到 src/agent/ 模块
  // 这里只保留 UI 相关的辅助函数
  
  /**
   * 添加 Agent 思维记录（UI 展示用）
   */
  function _addAgentThought(type: AgentThought["type"], content: string) {
    const thought: AgentThought = {
      id: uid(),
      type,
      content,
      timestamp: new Date(),
    };
    setAgentThoughts(prev => [...prev, thought]);
    return thought;
  }
  
  /**
   * 实时更新 Agent 消息（流式显示思考过程）
   */
  function _updateAgentMessage(step: string) {
    const msgId = currentAgentMsgIdRef.current;
    if (!msgId) return;
    
    // 添加新步骤到历史
    agentStepsRef.current.push(step);
    
    // 构建实时显示的消息内容
    const stepsText = agentStepsRef.current.join("\n");
    
    setMessages((prev) =>
      prev.map((msg) =>
        msg.id === msgId
          ? { ...msg, text: stepsText }
          : msg
      )
    );
  }

  // ===== 已删除旧的独立 Agent 函数（已迁移到 src/agent/ 模块）=====
  // buildAgentSystemPrompt, agentThink, agentExecuteTool, runReActAgent, runAgentLoop, getCurrentSelection
  // v2.9.8: 现在通过 useAgent hook 使用新的 Agent 模块

  // v2.9.8: 后端/API密钥管理函数现在使用 useApiSettings hook
  async function bootstrapBackendStatus(): Promise<void> {
    await apiSettings.bootstrap();
  }

  async function refreshBackendStatus(showToast: boolean = false): Promise<boolean> {
    const ok = await apiSettings.refreshBackendStatus(showToast);
    if (showToast && ok) {
      dispatchToast(<Text>后端服务已连接</Text>, { intent: "success" });
    } else if (showToast && !ok) {
      dispatchToast(<Text>后端服务不可用，请检查服务是否启动</Text>, { intent: "error" });
    }
    return ok;
  }

  async function checkApiKeyStatus(): Promise<ApiKeyStatus | null> {
    return await apiSettings.checkApiKeyStatus();
  }

  async function handleSetApiKey(): Promise<void> {
    if (!apiKeyInput.trim() || apiKeyBusy) return;

    const result = await apiSettings.setApiKey(apiKeyInput.trim());

    if (result.success) {
      dispatchToast(<Text>✓ API密钥设置成功</Text>, { intent: "success" });
      setApiKeyDialogOpen(false);
      setApiKeyInput("");
    } else {
      dispatchToast(<Text>✗ {result.message || "API密钥设置失败"}</Text>, { intent: "error" });
    }
  }

  async function handleClearApiKey(): Promise<void> {
    const result = await apiSettings.clearApiKey();
    if (result.success) {
      dispatchToast(<Text>✓ API密钥已清除</Text>, { intent: "success" });
    }
  }

  /**
   * 统一入口 - 所有请求都走 Agent 模式
   * v2.9.9: 删除 shouldUseAgentMode，Agent 自己判断是否需要工具
   * 
   * 原理：
   * - 用户说"你好" → Agent 判断不需要工具，直接回复
   * - 用户说"帮我建表" → Agent 调用工具执行
   * - 和 GitHub Copilot 的工作方式一致
   */
  async function onSend(text: string): Promise<void> {
    const t = text.trim();
    if (!t || busy) return;

    // v4.0: 简化逻辑，所有请求直接发送给新架构的 Agent

    const userMessage: ChatMessage = {
      id: uid(),
      role: "user",
      text: t,
      timestamp: new Date(),
    };
    setMessages((prev) => [...prev, userMessage]);
    setInput("");
    setBusy(true);

    try {
      const backendOk = backendHealthy === true ? true : await refreshBackendStatus(false);
      if (!backendOk) {
        setMessages((prev) => [
          ...prev,
          {
            id: uid(),
            role: "assistant",
            text: "后端不可用，无法处理请求。请确保后端服务已启动后重试。",
            timestamp: new Date(),
          },
        ]);
        setBusy(false);
        return;
      }

      let status = apiKeyStatus;
      if (!status) {
        status = await checkApiKeyStatus();
      }

      if (!status?.configured || !status?.isValid) {
        setMessages((prev) => [
          ...prev,
          {
            id: uid(),
            role: "assistant",
            text: "⚠️ 请先配置有效的 API 密钥以使用 AI 功能。点击右上角的⚙️设置按钮。",
            timestamp: new Date(),
          },
        ]);
        setBusy(false);
        return;
      }

      // ===== 统一 Agent 模式：所有请求都走这里 =====
      console.log("[App] 使用统一 Agent 模式处理请求");
      
      // 显示更自然的思考状态
        const thinkingMsgId = uid();
        setMessages((prev) => [
          ...prev,
          {
            id: thinkingMsgId,
            role: "assistant",
            text: "⏳ 执行中...",  // v2.9.50: 中性状态词，不承诺结果
            timestamp: new Date(),
          },
        ]);

        // 使用新的 Agent 模块执行任务
        setIsAgentRunning(true);
        setAgentThoughts([]);
        
        // 设置实时消息更新的 ref
        currentAgentMsgIdRef.current = thinkingMsgId;
        // v2.9.26: 不再累积步骤，只保留最后一个状态给用户看
        agentStepsRef.current = [];
        
        // v2.9.50: 实时更新只显示中性状态或错误，不显示承诺性文本
        // 用户不需要看到内部思考过程，只需要知道状态
        updateAgentMessageRef.current = (step: string) => {
          // 只更新最后一个状态，不累积
          const displayText = step.includes("失败") || step.includes("错误") 
            ? step  // 错误信息要显示
            : "⏳ 执行中...";  // 其他状态使用中性状态词
          setMessages((prev) =>
            prev.map((msg) =>
              msg.id === thinkingMsgId
                ? { ...msg, text: displayText }
                : msg
            )
          );
        };
        
        // v2.9.8: 设置 useAgent hook 的步骤回调
        agent.setStepCallback(updateAgentMessageRef.current);
        
        // v4.0: 使用新架构的 send() 方法
        try {
          // v4.0: 直接调用 agent.send()，传入上下文
          const result = await agent.send(t, {
            activeSheet: workbookContext?.sheets?.[0]?.name,
            workbookName: workbookContext?.fileName,
          });
          
          // v4.0: 处理执行结果（兼容 ExecutionResult 或 AgentTask）
          let agentResult: { success: boolean; message: string } = { success: false, message: "" };
          if (result && typeof result === "object" && "success" in result) {
            // ExecutionResult
            const r: any = result;
            agentResult = {
              success: Boolean(r.success),
              message: String(r.message || (r.success ? "任务完成" : r.error || "执行失败")),
            };
          } else if (result && typeof result === "object" && "id" in (result as any)) {
            // AgentTask - not completed immediately
            agentResult = { success: false, message: "任务已提交，正在执行中" };
          } else {
            agentResult = { success: false, message: "未知回复" };
          }
        
        // v2.9.26: 重构消息展示 - 只显示 Agent 的核心回复，像人一样说话
        // LLM 通过 respond_to_user 工具返回的 message 才是给用户看的！
        const mainMessage = agentResult.message;
        
        // 判断是否需要添加成功/失败状态
        // 如果 Agent 的回复已经很清晰了，就不需要额外添加状态
        const hasExplicitStatus = 
          mainMessage.includes("✅") || 
          mainMessage.includes("❌") || 
          mainMessage.includes("已完成") ||
          mainMessage.includes("完成") ||
          mainMessage.includes("失败");
        
        // 构建最终消息 - 简洁清晰
        let finalMessage = mainMessage;
        
        // 只有当消息没有明确状态且任务成功时，才添加成功图标
        if (!hasExplicitStatus) {
          if (agentResult.success) {
            // 不添加多余的 "任务完成"，Agent 的回复本身就应该说清楚
          } else {
            finalMessage = `⚠️ ${mainMessage}`;
          }
        }
        
        // 如果消息是默认的 "任务完成"，让它更自然一些
        if (finalMessage === "任务完成" || finalMessage === "") {
          finalMessage = "✅ 已完成！";
        }
        
        // 更新消息 - 只显示 Agent 的核心回复
        setMessages((prev) =>
          prev.map((msg) =>
            msg.id === thinkingMsgId
              ? {
                  ...msg,
                  text: finalMessage,
                }
              : msg
          )
        );
        
        // v2.9.65: 去掉多余的 Toast 通知，消息本身已经足够说明状态
        // 用户觉得每次都弹 Toast 很 low，而且消息里已经有状态了
        
        } catch (agentError) {
          // Agent 执行出错
          const existingSteps = agentStepsRef.current.join("\n");
          setMessages((prev) =>
            prev.map((msg) =>
              msg.id === thinkingMsgId
                ? { ...msg, text: `${existingSteps}\n\n❌ **Agent 执行失败**: ${agentError instanceof Error ? agentError.message : String(agentError)}` }
                : msg
            )
          );
          dispatchToast(<Text>Agent 任务失败</Text>, { intent: "error" });
        } finally {
          setIsAgentRunning(false);
          // 清理 ref
          currentAgentMsgIdRef.current = null;
          updateAgentMessageRef.current = null;
          setBusy(false);
        }
    } catch (e: unknown) {
      const result = await errorHandler.handleError(
        e,
        { operation: "send_message", parameters: { text: t } },
        { showToUser: true, userFriendlyMessage: "������Ϣʱ��������" }
      );

      const messageText = result.userMessage || (e instanceof Error ? e.message : String(e));
      setMessages((prev) => [
        ...prev,
        {
          id: uid(),
          role: "assistant",
          text: `? ${messageText}`,
          timestamp: new Date(),
        },
      ]);
    } finally {
      setBusy(false);
    }
  }

  // v2.9.12: 更新 onSendRef（用于 useSelectionListener）
  React.useEffect(() => {
    onSendRef.current = onSend;
  }, [onSend]);

  // v2.9.9: onApply 重构为通过 Agent 执行操作
  // UI 不再直接调用 Excel API，所有操作通过 Agent 工具层执行
  async function onApply(action: CopilotAction): Promise<void> {
    setBusy(true);
    try {
      const actionLabel = getActionLabel(action);
      
      // 构建自然语言请求让 Agent 执行
      let requestText = "";
      if (action.type === "executeCommand" && action.command) {
        const cmd = action.command;
        const params = JSON.stringify(cmd.parameters || {});
        requestText = `请执行以下 Excel 操作: ${cmd.action || cmd.type}，参数: ${params}`;
      } else if (action.type === "writeRange" || action.type === "setFormula" || action.type === "writeCell") {
        requestText = `请执行操作: ${action.type}，地址: ${action.address || "当前选区"}`;
      } else {
        requestText = `请执行操作: ${actionLabel}`;
      }
      
      // 通过 onSend 调用 Agent 来执行操作
      await onSend(requestText);
      
      addToHistory({
        id: uid(),
        operation: `应用${action.type}`,
        timestamp: new Date(),
        success: true,
        details: actionLabel,
      });
    } catch (e: unknown) {
      const result = await errorHandler.handleError(
        e,
        { operation: "apply_action" },
        { showToUser: true, userFriendlyMessage: "应用操作失败" }
      );

      const messageText = result.userMessage || (e instanceof Error ? e.message : String(e));
      setMessages((prev) => [
        ...prev,
        {
          id: uid(),
          role: "assistant",
          text: `❌ ${messageText}`,
          timestamp: new Date(),
        },
      ]);
    } finally {
      setBusy(false);
    }
  }

  function addToHistory(item: OperationHistoryItem): void {
    setHistory((prev) => [item, ...prev].slice(0, 50));
    setHistoryIndex(-1);
  }

  // �ж��Ƿ���ʾ��ӭ���棨����Ϣʱ��
  const showWelcome = messages.length <= 1;
  const isDisabled = busy || backendHealthy === false;

  // v2.9.8: 构建 HeaderBar 需要的 workbookSummary
  const workbookSummary = workbookContext ? {
    sheetCount: workbookContext.sheets.length,
    tableCount: workbookContext.tables.length,
    formulaCount: workbookContext.totalFormulas,
    qualityScore: workbookContext.overallQualityScore,
  } : undefined;

  // v2.9.8: 构建 InsightPanel 需要的数据
  const insightDataSummary: InsightDataSummary | undefined = dataSummary ? {
    rowCount: dataSummary.rowCount,
    columnCount: dataSummary.columnCount,
    numericColumns: dataSummary.numericColumns,
    qualityScore: dataSummary.qualityScore,
  } : undefined;

  const insightSuggestions: InsightSuggestion[] = proactiveSuggestions.map(s => ({
    id: s.id,
    title: s.title,
    description: s.description,
    icon: s.icon as InsightSuggestion["icon"],
    action: () => void s.action(),
  }));

  return (
    <FluentProvider theme={isDarkTheme ? webDarkTheme : webLightTheme} className={styles.app}>
      <Toaster toasterId={toasterId} />
      
      <div className={styles.container} data-testid="copilot-container">
        {/* ===== 顶部状态栏 ===== */}
        <HeaderBar
          backendHealthy={backendHealthy}
          workbookSummary={workbookSummary}
          isScanning={isScanning}
          scanProgress={scanProgress}
          selectionAddress={lastSelection?.address}
          undoCount={undoStack.length}
          apiKeyValid={apiKeyStatus?.isValid ?? false}
          onRefreshWorkbook={handleRefreshWorkbook}
          onUndo={handleUndo}
          onOpenSettings={handleOpenSettings}
        />

        {/* ===== 进度条 ===== */}
        {(analysisProgress > 0 || isScanning) && (
          <div className={styles.progressWrapper}>
            <ProgressBar value={isScanning ? scanProgress / 100 : analysisProgress / 100} />
          </div>
        )}

        {/* v4.3: 主动洞察 Agent 分析进度 */}
        {proactiveAgent.isAnalyzing && (
          <div className={styles.agentProgressWrapper}>
            <ProgressBar className={styles.agentProgressBar} />
            <Caption1 className={styles.agentProgressText}>
              🔍 正在分析工作表...
            </Caption1>
          </div>
        )}

        {/* v2.9.17: Agent 执行进度显示 */}
        {isAgentRunning && agent.state.progress && (
          <div className={styles.agentProgressWrapper}>
            <ProgressBar 
              value={agent.state.progress.percentage / 100} 
              className={styles.agentProgressBar}
            />
            <Caption1 className={styles.agentProgressText}>
              {agent.state.progress.currentPhase}
            </Caption1>
          </div>
        )}

        {/* ===== 数据洞察面板 ===== */}
        {(dataSummary || isAnalyzing) && (
          <InsightPanel
            isAnalyzing={isAnalyzing}
            dataSummary={insightDataSummary}
            selectionAddress={lastSelection?.address}
            suggestions={insightSuggestions}
          />
        )}

        {/* ===== v4.3: 主动洞察快速操作 ===== */}
        {proactiveAgent.quickActions.length > 0 && (
          <div className={styles.proactiveActionsWrapper}>
            <Caption1 className={styles.proactiveActionsTitle}>
              💡 根据分析，你可能想要：
            </Caption1>
            {proactiveAgent.quickActions.slice(0, 4).map((action, index) => (
              <button
                key={index}
                onClick={async () => {
                  await proactiveAgent.sendMessage(action.action);
                }}
                disabled={busy || proactiveAgent.isExecuting}
                className={styles.proactiveActionButton}
              >
                {action.label}
              </button>
            ))}
            {proactiveAgent.suggestions.length > 0 && (
              <button
                onClick={async () => {
                  await proactiveAgent.executeAll();
                }}
                disabled={busy || proactiveAgent.isExecuting}
                className={styles.proactiveExecuteAllButton}
              >
                ✓ 全部执行
              </button>
            )}
          </div>
        )}

        {/* ===== 聊天区域 ===== */}
        <div className={styles.chatContainer}>
          {showWelcome ? (
            <WelcomeView disabled={isDisabled} onSend={onSend} />
          ) : (
            <MessageList
              messages={messages}
              busy={busy}
              isAgentRunning={isAgentRunning}
              onApply={onApply}
            />
          )}
        </div>

        {/* ===== 输入区域 ===== */}
        <ChatInputArea
          value={input}
          onChange={setInput}
          onSend={onSend}
          busy={busy}
          backendHealthy={backendHealthy}
        />
      </div>

      {/* ===== API 配置对话框 ===== */}
      <ApiConfigDialog
        open={apiKeyDialogOpen}
        onOpenChange={setApiKeyDialogOpen}
        apiKeyInput={apiKeyInput}
        onApiKeyInputChange={setApiKeyInput}
        apiKeyStatus={apiKeyStatus}
        backendHealthy={backendHealthy}
        backendChecking={backendChecking}
        apiKeyBusy={apiKeyBusy}
        onRefreshBackend={() => void refreshBackendStatus(true)}
        onSaveApiKey={handleSetApiKey}
        onClearApiKey={handleClearApiKey}
      />

      {/* ===== 预览确认对话框 ===== */}
      <PreviewConfirmDialog
        open={previewDialogOpen}
        onOpenChange={setPreviewDialogOpen}
        previewMessage={previewMessage}
        onConfirm={() => {
          const handler = (window as unknown as Record<string, unknown>)._confirmHandler as ((v: boolean) => void) | undefined;
          if (handler) {
            handler(true);
            delete (window as unknown as Record<string, unknown>)._confirmHandler;
          }
        }}
        onCancel={() => {
          const handler = (window as unknown as Record<string, unknown>)._confirmHandler as ((v: boolean) => void) | undefined;
          if (handler) {
            handler(false);
            delete (window as unknown as Record<string, unknown>)._confirmHandler;
          }
        }}
      />

      {/* ===== v3.0: Agent 层审批确认对话框 ===== */}
      <ApprovalDialog
        open={agent.state.status === "awaiting_approval" && agent.state.pendingApproval !== null}
        approvalRequest={
          agent.state.pendingApproval && typeof agent.state.pendingApproval === "object"
            ? (agent.state.pendingApproval as any)
            : null
        }
        onConfirm={(approvalId) => {
          agent.approve(approvalId);
        }}
        onCancel={(approvalId) => {
          agent.reject(approvalId, "用户取消");
        }}
        onClose={() => {
          // 关闭弹窗但不做决定时，默认拒绝
          if (agent.state.pendingApproval && (agent.state.pendingApproval as any).approvalId) {
            agent.reject((agent.state.pendingApproval as any).approvalId, "用户关闭弹窗");
          }
        }}
      />
    </FluentProvider>
  );
};

export default App;
