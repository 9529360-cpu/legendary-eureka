import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Body1,
  Caption1,
  Button,
  Tooltip,
} from "@fluentui/react-components";
import { CopyRegular, PlayRegular, CheckmarkCircleRegular, ErrorCircleRegular } from "@fluentui/react-icons";

type Role = "user" | "assistant";

type CellValue = string | number | boolean | null;

type CopilotAction =
  | { type: "writeRange"; address: string; values: CellValue[][] }
  | { type: "setFormula"; address: string; formula: string }
  | { type: "writeCell"; address: string; value: CellValue }
  | { type: "executeCommand"; command: any; label: string };

export interface ChatMessageData {
  id: string;
  role: Role;
  text: string;
  actions?: CopilotAction[];
  timestamp: Date;
}

interface ChatMessageProps {
  message: ChatMessageData;
  onExecuteAction?: (action: CopilotAction) => Promise<void>;
  onCopyMessage?: (text: string) => void;
}

const useStyles = makeStyles({
  messageContainer: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.padding("8px", "16px"),
    animationName: {
      from: { opacity: 0, transform: "translateY(10px)" },
      to: { opacity: 1, transform: "translateY(0)" },
    },
    animationDuration: "0.2s",
    animationTimingFunction: "ease-out",
  },
  userMessage: {
    alignItems: "flex-end",
  },
  assistantMessage: {
    alignItems: "flex-start",
  },
  messageBubble: {
    maxWidth: "85%",
    ...shorthands.padding("12px", "16px"),
    ...shorthands.borderRadius("12px"),
    wordBreak: "break-word",
    whiteSpace: "pre-wrap",
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderBottomRightRadius: "4px",
  },
  assistantBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: "4px",
  },
  messageActions: {
    display: "flex",
    gap: "4px",
    marginTop: "8px",
  },
  timestamp: {
    marginTop: "4px",
    color: tokens.colorNeutralForeground4,
  },
  actionButtons: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    marginTop: "12px",
  },
  actionButton: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground4,
    ...shorthands.padding("8px", "12px"),
    ...shorthands.borderRadius("6px"),
    fontFamily: "Consolas, Monaco, monospace",
    fontSize: "13px",
    overflowX: "auto",
    marginTop: "8px",
  },
});

const ChatMessage: React.FC<ChatMessageProps> = ({
  message,
  onExecuteAction,
  onCopyMessage,
}) => {
  const styles = useStyles();
  const [executingAction, setExecutingAction] = React.useState<string | null>(null);
  const [executedActions, setExecutedActions] = React.useState<Set<string>>(new Set());
  const [failedActions, setFailedActions] = React.useState<Set<string>>(new Set());

  const isUser = message.role === "user";

  const handleExecuteAction = async (action: CopilotAction, index: number) => {
    if (!onExecuteAction) return;
    
    const actionKey = `${message.id}-${index}`;
    setExecutingAction(actionKey);
    
    try {
      await onExecuteAction(action);
      setExecutedActions((prev) => new Set(prev).add(actionKey));
      setFailedActions((prev) => {
        const newSet = new Set(prev);
        newSet.delete(actionKey);
        return newSet;
      });
    } catch (_error) {
      setFailedActions((prev) => new Set(prev).add(actionKey));
    } finally {
      setExecutingAction(null);
    }
  };

  const handleCopy = () => {
    if (onCopyMessage) {
      onCopyMessage(message.text);
    }
  };

  const getActionLabel = (action: CopilotAction): string => {
    switch (action.type) {
      case "writeRange":
        return `写入 ${action.address}`;
      case "writeCell":
        return `设置 ${action.address} = ${action.value}`;
      case "setFormula":
        return `设置公式 ${action.address}`;
      case "executeCommand":
        return action.label || "执行操作";
      default:
        return "执行";
    }
  };

  const renderContent = (text: string) => {
    // 简单的代码块渲染
    const parts = text.split(/(```[\s\S]*?```)/g);
    return parts.map((part, index) => {
      if (part.startsWith("```") && part.endsWith("```")) {
        const code = part.slice(3, -3).replace(/^\w*\n/, "");
        return (
          <div key={index} className={styles.codeBlock}>
            {code}
          </div>
        );
      }
      return <span key={index}>{part}</span>;
    });
  };

  return (
    <div
      className={`${styles.messageContainer} ${
        isUser ? styles.userMessage : styles.assistantMessage
      }`}
    >
      <div
        className={`${styles.messageBubble} ${
          isUser ? styles.userBubble : styles.assistantBubble
        }`}
      >
        <Body1>{renderContent(message.text)}</Body1>

        {/* AI 消息的操作按钮 */}
        {!isUser && message.actions && message.actions.length > 0 && (
          <div className={styles.actionButtons}>
            {message.actions.map((action, index) => {
              const actionKey = `${message.id}-${index}`;
              const isExecuting = executingAction === actionKey;
              const isExecuted = executedActions.has(actionKey);
              const isFailed = failedActions.has(actionKey);

              return (
                <Button
                  key={index}
                  appearance={isExecuted ? "secondary" : "primary"}
                  size="small"
                  disabled={isExecuting}
                  onClick={() => handleExecuteAction(action, index)}
                  icon={
                    isExecuted ? (
                      <CheckmarkCircleRegular />
                    ) : isFailed ? (
                      <ErrorCircleRegular />
                    ) : (
                      <PlayRegular />
                    )
                  }
                >
                  {isExecuting ? "执行中..." : getActionLabel(action)}
                </Button>
              );
            })}
          </div>
        )}
      </div>

      {/* 工具栏 */}
      <div className={styles.messageActions}>
        {!isUser && (
          <Tooltip content="复制消息" relationship="label">
            <Button
              appearance="subtle"
              size="small"
              icon={<CopyRegular />}
              onClick={handleCopy}
            />
          </Tooltip>
        )}
        <Caption1 className={styles.timestamp}>
          {message.timestamp.toLocaleTimeString("zh-CN", {
            hour: "2-digit",
            minute: "2-digit",
          })}
        </Caption1>
      </div>
    </div>
  );
};

export default ChatMessage;
