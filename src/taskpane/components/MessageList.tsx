/**
 * MessageList - 消息列表组件
 * @file src/taskpane/components/MessageList.tsx
 * @description v2.9.8 从 App.tsx 提取，包含消息渲染、加载状态
 */
import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Button,
} from "@fluentui/react-components";
import { CheckmarkCircleRegular } from "@fluentui/react-icons";
import type { ChatMessage, CopilotAction } from "../types/ui.types";
import { parseMessageContent, MessageStyles } from "../utils/messageParser";

const useStyles = makeStyles({
  messageList: {
    flex: 1,
    overflowY: "auto",
    ...shorthands.padding("16px"),
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  messageWrapper: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    animationName: {
      from: { opacity: 0, transform: "translateY(8px)" },
      to: { opacity: 1, transform: "translateY(0)" },
    },
    animationDuration: "0.25s",
    animationTimingFunction: "cubic-bezier(0.16, 1, 0.3, 1)",
    animationFillMode: "forwards",
    maxWidth: "100%",
  },
  messageWrapperUser: {
    alignItems: "flex-end",
  },
  messageWrapperAssistant: {
    alignItems: "flex-start",
  },
  messageBubble: {
    maxWidth: "85%",
    ...shorthands.padding("10px", "14px"),
    ...shorthands.borderRadius("12px"),
    wordBreak: "break-word",
    lineHeight: "1.5",
    fontSize: "13px",
  },
  messageBubbleUser: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderBottomRightRadius: "4px",
  },
  messageBubbleAssistant: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: "4px",
  },
  messageContentWrapper: {
    width: "100%",
  },
  messageTime: {
    fontSize: "10px",
    color: tokens.colorNeutralForeground4,
  },
  actionButtons: {
    marginTop: "8px",
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
  },
  loadingBubble: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    ...shorthands.padding("12px", "16px"),
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.borderRadius("12px"),
  },
  typingDots: {
    display: "flex",
    gap: "4px",
  },
  typingDot: {
    width: "6px",
    height: "6px",
    ...shorthands.borderRadius("50%"),
    backgroundColor: tokens.colorNeutralForeground3,
    animationName: {
      "0%, 60%, 100%": { opacity: 0.3 },
      "30%": { opacity: 1 },
    },
    animationDuration: "1.2s",
    animationIterationCount: "infinite",
  },
  typingDot2: {
    animationDelay: "0.2s",
  },
  typingDot3: {
    animationDelay: "0.4s",
  },
  loadingText: {
    fontSize: "13px",
    color: tokens.colorNeutralForeground3,
  },
  // 消息内容格式化样式
  messageListItem: {
    marginBottom: "4px",
    lineHeight: "1.6",
  },
  formattedList: {
    margin: "8px 0",
    paddingLeft: "20px",
  },
  messageBulletList: {
    listStyleType: "disc",
  },
  messageNumberedList: {
    listStyleType: "decimal",
  },
  codeBlock: {
    backgroundColor: tokens.colorNeutralBackground4,
    ...shorthands.padding("8px", "12px"),
    ...shorthands.borderRadius("6px"),
    fontFamily: "Consolas, monospace",
    fontSize: "12px",
    whiteSpace: "pre-wrap",
    marginTop: "8px",
    marginBottom: "8px",
  },
  inlineCode: {
    backgroundColor: tokens.colorNeutralBackground4,
    ...shorthands.padding("2px", "6px"),
    ...shorthands.borderRadius("4px"),
    fontFamily: "Consolas, monospace",
    fontSize: "12px",
  },
});

export interface MessageListProps {
  /** 消息列表 */
  messages: ChatMessage[];
  /** 是否正在加载 */
  busy: boolean;
  /** 是否 Agent 运行中 */
  isAgentRunning: boolean;
  /** 应用操作回调 */
  onApply: (action: CopilotAction) => Promise<void>;
}

/**
 * 获取操作按钮标签
 */
function getActionLabel(action: CopilotAction): string {
  switch (action.type) {
    case "writeRange":
      return `写入 ${action.address}`;
    case "setFormula":
      return "应用公式";
    case "writeCell":
      return "写入单元格";
    case "executeCommand":
      return action.label || "执行";
    default:
      return "应用";
  }
}

/**
 * 消息列表组件
 * 渲染聊天消息和加载状态
 */
export const MessageList: React.FC<MessageListProps> = ({
  messages,
  busy,
  isAgentRunning,
  onApply,
}) => {
  const styles = useStyles();
  const messageStyles: MessageStyles = styles;
  
  // v2.9.25: 添加自动滚动到最新消息
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  const containerRef = React.useRef<HTMLDivElement>(null);
  
  // 滚动到底部
  const scrollToBottom = React.useCallback(() => {
    if (messagesEndRef.current) {
      messagesEndRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, []);
  
  // 当消息列表变化或 busy 状态变化时，滚动到底部
  React.useEffect(() => {
    // 使用 setTimeout 确保 DOM 已更新
    const timer = setTimeout(scrollToBottom, 100);
    return () => clearTimeout(timer);
  }, [messages, busy, scrollToBottom]);

  return (
    <div className={styles.messageList} ref={containerRef}>
      {messages.map((msg) => (
        <div
          key={msg.id}
          className={`${styles.messageWrapper} ${
            msg.role === "user" ? styles.messageWrapperUser : styles.messageWrapperAssistant
          }`}
        >
          <div
            className={`${styles.messageBubble} ${
              msg.role === "user" ? styles.messageBubbleUser : styles.messageBubbleAssistant
            }`}
          >
            {msg.role === "user" ? (
              <span>{msg.text}</span>
            ) : (
              <div className={styles.messageContentWrapper}>
                {parseMessageContent(msg.text, messageStyles)}
              </div>
            )}

            {/* 操作按钮 */}
            {msg.actions && msg.actions.length > 0 && (
              <div className={styles.actionButtons}>
                {msg.actions.map((action, idx) => (
                  <Button
                    key={idx}
                    appearance="primary"
                    size="small"
                    icon={action.type === "executeCommand" ? <CheckmarkCircleRegular /> : undefined}
                    onClick={() => void onApply(action)}
                    disabled={busy}
                  >
                    {getActionLabel(action)}
                  </Button>
                ))}
              </div>
            )}
          </div>
          <span className={styles.messageTime}>
            {msg.timestamp.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
          </span>
        </div>
      ))}

      {/* 加载状态 - 打字动画 */}
      {busy && (
        <div className={`${styles.messageWrapper} ${styles.messageWrapperAssistant}`}>
          <div className={styles.loadingBubble}>
            <div className={styles.typingDots}>
              <span className={styles.typingDot} />
              <span className={`${styles.typingDot} ${styles.typingDot2}`} />
              <span className={`${styles.typingDot} ${styles.typingDot3}`} />
            </div>
            <span className={styles.loadingText}>
              {isAgentRunning ? "正在执行..." : "思考中..."}
            </span>
          </div>
        </div>
      )}
      
      {/* v2.9.25: 滚动锚点 - 用于自动滚动到最新消息 */}
      <div ref={messagesEndRef} style={{ height: 1 }} />
    </div>
  );
};

export default MessageList;
