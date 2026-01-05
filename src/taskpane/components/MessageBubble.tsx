/**
 * 消息气泡组件
 *
 * 职责单一：只负责渲染单条消息
 * - 不知道 Agent 内部结构
 * - 只接收 message model 和回调
 *
 * @module MessageBubble
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

// ========== Props ==========

export interface MessageBubbleProps {
  message: ChatMessage;
  onExecuteAction?: (action: CopilotAction) => Promise<void>;
  isBusy?: boolean;
}

// ========== Styles ==========

const useStyles = makeStyles({
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
    maxWidth: "92%",
    ...shorthands.padding("14px", "18px"),
    ...shorthands.borderRadius("18px"),
    fontSize: "13.5px",
    lineHeight: "1.65",
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
    WebkitFontSmoothing: "antialiased",
    MozOsxFontSmoothing: "grayscale",
    transition: "box-shadow 0.2s ease",
    ":hover": {
      boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
    },
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

  messageTime: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground4,
    ...shorthands.padding("0", "4px"),
  },

  messageContentWrapper: {
    width: "100%",
  },

  actionButtons: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    marginTop: "8px",
  },

  // 消息格式化样式
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
  messageStepsBlock: {
    backgroundColor: "rgba(0, 120, 212, 0.06)",
    ...shorthands.borderRadius("8px"),
    ...shorthands.padding("12px"),
    ...shorthands.margin("8px", "0"),
    fontSize: "12px",
  },
  messageStepItem: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    ...shorthands.padding("4px", "0"),
  },
  messageStepSuccess: {
    color: "#107c10",
  },
  messageStepError: {
    color: "#d13438",
  },
  messageSpacer: {
    height: "8px",
  },
  messageBold: {
    fontWeight: 600,
  },
});

// ========== Component ==========

export const MessageBubble: React.FC<MessageBubbleProps> = ({
  message,
  onExecuteAction,
  isBusy = false,
}) => {
  const styles = useStyles();
  const isUser = message.role === "user";

  // 将 styles 对象转换为 MessageStyles 接口
  const messageStyles: MessageStyles = {
    messageNumberedList: styles.messageNumberedList,
    messageBulletList: styles.messageBulletList,
    formattedList: styles.formattedList,
    messageListItem: styles.messageListItem,
    messageStepsBlock: styles.messageStepsBlock,
    messageStepItem: styles.messageStepItem,
    messageStepSuccess: styles.messageStepSuccess,
    messageStepError: styles.messageStepError,
    messageSpacer: styles.messageSpacer,
    messageBold: styles.messageBold,
  };

  const getActionLabel = (action: CopilotAction): string => {
    switch (action.type) {
      case "writeRange":
        return `写入 ${action.address}`;
      case "setFormula":
        return "应用公式";
      case "writeCell":
        return "写入单元格";
      case "executeCommand":
        return action.label;
      default:
        return "应用";
    }
  };

  return (
    <div
      className={`${styles.messageWrapper} ${
        isUser ? styles.messageWrapperUser : styles.messageWrapperAssistant
      }`}
    >
      <div
        className={`${styles.messageBubble} ${
          isUser ? styles.messageBubbleUser : styles.messageBubbleAssistant
        }`}
      >
        {isUser ? (
          <span>{message.text}</span>
        ) : (
          <div className={styles.messageContentWrapper}>
            {parseMessageContent(message.text, messageStyles)}
          </div>
        )}

        {/* 操作按钮 */}
        {!isUser && message.actions && message.actions.length > 0 && (
          <div className={styles.actionButtons}>
            {message.actions.map((action, idx) => (
              <Button
                key={idx}
                appearance="primary"
                size="small"
                icon={
                  action.type === "executeCommand" ? (
                    <CheckmarkCircleRegular />
                  ) : undefined
                }
                onClick={() => onExecuteAction?.(action)}
                disabled={isBusy}
              >
                {getActionLabel(action)}
              </Button>
            ))}
          </div>
        )}
      </div>
      <span className={styles.messageTime}>
        {message.timestamp.toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
        })}
      </span>
    </div>
  );
};

export default MessageBubble;
