/**
 * ChatInputArea - 聊天输入区域组件
 * @file src/taskpane/components/ChatInputArea.tsx
 * @description v2.9.8 从 App.tsx 提取，包含输入框、发送按钮和连接状态提示
 */
import * as React from "react";
import { Textarea, Button, makeStyles, tokens } from "@fluentui/react-components";
import { SendRegular, DismissCircleRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  composerWrapper: {
    padding: "12px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  composerInner: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-end",
  },
  textarea: {
    flex: 1,
    "& textarea": {
      minHeight: "40px",
      maxHeight: "120px",
    },
  },
  sendButton: {
    minWidth: "36px",
    height: "36px",
    padding: "8px",
  },
  composerHint: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    marginTop: "6px",
    fontSize: "12px",
  },
  textError: {
    color: tokens.colorPaletteRedForeground1,
  },
});

export interface ChatInputAreaProps {
  /** 当前输入内容 */
  value: string;
  /** 输入内容变更回调 */
  onChange: (value: string) => void;
  /** 发送消息回调 */
  onSend: (text: string) => void;
  /** 是否正在处理中 */
  busy: boolean;
  /** 后端连接状态 */
  backendHealthy: boolean | null;
  /** 占位符文本 */
  placeholder?: string;
}

/**
 * 聊天输入区域组件
 * 包含输入框、发送按钮和连接状态提示
 */
export const ChatInputArea: React.FC<ChatInputAreaProps> = ({
  value,
  onChange,
  onSend,
  busy,
  backendHealthy,
  placeholder = "输入自然语言指令",
}) => {
  const styles = useStyles();

  const handleKeyDown = React.useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        onSend(value);
      }
    },
    [onSend, value]
  );

  const handleSendClick = React.useCallback(() => {
    onSend(value);
  }, [onSend, value]);

  return (
    <div className={styles.composerWrapper}>
      <div className={styles.composerInner}>
        <Textarea
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          resize="none"
          className={styles.textarea}
          disabled={busy}
        />
        <Button
          appearance="primary"
          icon={<SendRegular />}
          className={styles.sendButton}
          onClick={handleSendClick}
          disabled={busy || !value.trim() || backendHealthy === false}
        />
      </div>
      {backendHealthy === false && (
        <div className={`${styles.composerHint} ${styles.textError}`}>
          <DismissCircleRegular /> 后端服务未连接
        </div>
      )}
    </div>
  );
};

export default ChatInputArea;
