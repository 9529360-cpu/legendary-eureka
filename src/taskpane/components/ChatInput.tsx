import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Button,
  Textarea,
  Tooltip,
  Spinner,
} from "@fluentui/react-components";
import {
  SendRegular,
  MicRegular,
  AttachRegular,
  SparkleRegular,
} from "@fluentui/react-icons";

interface ChatInputProps {
  value: string;
  onChange: (value: string) => void;
  onSend: () => void;
  isLoading?: boolean;
  disabled?: boolean;
  placeholder?: string;
  suggestions?: string[];
  onSuggestionClick?: (suggestion: string) => void;
}

const useStyles = makeStyles({
  container: {
    ...shorthands.padding("16px"),
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  inputWrapper: {
    display: "flex",
    alignItems: "flex-end",
    gap: "8px",
  },
  textareaContainer: {
    flex: 1,
    position: "relative",
  },
  textarea: {
    width: "100%",
    minHeight: "44px",
    maxHeight: "120px",
    resize: "none",
    "& textarea": {
      paddingRight: "80px",
    },
  },
  actions: {
    display: "flex",
    gap: "4px",
    position: "absolute",
    right: "8px",
    bottom: "8px",
  },
  suggestions: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
    marginBottom: "12px",
  },
  suggestionChip: {
    ...shorthands.padding("6px", "12px"),
    ...shorthands.borderRadius("16px"),
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: "pointer",
    fontSize: "13px",
    transition: "all 0.15s ease",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    alignItems: "center",
    gap: "4px",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground4,
      borderColor: tokens.colorBrandStroke1 as unknown as undefined,
    },
  },
  hint: {
    marginTop: "8px",
    fontSize: "12px",
    color: tokens.colorNeutralForeground4,
    textAlign: "center" as const,
  },
});

const ChatInput: React.FC<ChatInputProps> = ({
  value,
  onChange,
  onSend,
  isLoading = false,
  disabled = false,
  placeholder = "输入消息...",
  suggestions = [],
  onSuggestionClick,
}) => {
  const styles = useStyles();
  const textareaRef = React.useRef<HTMLTextAreaElement>(null);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (value.trim() && !isLoading && !disabled) {
        onSend();
      }
    }
  };

  const handleSuggestionClick = (suggestion: string) => {
    if (onSuggestionClick) {
      onSuggestionClick(suggestion);
    } else {
      onChange(suggestion);
      // 自动聚焦到输入框
      textareaRef.current?.focus();
    }
  };

  // 自动调整高度
  React.useEffect(() => {
    const textarea = textareaRef.current;
    if (textarea) {
      textarea.style.height = "auto";
      textarea.style.height = `${Math.min(textarea.scrollHeight, 120)}px`;
    }
  }, [value]);

  return (
    <div className={styles.container}>
      {/* 快捷建议 */}
      {suggestions.length > 0 && (
        <div className={styles.suggestions}>
          {suggestions.map((suggestion, index) => (
            <div
              key={index}
              className={styles.suggestionChip}
              onClick={() => handleSuggestionClick(suggestion)}
            >
              <SparkleRegular style={{ fontSize: 14 }} />
              {suggestion}
            </div>
          ))}
        </div>
      )}

      <div className={styles.inputWrapper}>
        <div className={styles.textareaContainer}>
          <Textarea
            ref={textareaRef}
            className={styles.textarea}
            placeholder={placeholder}
            value={value}
            onChange={(e, data) => onChange(data.value)}
            onKeyDown={handleKeyDown}
            disabled={disabled || isLoading}
            resize="none"
          />
          <div className={styles.actions}>
            <Tooltip content="附件 (即将推出)" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<AttachRegular />}
                disabled
              />
            </Tooltip>
            <Tooltip content="语音输入 (即将推出)" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<MicRegular />}
                disabled
              />
            </Tooltip>
          </div>
        </div>

        <Tooltip content="发送消息 (Enter)" relationship="label">
          <Button
            appearance="primary"
            size="medium"
            icon={isLoading ? <Spinner size="tiny" /> : <SendRegular />}
            onClick={onSend}
            disabled={!value.trim() || isLoading || disabled}
          />
        </Tooltip>
      </div>

      <div className={styles.hint}>
        按 Enter 发送，Shift + Enter 换行
      </div>
    </div>
  );
};

export default ChatInput;
