/**
 * 消息内容解析器
 *
 * 纯函数模块，用于将 Agent 返回的消息文本解析为结构化的 React 节点
 * 支持：粗体文本、emoji前缀、列表、执行步骤块
 *
 * @module messageParser
 */

import * as React from "react";

/**
 * 消息样式配置（可选）
 */
export interface MessageStyles {
  messageNumberedList?: string;
  messageBulletList?: string;
  formattedList?: string;
  messageListItem?: string;
  messageStepsBlock?: string;
  messageStepItem?: string;
  messageStepSuccess?: string;
  messageStepError?: string;
  messageSpacer?: string;
  messageBold?: string;
}

/**
 * 格式化行内文本（粗体等）
 *
 * @param text - 原始文本
 * @param styles - 样式映射
 * @returns React 节点
 */
export function formatInlineText(
  text: string,
  styles?: MessageStyles
): React.ReactNode {
  const parts: React.ReactNode[] = [];
  let lastIndex = 0;
  const boldRegex = /\*\*(.+?)\*\*/g;
  let match;

  while ((match = boldRegex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push(text.slice(lastIndex, match.index));
    }
    parts.push(
      <strong key={match.index} className={styles?.messageBold || ""}>
        {match[1]}
      </strong>
    );
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < text.length) {
    parts.push(text.slice(lastIndex));
  }

  return parts.length > 0 ? <>{parts}</> : text;
}

/**
 * 解析消息文本，提取结构化内容
 *
 * @param text - 原始消息文本
 * @param styles - 样式映射
 * @returns React 节点数组
 */
export function parseMessageContent(
  text: string,
  styles?: MessageStyles
): React.ReactNode[] {
  const lines = text.split("\n");
  const elements: React.ReactNode[] = [];
  let currentList: string[] = [];
  let listType: "bullet" | "numbered" | null = null;
  let isInStepsBlock = false;
  let stepsContent: string[] = [];

  const flushList = () => {
    if (currentList.length > 0) {
      const listClass =
        listType === "numbered"
          ? styles?.messageNumberedList || ""
          : styles?.messageBulletList || "";
      elements.push(
        <ul
          key={`list-${elements.length}`}
          className={`${styles?.formattedList || ""} ${listClass}`}
        >
          {currentList.map((item, i) => (
            <li key={i} className={styles?.messageListItem || ""}>
              {formatInlineText(item, styles)}
            </li>
          ))}
        </ul>
      );
      currentList = [];
      listType = null;
    }
  };

  const flushSteps = () => {
    if (stepsContent.length > 0) {
      elements.push(
        <div
          key={`steps-${elements.length}`}
          className={styles?.messageStepsBlock || ""}
        >
          {stepsContent.map((step, i) => {
            const isSuccess = step.includes("✅");
            const isError = step.includes("❌");
            const stepClass = isError
              ? styles?.messageStepError || ""
              : isSuccess
                ? styles?.messageStepSuccess || ""
                : "";
            return (
              <div
                key={i}
                className={`${styles?.messageStepItem || ""} ${stepClass}`}
              >
                {formatInlineText(step, styles)}
              </div>
            );
          })}
        </div>
      );
      stepsContent = [];
      isInStepsBlock = false;
    }
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmedLine = line.trim();

    // 检测执行步骤块
    if (
      trimmedLine.includes("**执行步骤:**") ||
      trimmedLine.includes("**已完成操作：**") ||
      trimmedLine.includes("📋")
    ) {
      flushList();
      isInStepsBlock = true;
      continue;
    }

    // 在步骤块中
    if (isInStepsBlock) {
      if (trimmedLine.match(/^\d+\.\s/) || trimmedLine.startsWith("•")) {
        stepsContent.push(trimmedLine);
        continue;
      } else if (trimmedLine === "") {
        flushSteps();
        continue;
      } else if (stepsContent.length > 0) {
        flushSteps();
      }
    }

    // 检测列表项
    if (trimmedLine.startsWith("• ") || trimmedLine.startsWith("- ")) {
      if (listType !== "bullet") {
        flushList();
        listType = "bullet";
      }
      currentList.push(trimmedLine.replace(/^[•-]\s/, ""));
      continue;
    }

    if (trimmedLine.match(/^\d+\.\s/) && !isInStepsBlock) {
      if (listType !== "numbered") {
        flushList();
        listType = "numbered";
      }
      currentList.push(trimmedLine.replace(/^\d+\.\s/, ""));
      continue;
    }

    // 普通文本行
    flushList();
    flushSteps();

    if (trimmedLine === "") {
      if (elements.length > 0) {
        elements.push(
          <div key={`spacer-${i}`} className={styles?.messageSpacer || ""} />
        );
      }
    } else {
      elements.push(
        <div key={`line-${i}`} className={styles?.messageListItem || ""}>
          {formatInlineText(trimmedLine, styles)}
        </div>
      );
    }
  }

  flushList();
  flushSteps();

  return elements;
}
