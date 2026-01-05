/**
 * WelcomeView - 欢迎界面组件
 * @file src/taskpane/components/WelcomeView.tsx
 * @description v2.9.8 从 App.tsx 提取，包含欢迎消息和快捷操作
 */
import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
} from "@fluentui/react-components";
import {
  SparkleRegular,
  TableSimpleRegular,
  MathFormulaRegular,
  ChartMultipleRegular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  welcomeState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    ...shorthands.padding("32px"),
    textAlign: "center",
  },
  welcomeIcon: {
    width: "64px",
    height: "64px",
    ...shorthands.borderRadius("50%"),
    backgroundColor: tokens.colorBrandBackground,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: "16px",
    "& svg": {
      fontSize: "28px",
      color: tokens.colorNeutralForegroundOnBrand,
    },
  },
  welcomeTitle: {
    fontSize: "20px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    marginBottom: "8px",
  },
  welcomeSubtitle: {
    fontSize: "14px",
    color: tokens.colorNeutralForeground3,
    marginBottom: "24px",
  },
  quickActionsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "12px",
    width: "100%",
    maxWidth: "320px",
  },
  actionCard: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
    ...shorthands.padding("16px"),
    ...shorthands.borderRadius("12px"),
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    cursor: "pointer",
    transitionProperty: "background-color",
    transitionDuration: "0.15s",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3Hover,
    },
  },
  actionCardDisabled: {
    opacity: 0.5,
    cursor: "not-allowed",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground3,
    },
  },
  actionIcon: {
    fontSize: "24px",
    color: tokens.colorBrandForeground1,
  },
  actionLabel: {
    fontSize: "13px",
    fontWeight: 500,
    color: tokens.colorNeutralForeground1,
  },
});

export interface QuickAction {
  id: string;
  label: string;
  icon: "summarize" | "formula" | "chart" | "analyze";
  prompt: string;
}

const defaultActions: QuickAction[] = [
  { id: "summarize", label: "总结数据", icon: "summarize", prompt: "总结当前选中的数据" },
  { id: "formula", label: "生成公式", icon: "formula", prompt: "为当前数据生成公式" },
  { id: "chart", label: "创建图表", icon: "chart", prompt: "为数据创建合适的图表" },
  { id: "analyze", label: "分析趋势", icon: "analyze", prompt: "分析数据趋势和异常" },
];

export interface WelcomeViewProps {
  /** 是否禁用操作 */
  disabled: boolean;
  /** 发送消息回调 */
  onSend: (text: string) => void;
  /** 自定义快捷操作 */
  actions?: QuickAction[];
}

/**
 * 欢迎界面组件
 * 显示欢迎消息和快捷操作按钮
 */
export const WelcomeView: React.FC<WelcomeViewProps> = ({
  disabled,
  onSend,
  actions = defaultActions,
}) => {
  const styles = useStyles();

  const getIcon = (iconType: QuickAction["icon"]) => {
    switch (iconType) {
      case "summarize":
        return <TableSimpleRegular className={styles.actionIcon} />;
      case "formula":
        return <MathFormulaRegular className={styles.actionIcon} />;
      case "chart":
        return <ChartMultipleRegular className={styles.actionIcon} />;
      case "analyze":
        return <SparkleRegular className={styles.actionIcon} />;
    }
  };

  return (
    <div className={styles.welcomeState}>
      <div className={styles.welcomeIcon}>
        <SparkleRegular />
      </div>
      <div className={styles.welcomeTitle}>Excel 智能助手</div>
      <div className={styles.welcomeSubtitle}>选中数据区域，告诉我你想做什么</div>

      <div className={styles.quickActionsGrid}>
        {actions.map((action) => (
          <button
            key={action.id}
            className={`${styles.actionCard} ${disabled ? styles.actionCardDisabled : ""}`}
            onClick={() => !disabled && onSend(action.prompt)}
            disabled={disabled}
          >
            {getIcon(action.icon)}
            <span className={styles.actionLabel}>{action.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
};

export default WelcomeView;
