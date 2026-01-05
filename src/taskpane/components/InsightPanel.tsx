/**
 * InsightPanel - 数据洞察面板组件
 * @file src/taskpane/components/InsightPanel.tsx
 * @description v2.9.8 从 App.tsx 提取，显示数据摘要和智能建议
 */
import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Spinner,
} from "@fluentui/react-components";
import {
  LightbulbRegular,
  DataBarVerticalRegular,
  ChartMultipleRegular,
  BroomRegular,
  MathFormulaRegular,
  PaintBrushRegular,
  FlashRegular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  insightPanel: {
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius("0", "0", "12px", "12px"),
    ...shorthands.padding("16px"),
    ...shorthands.margin("0", "16px"),
    marginBottom: "8px",
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    borderTop: "none",
  },
  insightHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "12px",
  },
  insightHeaderLeft: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  insightHeaderIcon: {
    width: "32px",
    height: "32px",
    ...shorthands.borderRadius("8px"),
    backgroundColor: tokens.colorBrandBackground,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    "& svg": {
      fontSize: "16px",
      color: tokens.colorNeutralForegroundOnBrand,
    },
  },
  insightTitle: {
    fontSize: "14px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
  },
  insightSubtitle: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
  },
  qualityBadge: {
    ...shorthands.padding("4px", "8px"),
    ...shorthands.borderRadius("12px"),
    fontSize: "11px",
    fontWeight: 500,
  },
  qualityGood: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  qualityWarning: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    color: tokens.colorPaletteYellowForeground1,
  },
  qualityPoor: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  analyzingState: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    ...shorthands.padding("16px"),
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
  },
  dataSummaryGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(3, 1fr)",
    gap: "12px",
    marginBottom: "12px",
  },
  summaryCell: {
    textAlign: "center",
    ...shorthands.padding("8px"),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius("8px"),
  },
  summaryValue: {
    fontSize: "18px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
  },
  summaryLabel: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
  },
  suggestionsContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  suggestionCard: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    ...shorthands.padding("10px", "12px"),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.borderRadius("8px"),
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke2),
    cursor: "pointer",
    transitionProperty: "background-color, border-color",
    transitionDuration: "0.15s",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  suggestionIcon: {
    width: "32px",
    height: "32px",
    ...shorthands.borderRadius("8px"),
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    "& svg": {
      fontSize: "16px",
    },
  },
  suggestionIconAnalyze: {
    backgroundColor: tokens.colorPaletteBlueBorderActive,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  suggestionIconChart: {
    backgroundColor: tokens.colorPaletteGreenBorderActive,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  suggestionIconClean: {
    backgroundColor: tokens.colorPaletteYellowBorderActive,
    color: tokens.colorNeutralForeground1,
  },
  suggestionIconFormula: {
    backgroundColor: tokens.colorPalettePurpleBorderActive,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  suggestionIconFormat: {
    backgroundColor: tokens.colorPalettePinkBorderActive,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  suggestionContent: {
    flex: 1,
  },
  suggestionTitle: {
    fontSize: "13px",
    fontWeight: 500,
    color: tokens.colorNeutralForeground1,
  },
  suggestionDesc: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
  },
  suggestionArrow: {
    fontSize: "16px",
    color: tokens.colorNeutralForeground3,
  },
});

export interface DataSummary {
  rowCount: number;
  columnCount: number;
  numericColumns: number;
  qualityScore: number;
}

export interface ProactiveSuggestion {
  id: string;
  title: string;
  description: string;
  icon: "analyze" | "chart" | "clean" | "formula" | "format";
  action: () => void;
}

export interface InsightPanelProps {
  /** 是否正在分析 */
  isAnalyzing: boolean;
  /** 数据摘要 */
  dataSummary?: DataSummary;
  /** 当前选区地址 */
  selectionAddress?: string;
  /** 智能建议列表 */
  suggestions: ProactiveSuggestion[];
}

/**
 * 数据洞察面板组件
 * 显示数据摘要和智能建议
 */
export const InsightPanel: React.FC<InsightPanelProps> = ({
  isAnalyzing,
  dataSummary,
  selectionAddress,
  suggestions,
}) => {
  const styles = useStyles();

  const getQualityLabel = (score: number): string => {
    if (score >= 90) return "✓ 优质";
    if (score >= 70) return "△ 一般";
    return "! 待改善";
  };

  const getQualityClass = (score: number): string => {
    if (score >= 90) return styles.qualityGood;
    if (score >= 70) return styles.qualityWarning;
    return styles.qualityPoor;
  };

  const getIconElement = (icon: ProactiveSuggestion["icon"]) => {
    switch (icon) {
      case "analyze":
        return <DataBarVerticalRegular />;
      case "chart":
        return <ChartMultipleRegular />;
      case "clean":
        return <BroomRegular />;
      case "formula":
        return <MathFormulaRegular />;
      case "format":
        return <PaintBrushRegular />;
    }
  };

  const getIconClass = (icon: ProactiveSuggestion["icon"]): string => {
    switch (icon) {
      case "analyze":
        return styles.suggestionIconAnalyze;
      case "chart":
        return styles.suggestionIconChart;
      case "clean":
        return styles.suggestionIconClean;
      case "formula":
        return styles.suggestionIconFormula;
      case "format":
        return styles.suggestionIconFormat;
    }
  };

  return (
    <div className={styles.insightPanel}>
      {/* 头部 */}
      <div className={styles.insightHeader}>
        <div className={styles.insightHeaderLeft}>
          <div className={styles.insightHeaderIcon}>
            <LightbulbRegular />
          </div>
          <div>
            <div className={styles.insightTitle}>
              {isAnalyzing ? "正在分析..." : "数据洞察"}
            </div>
            {dataSummary && !isAnalyzing && selectionAddress && (
              <div className={styles.insightSubtitle}>已选中 {selectionAddress}</div>
            )}
          </div>
        </div>
        {dataSummary && (
          <div className={`${styles.qualityBadge} ${getQualityClass(dataSummary.qualityScore)}`}>
            {getQualityLabel(dataSummary.qualityScore)}
          </div>
        )}
      </div>

      {isAnalyzing ? (
        <div className={styles.analyzingState}>
          <Spinner size="tiny" />
          <span>正在分析选中的数据...</span>
        </div>
      ) : (
        dataSummary && (
          <>
            {/* 数据摘要网格 */}
            <div className={styles.dataSummaryGrid}>
              <div className={styles.summaryCell}>
                <div className={styles.summaryValue}>{dataSummary.rowCount}</div>
                <div className={styles.summaryLabel}>行</div>
              </div>
              <div className={styles.summaryCell}>
                <div className={styles.summaryValue}>{dataSummary.columnCount}</div>
                <div className={styles.summaryLabel}>列</div>
              </div>
              <div className={styles.summaryCell}>
                <div className={styles.summaryValue}>{dataSummary.numericColumns}</div>
                <div className={styles.summaryLabel}>数值列</div>
              </div>
            </div>

            {/* 智能建议 */}
            {suggestions.length > 0 && (
              <div className={styles.suggestionsContainer}>
                {suggestions.map((suggestion) => (
                  <div
                    key={suggestion.id}
                    className={styles.suggestionCard}
                    onClick={suggestion.action}
                  >
                    <div className={`${styles.suggestionIcon} ${getIconClass(suggestion.icon)}`}>
                      {getIconElement(suggestion.icon)}
                    </div>
                    <div className={styles.suggestionContent}>
                      <div className={styles.suggestionTitle}>{suggestion.title}</div>
                      <div className={styles.suggestionDesc}>{suggestion.description}</div>
                    </div>
                    <FlashRegular className={styles.suggestionArrow} />
                  </div>
                ))}
              </div>
            )}
          </>
        )
      )}
    </div>
  );
};

export default InsightPanel;
