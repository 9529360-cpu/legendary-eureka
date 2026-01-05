import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Title3,
  Body2,
  Caption1,
  Button,
  Badge,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  ProgressBar,
} from "@fluentui/react-components";
import {
  ChartMultipleRegular,
  DataBarVerticalRegular,
  LightbulbRegular,
  PlayRegular,
  ArrowSyncRegular,
} from "@fluentui/react-icons";

export interface DataInsight {
  id: string;
  type: "statistic" | "quality" | "trend" | "recommendation";
  title: string;
  description: string;
  value?: string | number;
  severity?: "info" | "warning" | "error" | "success";
  actionable?: boolean;
}

interface AnalysisPanelProps {
  insights: DataInsight[];
  isAnalyzing: boolean;
  dataQuality?: {
    score: number;
    rating: string;
    completeness: number;
    consistency: number;
    validRecords: number;
    totalRecords: number;
  };
  onRefreshAnalysis?: () => void;
  onApplyRecommendation?: (insight: DataInsight) => void;
}

const useStyles = makeStyles({
  panel: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground2,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  header: {
    ...shorthands.padding("16px"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  content: {
    flex: 1,
    overflow: "auto",
    ...shorthands.padding("16px"),
  },
  section: {
    marginBottom: "20px",
  },
  sectionTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "12px",
  },
  qualityCard: {
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.padding("16px"),
    ...shorthands.borderRadius("8px"),
    marginBottom: "16px",
  },
  qualityScore: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    marginBottom: "12px",
  },
  scoreValue: {
    fontSize: "28px",
    fontWeight: "bold",
    color: tokens.colorBrandForeground1,
  },
  qualityMetric: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  progressContainer: {
    width: "60%",
  },
  insightCard: {
    backgroundColor: tokens.colorNeutralBackground3,
    ...shorthands.padding("12px"),
    ...shorthands.borderRadius("6px"),
    marginBottom: "8px",
    cursor: "pointer",
    transition: "background-color 0.15s ease",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground4,
    },
  },
  insightHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "4px",
  },
  insightIcon: {
    marginRight: "8px",
  },
  emptyState: {
    textAlign: "center",
    ...shorthands.padding("32px"),
    color: tokens.colorNeutralForeground4,
  },
  loadingState: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    ...shorthands.padding("32px"),
  },
});

const _getInsightIcon = (type: DataInsight["type"]) => {
  switch (type) {
    case "statistic":
      return <DataBarVerticalRegular />;
    case "quality":
      return <ChartMultipleRegular />;
    case "trend":
      return <ChartMultipleRegular />;
    case "recommendation":
      return <LightbulbRegular />;
    default:
      return <LightbulbRegular />;
  }
};

const getSeverityColor = (severity?: DataInsight["severity"]) => {
  switch (severity) {
    case "success":
      return "success";
    case "warning":
      return "warning";
    case "error":
      return "danger";
    default:
      return "informative";
  }
};

const AnalysisPanel: React.FC<AnalysisPanelProps> = ({
  insights,
  isAnalyzing,
  dataQuality,
  onRefreshAnalysis,
  onApplyRecommendation,
}) => {
  const styles = useStyles();

  const statisticInsights = insights.filter((i) => i.type === "statistic");
  const qualityInsights = insights.filter((i) => i.type === "quality");
  const recommendationInsights = insights.filter((i) => i.type === "recommendation");

  return (
    <div className={styles.panel}>
      <div className={styles.header}>
        <Title3>数据分析</Title3>
        <Button
          appearance="subtle"
          size="small"
          icon={<ArrowSyncRegular />}
          onClick={onRefreshAnalysis}
          disabled={isAnalyzing}
        >
          刷新
        </Button>
      </div>

      <div className={styles.content}>
        {isAnalyzing ? (
          <div className={styles.loadingState}>
            <ProgressBar />
            <Caption1>正在分析数据...</Caption1>
          </div>
        ) : insights.length === 0 ? (
          <div className={styles.emptyState}>
            <ChartMultipleRegular style={{ fontSize: 48, marginBottom: 16 }} />
            <Body2>选择数据区域开始分析</Body2>
            <Caption1>选中 Excel 中的数据后，点击刷新按钮</Caption1>
          </div>
        ) : (
          <>
            {/* 数据质量概览 */}
            {dataQuality && (
              <div className={styles.section}>
                <div className={styles.sectionTitle}>
                  <ChartMultipleRegular />
                  <Body2>数据质量</Body2>
                </div>
                <div className={styles.qualityCard}>
                  <div className={styles.qualityScore}>
                    <span className={styles.scoreValue}>{dataQuality.score}</span>
                    <Badge appearance="filled" color={dataQuality.score >= 80 ? "success" : dataQuality.score >= 60 ? "warning" : "danger"}>
                      {dataQuality.rating}
                    </Badge>
                  </div>
                  <div className={styles.qualityMetric}>
                    <Caption1>完整性</Caption1>
                    <div className={styles.progressContainer}>
                      <ProgressBar value={dataQuality.completeness / 100} />
                    </div>
                    <Caption1>{dataQuality.completeness.toFixed(0)}%</Caption1>
                  </div>
                  <div className={styles.qualityMetric}>
                    <Caption1>一致性</Caption1>
                    <div className={styles.progressContainer}>
                      <ProgressBar value={dataQuality.consistency / 100} />
                    </div>
                    <Caption1>{dataQuality.consistency.toFixed(0)}%</Caption1>
                  </div>
                  <Caption1 style={{ marginTop: "8px" }}>
                    有效记录: {dataQuality.validRecords} / {dataQuality.totalRecords}
                  </Caption1>
                </div>
              </div>
            )}

            {/* 统计洞察 */}
            <Accordion collapsible defaultOpenItems={["statistics"]}>
              {statisticInsights.length > 0 && (
                <AccordionItem value="statistics">
                  <AccordionHeader>
                    <div className={styles.sectionTitle}>
                      <DataBarVerticalRegular />
                      <span>统计信息</span>
                      <Badge appearance="tint" size="small">
                        {statisticInsights.length}
                      </Badge>
                    </div>
                  </AccordionHeader>
                  <AccordionPanel>
                    {statisticInsights.map((insight) => (
                      <div key={insight.id} className={styles.insightCard}>
                        <div className={styles.insightHeader}>
                          <Body2>{insight.title}</Body2>
                          {insight.value !== undefined && (
                            <Badge appearance="filled" color="brand">
                              {insight.value}
                            </Badge>
                          )}
                        </div>
                        <Caption1>{insight.description}</Caption1>
                      </div>
                    ))}
                  </AccordionPanel>
                </AccordionItem>
              )}

              {qualityInsights.length > 0 && (
                <AccordionItem value="quality">
                  <AccordionHeader>
                    <div className={styles.sectionTitle}>
                      <ChartMultipleRegular />
                      <span>数据质量</span>
                      <Badge appearance="tint" size="small">
                        {qualityInsights.length}
                      </Badge>
                    </div>
                  </AccordionHeader>
                  <AccordionPanel>
                    {qualityInsights.map((insight) => (
                      <div key={insight.id} className={styles.insightCard}>
                        <div className={styles.insightHeader}>
                          <Body2>{insight.title}</Body2>
                          <Badge
                            appearance="filled"
                            color={getSeverityColor(insight.severity) as any}
                          >
                            {insight.severity === "success" ? "正常" : insight.severity === "warning" ? "警告" : "关注"}
                          </Badge>
                        </div>
                        <Caption1>{insight.description}</Caption1>
                      </div>
                    ))}
                  </AccordionPanel>
                </AccordionItem>
              )}

              {recommendationInsights.length > 0 && (
                <AccordionItem value="recommendations">
                  <AccordionHeader>
                    <div className={styles.sectionTitle}>
                      <LightbulbRegular />
                      <span>建议</span>
                      <Badge appearance="tint" size="small">
                        {recommendationInsights.length}
                      </Badge>
                    </div>
                  </AccordionHeader>
                  <AccordionPanel>
                    {recommendationInsights.map((insight) => (
                      <div
                        key={insight.id}
                        className={styles.insightCard}
                        onClick={() =>
                          insight.actionable &&
                          onApplyRecommendation &&
                          onApplyRecommendation(insight)
                        }
                      >
                        <div className={styles.insightHeader}>
                          <Body2>{insight.title}</Body2>
                          {insight.actionable && (
                            <Button
                              appearance="subtle"
                              size="small"
                              icon={<PlayRegular />}
                            >
                              应用
                            </Button>
                          )}
                        </div>
                        <Caption1>{insight.description}</Caption1>
                      </div>
                    ))}
                  </AccordionPanel>
                </AccordionItem>
              )}
            </Accordion>
          </>
        )}
      </div>
    </div>
  );
};

export default AnalysisPanel;
