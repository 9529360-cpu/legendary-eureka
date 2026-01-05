/**
 * PreviewConfirmDialog - 预览确认对话框组件 v2.0
 * @file src/taskpane/components/PreviewConfirmDialog.tsx
 * @description v2.9.39 升级版，支持风险等级显示和详细操作预览
 */
import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Body2,
  Badge,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  WarningRegular,
  InfoRegular,
  ShieldCheckmarkRegular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  textMuted: {
    color: tokens.colorNeutralForeground3,
  },
  previewBox: {
    marginTop: "12px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "8px",
    whiteSpace: "pre-wrap",
    fontFamily: "Consolas, monospace",
    fontSize: "12px",
    maxHeight: "200px",
    overflow: "auto",
    lineHeight: "1.5",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "8px",
  },
  riskBadge: {
    marginLeft: "auto",
  },
  detailsSection: {
    marginTop: "12px",
    padding: "8px 12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "6px",
    fontSize: "13px",
  },
  detailRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "4px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    "&:last-child": {
      borderBottom: "none",
    },
  },
  detailLabel: {
    color: tokens.colorNeutralForeground3,
  },
  detailValue: {
    fontWeight: 500,
  },
  warningText: {
    color: tokens.colorPaletteYellowForeground1,
    display: "flex",
    alignItems: "center",
    gap: "4px",
    marginTop: "8px",
    fontSize: "12px",
  },
});

export interface OperationDetails {
  /** 操作类型 */
  operationType: string;
  /** 目标范围 */
  targetRange?: string;
  /** 影响的单元格数量 */
  cellCount?: number;
  /** 是否可撤销 */
  reversible?: boolean;
  /** 风险等级 */
  riskLevel?: "low" | "medium" | "high";
  /** 额外信息 */
  additionalInfo?: string;
}

export interface PreviewConfirmDialogProps {
  /** 对话框是否打开 */
  open: boolean;
  /** 对话框打开状态变更回调 */
  onOpenChange: (open: boolean) => void;
  /** 预览消息内容 */
  previewMessage: string;
  /** 操作详情 */
  operationDetails?: OperationDetails;
  /** 确认回调 */
  onConfirm: () => void;
  /** 取消回调 */
  onCancel: () => void;
}

/**
 * 预览确认对话框组件
 * 用于在执行 Agent 操作前显示预览并要求用户确认
 */
export const PreviewConfirmDialog: React.FC<PreviewConfirmDialogProps> = ({
  open,
  onOpenChange,
  previewMessage,
  operationDetails,
  onConfirm,
  onCancel,
}) => {
  const styles = useStyles();

  const handleOpenChange = React.useCallback(
    (_: unknown, data: { open: boolean }) => {
      if (!data.open) {
        onCancel();
      }
      onOpenChange(data.open);
    },
    [onCancel, onOpenChange]
  );

  const handleCancel = React.useCallback(() => {
    onCancel();
    onOpenChange(false);
  }, [onCancel, onOpenChange]);

  const handleConfirm = React.useCallback(() => {
    onConfirm();
    onOpenChange(false);
  }, [onConfirm, onOpenChange]);

  // 风险等级显示
  const getRiskBadge = () => {
    const level = operationDetails?.riskLevel || "low";
    switch (level) {
      case "high":
        return <Badge color="danger" icon={<WarningRegular />}>高风险</Badge>;
      case "medium":
        return <Badge color="warning" icon={<InfoRegular />}>中风险</Badge>;
      default:
        return <Badge color="success" icon={<ShieldCheckmarkRegular />}>安全</Badge>;
    }
  };

  // 操作类型中文映射
  const getOperationTypeName = (type: string): string => {
    const typeMap: Record<string, string> = {
      "excel_write_range": "写入数据",
      "excel_write_cell": "写入单元格",
      "excel_clear": "清空内容",
      "excel_format_range": "格式化",
      "excel_set_formula": "设置公式",
      "excel_sort": "排序数据",
      "excel_create_chart": "创建图表",
      "excel_delete_sheet": "删除工作表",
    };
    return typeMap[type] || type;
  };

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>
            <div className={styles.header}>
              <span>确认操作</span>
              <div className={styles.riskBadge}>{getRiskBadge()}</div>
            </div>
          </DialogTitle>
          <DialogContent>
            <Body2 className={styles.textMuted}>即将执行以下操作：</Body2>
            <div className={styles.previewBox}>{previewMessage}</div>
            
            {operationDetails && (
              <div className={styles.detailsSection}>
                {operationDetails.operationType && (
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>操作类型</span>
                    <span className={styles.detailValue}>
                      {getOperationTypeName(operationDetails.operationType)}
                    </span>
                  </div>
                )}
                {operationDetails.targetRange && (
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>目标范围</span>
                    <span className={styles.detailValue}>{operationDetails.targetRange}</span>
                  </div>
                )}
                {operationDetails.cellCount !== undefined && (
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>影响单元格</span>
                    <span className={styles.detailValue}>{operationDetails.cellCount} 个</span>
                  </div>
                )}
                {operationDetails.reversible !== undefined && (
                  <div className={styles.detailRow}>
                    <span className={styles.detailLabel}>可撤销</span>
                    <span className={styles.detailValue}>
                      {operationDetails.reversible ? "✅ 是" : "❌ 否"}
                    </span>
                  </div>
                )}
              </div>
            )}
            
            {operationDetails?.riskLevel === "high" && (
              <div className={styles.warningText}>
                <WarningRegular />
                <span>此操作风险较高，请确认后再执行</span>
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={handleCancel}>
              取消
            </Button>
            <Button appearance="primary" onClick={handleConfirm}>
              确认执行
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default PreviewConfirmDialog;
