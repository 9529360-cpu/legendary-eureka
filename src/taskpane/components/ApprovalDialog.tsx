/**
 * ApprovalDialog - Agent 层审批确认弹窗 v1.0
 * 
 * 特性：
 * 1. 显示 approvalId 便于追溯
 * 2. 展示风险等级和影响范围
 * 3. 确认/取消两个按钮
 * 4. 支持回滚策略展示
 * 5. 高风险操作特殊警告样式
 * 
 * 设计原则：
 * - 清晰展示"将改动什么"
 * - 明确标注"是否可撤销"
 * - 使用醒目颜色区分风险等级
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
  Text,
  Badge,
  Divider,
  makeStyles,
  tokens,
  mergeClasses,
} from "@fluentui/react-components";
import {
  WarningFilled,
  ShieldErrorFilled,
  InfoFilled,
  CheckmarkCircleFilled,
  DismissCircleFilled,
  ArrowUndoRegular,
  DocumentTableRegular,
} from "@fluentui/react-icons";
import type { ApprovalRequest, RiskLevel } from "../../agent/ApprovalManager";

// ==================== 样式定义 ====================

const useStyles = makeStyles({
  surface: {
    maxWidth: "480px",
  },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "12px",
  },
  headerRight: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  approvalId: {
    fontFamily: "Consolas, monospace",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "2px 8px",
    borderRadius: "4px",
  },
  riskBadgeCritical: {
    backgroundColor: tokens.colorPaletteRedBackground3,
    color: tokens.colorPaletteRedForeground1,
  },
  riskBadgeHigh: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
  },
  riskBadgeMedium: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorPaletteYellowForeground2,
  },
  riskBadgeLow: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    color: tokens.colorPaletteGreenForeground1,
  },
  contentBox: {
    marginTop: "16px",
  },
  operationSection: {
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "8px",
    padding: "12px 16px",
    marginBottom: "12px",
  },
  operationTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: 600,
    marginBottom: "8px",
  },
  detailGrid: {
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    gap: "8px 16px",
    fontSize: "13px",
  },
  detailLabel: {
    color: tokens.colorNeutralForeground3,
  },
  detailValue: {
    fontWeight: 500,
  },
  warningBox: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    backgroundColor: tokens.colorPaletteRedBackground1,
    borderRadius: "6px",
    padding: "10px 12px",
    marginTop: "12px",
    border: `1px solid ${tokens.colorPaletteRedBorder1}`,
  },
  warningBoxMedium: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    border: `1px solid ${tokens.colorPaletteYellowBorder1}`,
  },
  warningIcon: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: "20px",
    flexShrink: 0,
  },
  warningIconMedium: {
    color: tokens.colorPaletteYellowForeground2,
  },
  warningText: {
    fontSize: "13px",
    lineHeight: "1.4",
  },
  reversibleInfo: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    marginTop: "12px",
    fontSize: "12px",
  },
  reversibleYes: {
    color: tokens.colorPaletteGreenForeground1,
  },
  reversibleNo: {
    color: tokens.colorPaletteRedForeground1,
  },
  divider: {
    margin: "16px 0",
  },
  actions: {
    paddingTop: "8px",
  },
  cancelButton: {
    minWidth: "80px",
  },
  confirmButton: {
    minWidth: "100px",
  },
  confirmButtonDanger: {
    backgroundColor: tokens.colorPaletteRedBackground3,
    color: "white",
    ":hover": {
      backgroundColor: tokens.colorPaletteRedForeground1,
    },
  },
});

// ==================== Props 定义 ====================

export interface ApprovalDialogProps {
  /** 是否打开 */
  open: boolean;
  /** 审批请求数据 */
  approvalRequest: ApprovalRequest | null;
  /** 确认回调 */
  onConfirm: (approvalId: string) => void;
  /** 取消回调 */
  onCancel: (approvalId: string) => void;
  /** 关闭回调 */
  onClose: () => void;
}

// ==================== 组件实现 ====================

/**
 * Agent 层审批确认弹窗
 */
export const ApprovalDialog: React.FC<ApprovalDialogProps> = ({
  open,
  approvalRequest,
  onConfirm,
  onCancel,
  onClose,
}) => {
  const styles = useStyles();

  if (!approvalRequest) return null;

  const { approvalId, operationName, parameters, riskAssessment } = approvalRequest;
  const { riskLevel, impactDescription, reversible, estimatedImpact, reason } = riskAssessment;

  // 风险等级 Badge
  const getRiskBadge = () => {
    const config: Record<RiskLevel, { label: string; className: string; icon: React.ReactElement }> = {
      critical: { label: "严重风险", className: styles.riskBadgeCritical, icon: <ShieldErrorFilled /> },
      high: { label: "高风险", className: styles.riskBadgeHigh, icon: <WarningFilled /> },
      medium: { label: "中风险", className: styles.riskBadgeMedium, icon: <InfoFilled /> },
      low: { label: "低风险", className: styles.riskBadgeLow, icon: <CheckmarkCircleFilled /> },
    };
    const c = config[riskLevel];
    return (
      <Badge className={c.className} icon={c.icon} appearance="filled">
        {c.label}
      </Badge>
    );
  };

  // 操作显示名称
  const getOperationDisplayName = (name: string): string => {
    const displayNames: Record<string, string> = {
      'delete_rows': '删除行',
      'delete_row': '删除行',
      'delete_columns': '删除列',
      'delete_column': '删除列',
      'delete_sheet': '删除工作表',
      'clear_range': '清空区域',
      'clear_all': '清空全部',
      'batch_update': '批量更新',
      'batch_write': '批量写入',
      'batch_formula': '批量公式',
      'fill_formula': '填充公式',
      'remove_duplicates': '删除重复项',
      'protect_sheet': '保护工作表',
      'unprotect_sheet': '取消保护',
      'write_range': '写入数据',
      'set_formula': '设置公式',
      'sort_range': '排序数据',
    };
    return displayNames[name] || name;
  };

  // 处理确认
  const handleConfirm = () => {
    onConfirm(approvalId);
    onClose();
  };

  // 处理取消
  const handleCancel = () => {
    onCancel(approvalId);
    onClose();
  };

  const isHighRisk = riskLevel === 'high' || riskLevel === 'critical';
  const isMediumRisk = riskLevel === 'medium';

  return (
    <Dialog open={open} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle>
            <div className={styles.header}>
              <span>确认操作</span>
              <div className={styles.headerRight}>
                {getRiskBadge()}
                <span className={styles.approvalId}>{approvalId}</span>
              </div>
            </div>
          </DialogTitle>

          <DialogContent className={styles.contentBox}>
            {/* 操作信息区 */}
            <div className={styles.operationSection}>
              <div className={styles.operationTitle}>
                <DocumentTableRegular />
                <span>{getOperationDisplayName(operationName)}</span>
              </div>
              
              <div className={styles.detailGrid}>
                {/* 目标范围 */}
                {Boolean(parameters.range || parameters.address) && (
                  <>
                    <span className={styles.detailLabel}>目标范围</span>
                    <span className={styles.detailValue}>
                      {(parameters.range || parameters.address) as string}
                    </span>
                  </>
                )}
                
                {/* 影响行数 */}
                {Boolean((estimatedImpact as { rowCount?: number })?.rowCount) && (
                  <>
                    <span className={styles.detailLabel}>影响行数</span>
                    <span className={styles.detailValue}>
                      约 {(estimatedImpact as { rowCount?: number }).rowCount} 行
                    </span>
                  </>
                )}
                
                {/* 影响单元格数 */}
                {Boolean((estimatedImpact as { cellCount?: number })?.cellCount) && (
                  <>
                    <span className={styles.detailLabel}>影响单元格</span>
                    <span className={styles.detailValue}>
                      约 {(estimatedImpact as { cellCount?: number }).cellCount} 个
                    </span>
                  </>
                )}

                {/* 工作表 */}
                {Boolean(parameters.sheetName) && (
                  <>
                    <span className={styles.detailLabel}>工作表</span>
                    <span className={styles.detailValue}>
                      {parameters.sheetName as string}
                    </span>
                  </>
                )}
              </div>
            </div>

            {/* 风险警告框 */}
            {(isHighRisk || isMediumRisk) && (
              <div className={mergeClasses(
                styles.warningBox,
                isMediumRisk && styles.warningBoxMedium
              )}>
                <WarningFilled className={mergeClasses(
                  styles.warningIcon,
                  isMediumRisk && styles.warningIconMedium
                )} />
                <div className={styles.warningText}>
                  <Text weight="semibold">
                    {isHighRisk ? '⚠️ 高风险操作' : '⚡ 请注意'}
                  </Text>
                  <br />
                  {impactDescription}
                  {reason && (
                    <>
                      <br />
                      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                        触发原因: {reason}
                      </Text>
                    </>
                  )}
                </div>
              </div>
            )}

            {/* 可撤销信息 */}
            <div className={styles.reversibleInfo}>
              <ArrowUndoRegular />
              <span>可撤销：</span>
              {reversible ? (
                <span className={styles.reversibleYes}>
                  <CheckmarkCircleFilled /> 是，操作后可恢复
                </span>
              ) : (
                <span className={styles.reversibleNo}>
                  <DismissCircleFilled /> 否，此操作不可逆
                </span>
              )}
            </div>
          </DialogContent>

          <Divider className={styles.divider} />

          <DialogActions className={styles.actions}>
            <Button 
              appearance="secondary" 
              onClick={handleCancel}
              className={styles.cancelButton}
            >
              取消
            </Button>
            <Button 
              appearance="primary" 
              onClick={handleConfirm}
              className={mergeClasses(
                styles.confirmButton,
                isHighRisk && styles.confirmButtonDanger
              )}
            >
              {isHighRisk ? '确认执行' : '确认'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default ApprovalDialog;
