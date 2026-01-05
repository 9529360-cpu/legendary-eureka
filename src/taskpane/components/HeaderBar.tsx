/**
 * HeaderBar - 顶部状态栏组件
 * @file src/taskpane/components/HeaderBar.tsx
 * @description v2.9.14 移除 Tooltip 防止闪烁
 */
import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Button,
  Caption1,
  Spinner,
} from "@fluentui/react-components";
import {
  SettingsRegular,
  ArrowSyncRegular,
  ArrowUndoRegular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  statusBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    ...shorthands.padding("12px", "16px"),
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    minHeight: "48px",
  },
  statusLeft: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  statusIndicator: {
    width: "8px",
    height: "8px",
    ...shorthands.borderRadius("50%"),
    backgroundColor: tokens.colorPaletteGreenBackground3,
  },
  statusIndicatorOffline: {
    backgroundColor: tokens.colorPaletteRedBackground3,
  },
  brandText: {
    fontSize: "14px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
  },
  flexCenter: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  textMuted: {
    color: tokens.colorNeutralForeground3,
    fontSize: "12px",
  },
  workbookInfo: {
    marginLeft: "8px",
    color: tokens.colorNeutralForeground3,
    fontSize: "12px",
    cursor: "default",
  },
  actionButton: {
    minWidth: "32px",
    minHeight: "32px",
    width: "32px",
    height: "32px",
  },
});

export interface WorkbookSummary {
  sheetCount: number;
  tableCount: number;
  formulaCount: number;
  qualityScore: number;
}

// ========== 独立的操作按钮组件（完全稳定，无 Tooltip） ==========
interface ActionButtonsProps {
  onRefreshWorkbook: () => void;
  onUndo: () => void;
  onOpenSettings: () => void;
}

const ActionButtons = React.memo<ActionButtonsProps>(
  function ActionButtons({ onRefreshWorkbook, onUndo, onOpenSettings }) {
    const styles = useStyles();
    
    return (
      <>
        <Button
          className={styles.actionButton}
          appearance="subtle"
          size="small"
          icon={<ArrowSyncRegular />}
          onClick={onRefreshWorkbook}
          title="刷新"
        />
        <Button
          className={styles.actionButton}
          appearance="subtle"
          size="small"
          icon={<ArrowUndoRegular />}
          onClick={onUndo}
          title="撤销"
        />
        <Button
          className={styles.actionButton}
          appearance="subtle"
          size="small"
          icon={<SettingsRegular />}
          onClick={onOpenSettings}
          title="设置"
        />
      </>
    );
  }
);

export interface HeaderBarProps {
  backendHealthy: boolean | null;
  workbookSummary?: WorkbookSummary;
  isScanning: boolean;
  scanProgress: number;
  selectionAddress?: string;
  undoCount: number;
  apiKeyValid: boolean;
  onRefreshWorkbook: () => void;
  onUndo: () => void;
  onOpenSettings: () => void;
}

/**
 * 顶部状态栏组件
 */
export const HeaderBar: React.FC<HeaderBarProps> = ({
  backendHealthy,
  workbookSummary,
  isScanning,
  scanProgress,
  selectionAddress,
  onRefreshWorkbook,
  onUndo,
  onOpenSettings,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.statusBar}>
      <div className={styles.statusLeft}>
        <div
          className={`${styles.statusIndicator} ${
            backendHealthy === false ? styles.statusIndicatorOffline : ""
          }`}
        />
        <span className={styles.brandText}>Copilot</span>

        {workbookSummary && (
          <Caption1 className={styles.workbookInfo} title={`${workbookSummary.sheetCount} 表 · ${workbookSummary.tableCount} 数据表 · ${workbookSummary.formulaCount} 公式`}>
            📄 {workbookSummary.sheetCount}表 · ⭐ {workbookSummary.qualityScore}分
          </Caption1>
        )}

        {isScanning && (
          <Caption1
            style={{
              marginLeft: "8px",
              display: "flex",
              alignItems: "center",
              gap: "4px",
            }}
          >
            <Spinner size="extra-tiny" /> 扫描中 {scanProgress}%
          </Caption1>
        )}
      </div>

      <div className={styles.flexCenter}>
        {selectionAddress && (
          <Caption1 className={styles.textMuted}>{selectionAddress}</Caption1>
        )}
        <ActionButtons
          onRefreshWorkbook={onRefreshWorkbook}
          onUndo={onUndo}
          onOpenSettings={onOpenSettings}
        />
      </div>
    </div>
  );
};

export default React.memo(HeaderBar);
