/**
 * ApiConfigDialog - API 配置对话框组件
 * @file src/taskpane/components/ApiConfigDialog.tsx
 * @description v2.9.8 从 App.tsx 提取，包含 API 密钥配置界面
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
  Input,
  Field,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  CheckmarkCircleRegular,
  DismissCircleRegular,
  ArrowSyncRegular,
} from "@fluentui/react-icons";
import type { ApiKeyStatus } from "../../services/ApiService";

const useStyles = makeStyles({
  dialogField: {
    marginBottom: "8px",
  },
  fullWidth: {
    width: "100%",
  },
  marginTop8: {
    marginTop: "8px",
  },
  marginTop16: {
    marginTop: "16px",
  },
  flexBetween: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  flexCenter: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
  },
  textSuccess: {
    color: tokens.colorPaletteGreenForeground1,
  },
  textError: {
    color: tokens.colorPaletteRedForeground1,
  },
  textMuted: {
    color: tokens.colorNeutralForeground3,
  },
});

export interface ApiConfigDialogProps {
  /** 对话框是否打开 */
  open: boolean;
  /** 对话框打开状态变更回调 */
  onOpenChange: (open: boolean) => void;
  /** API 密钥输入值 */
  apiKeyInput: string;
  /** API 密钥输入变更回调 */
  onApiKeyInputChange: (value: string) => void;
  /** API 密钥状态 */
  apiKeyStatus: ApiKeyStatus | null;
  /** 后端连接状态 */
  backendHealthy: boolean | null;
  /** 正在检查后端状态 */
  backendChecking: boolean;
  /** API 密钥操作正在进行 */
  apiKeyBusy: boolean;
  /** 刷新后端状态回调 */
  onRefreshBackend: () => void;
  /** 保存 API 密钥回调 */
  onSaveApiKey: () => void;
  /** 清除 API 密钥回调 */
  onClearApiKey: () => void;
}

/**
 * API 配置对话框组件
 * 用于配置和管理 DeepSeek API 密钥
 */
export const ApiConfigDialog: React.FC<ApiConfigDialogProps> = ({
  open,
  onOpenChange,
  apiKeyInput,
  onApiKeyInputChange,
  apiKeyStatus,
  backendHealthy,
  backendChecking,
  apiKeyBusy,
  onRefreshBackend,
  onSaveApiKey,
  onClearApiKey,
}) => {
  const styles = useStyles();

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>配置 API</DialogTitle>
          <DialogContent>
            <Field label="DeepSeek API 密钥" className={styles.dialogField}>
              <Input
                type="password"
                value={apiKeyInput}
                onChange={(e) => onApiKeyInputChange(e.target.value)}
                placeholder="sk-..."
                className={styles.fullWidth}
              />
            </Field>

            <Divider className={styles.marginTop16} />

            <div className={`${styles.flexBetween} ${styles.marginTop16}`}>
              <div className={styles.flexCenter}>
                {backendHealthy ? (
                  <CheckmarkCircleRegular className={styles.textSuccess} />
                ) : (
                  <DismissCircleRegular className={styles.textError} />
                )}
                <Caption1 className={styles.textMuted}>
                  后端 {backendHealthy ? "已连接" : "未连接"}
                </Caption1>
              </div>
              <Button
                appearance="subtle"
                size="small"
                icon={<ArrowSyncRegular />}
                onClick={onRefreshBackend}
                disabled={backendChecking}
              />
            </div>

            {apiKeyStatus?.configured && (
              <div className={styles.marginTop8}>
                <Caption1 className={styles.textMuted}>
                  当前: {apiKeyStatus.maskedKey}
                  {apiKeyStatus.isValid ? " ✓" : " ✗"}
                </Caption1>
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>
              取消
            </Button>
            {apiKeyStatus?.configured && (
              <Button
                appearance="secondary"
                onClick={onClearApiKey}
                disabled={apiKeyBusy || backendHealthy === false}
              >
                清除
              </Button>
            )}
            <Button
              appearance="primary"
              onClick={onSaveApiKey}
              disabled={!apiKeyInput.trim() || apiKeyBusy || backendHealthy === false}
            >
              保存
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default ApiConfigDialog;
