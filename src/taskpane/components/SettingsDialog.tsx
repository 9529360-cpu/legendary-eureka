import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Body1,
  Button,
  Switch,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Field,
  Input,
  Caption1,
} from "@fluentui/react-components";
import {
  SettingsRegular,
  WeatherMoonRegular,
  WeatherSunnyRegular,
  KeyRegular,
  InfoRegular,
} from "@fluentui/react-icons";

interface SettingsDialogProps {
  isOpen: boolean;
  onClose: () => void;
  isDarkTheme: boolean;
  onThemeChange: (isDark: boolean) => void;
  apiKey?: string;
  onApiKeyChange?: (key: string) => void;
  apiKeyStatus?: "valid" | "invalid" | "unchecked";
}

const useStyles = makeStyles({
  dialog: {
    maxWidth: "480px",
  },
  section: {
    marginBottom: "24px",
  },
  sectionTitle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    marginBottom: "12px",
    fontWeight: "600",
  },
  settingRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    ...shorthands.padding("12px"),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius("8px"),
    marginBottom: "8px",
  },
  settingLabel: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  themeToggle: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  apiKeyInput: {
    width: "100%",
  },
  apiKeyStatus: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    marginTop: "8px",
  },
  statusValid: {
    color: tokens.colorPaletteGreenForeground1,
  },
  statusInvalid: {
    color: tokens.colorPaletteRedForeground1,
  },
  infoText: {
    color: tokens.colorNeutralForeground4,
    fontSize: "12px",
    marginTop: "8px",
  },
});

const SettingsDialog: React.FC<SettingsDialogProps> = ({
  isOpen,
  onClose,
  isDarkTheme,
  onThemeChange,
  apiKey = "",
  onApiKeyChange,
  apiKeyStatus = "unchecked",
}) => {
  const styles = useStyles();
  const [localApiKey, setLocalApiKey] = React.useState(apiKey);

  React.useEffect(() => {
    setLocalApiKey(apiKey);
  }, [apiKey]);

  const handleSave = () => {
    if (onApiKeyChange && localApiKey !== apiKey) {
      onApiKeyChange(localApiKey);
    }
    onClose();
  };

  return (
    <Dialog open={isOpen} onOpenChange={(e, data) => !data.open && onClose()}>
      <DialogSurface className={styles.dialog}>
        <DialogTitle>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <SettingsRegular />
            设置
          </div>
        </DialogTitle>
        <DialogBody>
          <DialogContent>
            {/* 外观设置 */}
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                {isDarkTheme ? <WeatherMoonRegular /> : <WeatherSunnyRegular />}
                <span>外观</span>
              </div>
              <div className={styles.settingRow}>
                <div className={styles.settingLabel}>
                  <Body1>深色模式</Body1>
                  <Caption1>切换应用的主题外观</Caption1>
                </div>
                <div className={styles.themeToggle}>
                  <WeatherSunnyRegular />
                  <Switch
                    checked={isDarkTheme}
                    onChange={(e, data) => onThemeChange(data.checked)}
                  />
                  <WeatherMoonRegular />
                </div>
              </div>
            </div>

            {/* API 设置 */}
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                <KeyRegular />
                <span>API 配置</span>
              </div>
              <Field label="API 密钥" hint="用于连接 AI 服务">
                <Input
                  className={styles.apiKeyInput}
                  type="password"
                  placeholder="输入您的 API 密钥..."
                  value={localApiKey}
                  onChange={(e, data) => setLocalApiKey(data.value)}
                />
              </Field>
              {apiKeyStatus !== "unchecked" && (
                <div
                  className={`${styles.apiKeyStatus} ${
                    apiKeyStatus === "valid" ? styles.statusValid : styles.statusInvalid
                  }`}
                >
                  <InfoRegular />
                  <Caption1>
                    {apiKeyStatus === "valid" ? "API 密钥有效" : "API 密钥无效"}
                  </Caption1>
                </div>
              )}
              <Caption1 className={styles.infoText}>
                如果未设置 API 密钥，将使用内置的后端服务。
              </Caption1>
            </div>

            {/* 关于 */}
            <div className={styles.section}>
              <div className={styles.sectionTitle}>
                <InfoRegular />
                <span>关于</span>
              </div>
              <div className={styles.settingRow}>
                <div className={styles.settingLabel}>
                  <Body1>Excel 智能助手</Body1>
                  <Caption1>版本 2.0.0</Caption1>
                </div>
              </div>
            </div>
          </DialogContent>
        </DialogBody>
        <DialogActions>
          <Button appearance="secondary" onClick={onClose}>
            取消
          </Button>
          <Button appearance="primary" onClick={handleSave}>
            保存
          </Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
};

export default SettingsDialog;
