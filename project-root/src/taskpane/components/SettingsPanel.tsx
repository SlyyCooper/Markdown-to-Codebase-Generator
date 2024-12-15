import * as React from "react";
import {
  Dialog,
  DialogBody,
  DialogContent,
  DialogSurface,
  Button,
  makeStyles,
  Input,
  Label,
  tokens,
  shorthands,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";

interface SettingsPanelProps {
  onClose: () => void;
  apiKey: string;
  onApiKeyChange: (apiKey: string) => void;
}

const useStyles = makeStyles({
  dialogSurface: {
    width: "400px",
    maxWidth: "90vw",
    ...shorthands.borderRadius("12px"),
    boxShadow: tokens.shadow64,
  },
  dialogContent: {
    display: "flex",
    flexDirection: "column",
    gap: "20px",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    ...shorthands.margin("0", "0", "16px"),
  },
  title: {
    fontSize: tokens.fontSizeBase600,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  closeButton: {
    ...shorthands.margin("0"),
    ...shorthands.padding("4px"),
    color: tokens.colorNeutralForeground3,
    "&:hover": {
      color: tokens.colorNeutralForeground2,
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  inputContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  label: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  input: {
    width: "100%",
    ...shorthands.borderRadius("4px"),
    "&:focus-within": {
      ...shorthands.borderColor(tokens.colorBrandBackground),
    },
  },
  saveButton: {
    marginTop: "16px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    "&:hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
  },
});

const DEFAULT_API_KEY = "sk-...";

const SettingsPanel: React.FC<SettingsPanelProps> = ({ onClose, apiKey, onApiKeyChange }) => {
  const styles = useStyles();
  const [tempApiKey, setTempApiKey] = React.useState(apiKey || DEFAULT_API_KEY);

  const handleSave = () => {
    onApiKeyChange(tempApiKey);
    onClose();
  };

  return (
    <Dialog open={true} onOpenChange={(_, { open }) => !open && onClose()}>
      <DialogSurface className={styles.dialogSurface}>
        <DialogBody>
          <DialogContent className={styles.dialogContent}>
            <div className={styles.header}>
              <h2 className={styles.title}>Settings</h2>
              <Button
                icon={<Dismiss24Regular />}
                appearance="subtle"
                className={styles.closeButton}
                onClick={onClose}
              />
            </div>
            <div className={styles.inputContainer}>
              <Label htmlFor="apiKeyInput" className={styles.label}>
                API Key
              </Label>
              <Input
                id="apiKeyInput"
                type="password"
                value={tempApiKey}
                onChange={(e) => setTempApiKey(e.target.value)}
                placeholder="Enter your API key"
                className={styles.input}
              />
            </div>
            <Button className={styles.saveButton} onClick={handleSave}>
              Save Changes
            </Button>
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default SettingsPanel;
