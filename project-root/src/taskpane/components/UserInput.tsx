import * as React from "react";
import { Button, makeStyles, tokens, shorthands, Textarea } from "@fluentui/react-components";
import { Send24Regular } from "@fluentui/react-icons";

interface UserInputProps {
  onSendMessage: (text: string, taggedSheets: string[]) => void;
  worksheetNames: string[];
}

const useStyles = makeStyles({
  userInput: {
    display: "flex",
    padding: "10px",
    backgroundColor: tokens.colorNeutralBackground1,
    minHeight: "60px",
    boxShadow: tokens.shadow4,
    position: "relative",
    alignItems: "flex-end",
  },
  textAreaWrapper: {
    flexGrow: 1,
    marginRight: "10px",
    position: "relative",
  },
  textArea: {
    width: "100%",
    padding: "12px 16px",
    paddingRight: "40px",
    fontSize: "14px",
    backgroundColor: "transparent",
    borderRadius: "20px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    resize: "none",
    maxHeight: "150px",
    overflowY: "auto",
    lineHeight: "1.5",
    scrollbarWidth: "none",
    "&::-webkit-scrollbar": {
      display: "none",
    },
    "-ms-overflow-style": "none",
  },
  suggestion: {
    position: "absolute",
    pointerEvents: "none",
    color: tokens.colorNeutralForeground4,
    opacity: 0.7,
  },
  taggedWorksheet: {
    fontWeight: "bold",
    color: tokens.colorBrandForeground1,
  },
  button: {
    position: "absolute",
    right: "10px",
    bottom: "10px",
    backgroundColor: "transparent",
    color: tokens.colorBrandForeground1,
    borderRadius: "50%",
    width: "32px",
    height: "32px",
    minWidth: "32px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    border: "none",
    cursor: "pointer",
    ...shorthands.padding("0"),
    zIndex: 3,
  },
  icon: {
    fontSize: "1.2em",
  },
  textOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    padding: "12px 16px",
    pointerEvents: "none",
    whiteSpace: "pre-wrap",
    overflowWrap: "break-word",
    wordWrap: "break-word",
    color: "transparent",
    caretColor: tokens.colorNeutralForeground1,
  },
});

const UserInput: React.FC<UserInputProps> = ({ onSendMessage, worksheetNames }) => {
  const [inputText, setInputText] = React.useState("");
  const [suggestion, setSuggestion] = React.useState("");
  const [cursorPosition, setCursorPosition] = React.useState(0);
  const styles = useStyles();
  const textAreaRef = React.useRef<HTMLTextAreaElement>(null);

  const handleInputChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    const newText = event.target.value;
    setInputText(newText);
    setCursorPosition(event.target.selectionStart);
    checkForTagging(newText, event.target.selectionStart);
    adjustTextAreaHeight();
  };

  const checkForTagging = (text: string, cursorPos: number) => {
    const lastAtIndex = text.lastIndexOf("@", cursorPos - 1);
    if (lastAtIndex !== -1) {
      const tagText = text.slice(lastAtIndex + 1, cursorPos).toLowerCase();
      const matchingWorksheet = worksheetNames.find((name) => name.toLowerCase().startsWith(tagText));
      if (matchingWorksheet) {
        setSuggestion(matchingWorksheet.slice(tagText.length));
      } else {
        setSuggestion("");
      }
    } else {
      setSuggestion("");
    }
  };

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === "Tab" && suggestion) {
      event.preventDefault();
      const newText = inputText.slice(0, cursorPosition) + suggestion + inputText.slice(cursorPosition);
      setInputText(newText);
      setCursorPosition(cursorPosition + suggestion.length);
      setSuggestion("");
    }
  };

  const adjustTextAreaHeight = () => {
    if (textAreaRef.current) {
      textAreaRef.current.style.height = "auto";
      textAreaRef.current.style.height = `${textAreaRef.current.scrollHeight}px`;
    }
  };

  const handleSend = () => {
    if (inputText.trim()) {
      const taggedSheets = inputText.match(/@([^\s]+)/g)?.map((tag) => tag.slice(1)) || [];
      onSendMessage(inputText, taggedSheets);
      setInputText("");
      setSuggestion("");
      if (textAreaRef.current) {
        textAreaRef.current.style.height = "auto";
      }
    }
  };

  const handleKeyPress = (event: React.KeyboardEvent) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      handleSend();
    }
  };

  const renderInputText = () => {
    const parts = inputText.split(/(@[^\s]+)/);
    return parts.map((part, index) => {
      if (part.startsWith('@') && worksheetNames.includes(part.slice(1))) {
        return (
          <span key={index} className={styles.taggedWorksheet}>
            {part}
          </span>
        );
      }
      return <span key={index}>{part}</span>;
    });
  };

  return (
    <div className={styles.userInput}>
      <div className={styles.textAreaWrapper}>
        <Textarea
          ref={textAreaRef}
          className={styles.textArea}
          value={inputText}
          onChange={handleInputChange}
          onKeyPress={handleKeyPress}
          onKeyDown={handleKeyDown}
          placeholder="Type your message here..."
          rows={1}
        />
        <div className={styles.textOverlay} aria-hidden="true">
          {renderInputText()}
        </div>
        <div className={styles.suggestion} style={{ left: `${12 + cursorPosition * 8}px`, top: "12px" }}>
          {suggestion}
        </div>
        <Button
          icon={<Send24Regular className={styles.icon} />}
          onClick={handleSend}
          className={styles.button}
          disabled={!inputText.trim()}
        />
      </div>
    </div>
  );
};

export default UserInput;
