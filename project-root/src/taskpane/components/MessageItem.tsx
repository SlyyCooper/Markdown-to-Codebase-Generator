import * as React from "react";
import { Text, makeStyles, tokens } from "@fluentui/react-components";

interface MessageItemProps {
  isUser: boolean;
  children: React.ReactNode;
  worksheetNames: string[];
}

const useStyles = makeStyles({
  messageContainer: {
    display: "flex",
    alignItems: "flex-start",
    marginBottom: "16px",
  },
  messageItem: {
    maxWidth: "80%",
    padding: "12px 16px",
    borderRadius: "18px",
    fontSize: "14px",
    lineHeight: "1.5",
    boxShadow: tokens.shadow4,
  },
  userMessage: {
    backgroundColor: tokens.colorBrandBackgroundStatic,
    color: tokens.colorNeutralForegroundOnBrand,
    borderBottomRightRadius: "4px",
    marginLeft: "auto",
  },
  aiMessage: {
    backgroundColor: "rgba(255, 255, 255, 0.9)",
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: "4px",
  },
  avatar: {
    width: "24px",
    height: "24px",
    marginRight: "8px",
  },
});

const MessageItem: React.FC<MessageItemProps> = ({ isUser, children, worksheetNames }) => {
  const styles = useStyles();

  const renderMessageWithTaggedSheets = (text: string) => {
    const parts = text.split(/@(\w+)/);
    return parts.map((part, index) => {
      if (index % 2 === 1 && worksheetNames.includes(part)) {
        return <span key={index} style={{ fontWeight: 'bold', color: '#0078d4' }}>@{part}</span>;
      }
      return part;
    });
  };

  return (
    <div className={styles.messageContainer}>
      {!isUser && <img src="assets/logo-filled.png" alt="AI Avatar" className={styles.avatar} />}
      <div className={`${styles.messageItem} ${isUser ? styles.userMessage : styles.aiMessage}`}>
        {typeof children === 'string' ? renderMessageWithTaggedSheets(children) : children}
      </div>
    </div>
  );
};

export default MessageItem;
