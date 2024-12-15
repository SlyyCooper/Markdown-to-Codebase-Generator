import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import MessageItem from "./MessageItem";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

interface MessageListProps {
  messages: Array<{ text: string; isUser: boolean }>;
  worksheetNames: string[];
}

const useStyles = makeStyles({
  messageList: {
    height: "100%",
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    padding: "10px",
  },
  messagesContainer: {
    marginTop: "0",
  },
  markdownContent: {
    "& p": { margin: "0 0 10px 0" },
    "& ul, & ol": { paddingLeft: "20px" },
    "& pre": {
      backgroundColor: "#f0f0f0",
      padding: "10px",
      borderRadius: "4px",
      overflowX: "auto",
    },
    "& code": {
      backgroundColor: "#f0f0f0",
      padding: "2px 4px",
      borderRadius: "2px",
    },
    "& table": {
      borderCollapse: "collapse",
      margin: "15px 0",
      fontSize: "0.9em",
      fontFamily: "sans-serif",
      minWidth: "400px",
      boxShadow: "0 0 20px rgba(0, 0, 0, 0.15)",
    },
    "& thead tr": {
      backgroundColor: "#009879",
      color: "#ffffff",
      textAlign: "left",
    },
    "& th, & td": {
      padding: "12px 15px",
    },
    "& tbody tr": {
      borderBottom: "1px solid #dddddd",
    },
    "& tbody tr:nth-of-type(even)": {
      backgroundColor: "#f3f3f3",
    },
    "& tbody tr:last-of-type": {
      borderBottom: "2px solid #009879",
    },
  },
});

const MessageList: React.FC<MessageListProps> = ({ messages, worksheetNames }) => {
  const styles = useStyles();
  const listRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (listRef.current) {
      listRef.current.scrollTop = listRef.current.scrollHeight;
    }
  }, [messages]);

  return (
    <div className={styles.messageList} ref={listRef}>
      <div className={styles.messagesContainer}>
        {messages.map((message, index) => (
          <MessageItem key={index} isUser={message.isUser} worksheetNames={worksheetNames}>
            {message.isUser ? (
              message.text
            ) : (
              <ReactMarkdown className={styles.markdownContent} remarkPlugins={[remarkGfm]}>
                {message.text}
              </ReactMarkdown>
            )}
          </MessageItem>
        ))}
      </div>
    </div>
  );
};

export default MessageList;
