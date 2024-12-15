/* global Office, console */
import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import MessageList from "./MessageList";
import UserInput from "./UserInput";
import OpenAI from "openai";
import { initializeOpenAI } from "../../embedding_operations";
import { getActiveWorksheetName, analyzeData, getRangeData, writeToCellOperation, readFromCellOperation, formatCellOperation, addChartOperation, writeToSelectedRange, addPivotTableOperation, manageWorksheet, filterDataOperation, sortDataOperation, mergeCellsOperation, unmergeCellsOperation, autofitColumnsOperation, autofitRowsOperation, applyConditionalFormat, clearConditionalFormats, getWorksheetNames, embedWorksheet, embedAllWorksheets } from "../../excelOperations";
import { tools } from "../../tools/tools";

const DEFAULT_API_KEY = "YOUR_API_KEY";

const useStyles = makeStyles({
  chatInterface: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    backgroundColor: "rgba(255, 255, 255, 0.8)",
    opacity: 0,
    transition: "opacity 1s ease-in",
  },
  fadeIn: {
    opacity: 1,
  },
  messageListContainer: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "20px",
  },
});

interface ChatInterfaceProps {
  apiKey: string;
  setIsLoading: React.Dispatch<React.SetStateAction<boolean>>;
}

const ChatInterface: React.FC<ChatInterfaceProps> = ({ apiKey, setIsLoading }) => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<Array<{ text: string; isUser: boolean }>>([]);
  const [worksheetNames, setWorksheetNamesState] = React.useState<string[]>([]);
  const [activeWorksheet, setActiveWorksheet] = React.useState<string>("");

  const effectiveApiKey = apiKey || DEFAULT_API_KEY;

  const openaiRef = React.useRef<OpenAI | null>(null);
  React.useEffect(() => {
    openaiRef.current = new OpenAI({ apiKey: effectiveApiKey, dangerouslyAllowBrowser: true });
  }, [effectiveApiKey]);

  const openai = React.useMemo(() => initializeOpenAI(effectiveApiKey), [effectiveApiKey]);

  const handleSendMessage = async (text: string, taggedSheets: string[]) => {
    setIsLoading(true);
    const currentActiveWorksheet = await getActiveWorksheetName();
    setMessages((prevMessages) => [
      ...prevMessages,
      { text: `[Active Worksheet: ${currentActiveWorksheet}] ${text}`, isUser: true },
    ]);

    if (!effectiveApiKey) {
      setMessages((prevMessages) => [...prevMessages, { text: "Please set your API key in the settings.", isUser: false }]);
      setIsLoading(false);
      return;
    }

    try {
      if (taggedSheets.includes("workbook")) {
        await handleEmbedAllWorksheets();
      } else {
        for (const sheetName of taggedSheets) {
          await handleEmbedding(sheetName);
        }
      }

      const { address, rowCount, columnCount } = await (async () => {
        try {
          return await getSelectedRangeInfo();
        } catch {
          return { address: "A1", rowCount: 1, columnCount: 1 };
        }
      })();

      const completion = await openaiRef.current?.chat.completions.create({
        messages: [
          {
            role: "system",
            content: `You are a helpful assistant that interacts with Excel...`,
          },
          ...messages.map((msg) => ({
            role: msg.isUser ? "user" as const : "assistant" as const,
            content: msg.text,
          })),
          { role: "user", content: text },
        ],
        model: "gpt-4o-2024-08-06",
        tools: tools,
        tool_choice: "auto",
      });

      if (completion?.choices[0]?.finish_reason === "tool_calls") {
        const toolCalls = completion.choices[0].message.tool_calls;
        const toolResults = await Promise.all(
          toolCalls.map(async (toolCall) => {
            const args = JSON.parse(toolCall.function.arguments);
            let functionResult = "";

            try {
              switch (toolCall.function.name) {
                case "write_to_excel":
                  await writeToCellOperation(args.startCell, args.values);
                  functionResult = `Values written to ${args.startCell}`;
                  break;
                case "read_from_excel":
                  const cellValue = await readFromCellOperation(args.cellAddress);
                  functionResult = `The value in ${args.cellAddress} is "${cellValue}"`;
                  break;
                case "format_cell":
                  await formatCellOperation(args.cellAddress, {
                    fontColor: args.fontColor,
                    backgroundColor: args.backgroundColor,
                    bold: args.bold,
                  });
                  functionResult = `Formatted cell ${args.cellAddress}`;
                  break;
                case "add_chart":
                  await addChartOperation(args.dataRange, args.chartType);
                  functionResult = `Added ${args.chartType} chart from ${args.dataRange}`;
                  break;
                case "analyze_selected_range":
                  const { values } = await getRangeData();
                  functionResult = `Analysis:\n${await analyzeData(values, args.analysisType)}`;
                  break;
                case "write_to_selected_range":
                  const resultAddress = await writeToSelectedRange(args.values);
                  functionResult = `Wrote data to selected range: ${resultAddress}`;
                  break;
                case "read_range":
                  const readRes = await getRangeData(args.rangeAddress);
                  functionResult = `Values in ${readRes.address}:\n${JSON.stringify(readRes.values)}`;
                  break;
                case "add_pivot":
                  await addPivotTableOperation(
                    args.sourceDataRange,
                    args.destinationCell,
                    args.rowFields,
                    args.columnFields,
                    args.dataFields,
                    args.filterFields
                  );
                  functionResult = `Added pivot table at ${args.destinationCell}`;
                  break;
                case "manage_worksheet":
                  functionResult = await manageWorksheet(args.action, args.sheetName);
                  break;
                case "filter_data":
                  const filterResult = await filterDataOperation(args.range, args.column, args.filterType, args.criteria);
                  functionResult = `Filtered ${filterResult.range}, ${filterResult.filteredCount} rows match.`;
                  break;
                case "sort_data":
                  functionResult = await sortDataOperation(args.range, args.sortFields, args.matchCase, args.hasHeaders);
                  break;
                case "merge_cells":
                  functionResult = await mergeCellsOperation(args.range, args.across);
                  break;
                case "unmerge_cells":
                  functionResult = await unmergeCellsOperation(args.range);
                  break;
                case "autofit_columns":
                  functionResult = await autofitColumnsOperation(args.range);
                  break;
                case "autofit_rows":
                  functionResult = await autofitRowsOperation(args.range);
                  break;
                case "apply_conditional_format":
                  functionResult = await applyConditionalFormat(args.range, args.formatType, args.rule, args.format);
                  break;
                case "clear_conditional_formats":
                  functionResult = await clearConditionalFormats(args.range);
                  break;
                case "get_worksheet_names":
                  const names = await getWorksheetNames();
                  functionResult = `Worksheets: ${names.join(", ")}`;
                  break;
                case "get_active_worksheet_name":
                  const activeName = await getActiveWorksheetName();
                  functionResult = `Active worksheet: ${activeName}`;
                  break;
                default:
                  functionResult = "Unrecognized function call.";
              }
            } catch (error: any) {
              functionResult = `Error: ${error.message || String(error)}`;
            }

            return {
              tool_call_id: toolCall.id,
              role: "tool" as const,
              content: functionResult,
            };
          })
        );

        const completionPayload = {
          model: "gpt-4o-2024-08-06",
          messages: [
            {
              role: "system",
              content: `You are a helpful assistant that interacts with Excel...`,
            },
            ...messages.map((msg) => ({
              role: msg.isUser ? "user" : "assistant",
              content: msg.text,
            })),
            { role: "user", content: text },
            completion.choices[0].message,
            ...toolResults,
          ],
          tools: tools,
          tool_choice: "auto",
        };

        const finalResponse = await openaiRef.current?.chat.completions.create(completionPayload);

        if (finalResponse?.choices[0]?.message?.content) {
          const aiResponse = finalResponse.choices[0].message.content;
          setMessages((prevMessages) => [...prevMessages, { text: aiResponse, isUser: false }]);
        }
      } else if (completion?.choices[0]?.message?.content) {
        const aiResponse = completion.choices[0].message.content;
        setMessages((prevMessages) => [...prevMessages, { text: aiResponse, isUser: false }]);
      }
    } catch (error: any) {
      console.error("Error calling OpenAI API:", error);
      let errorMessage = "Sorry, I encountered an error. Please try again.";

      if (error instanceof Error) {
        errorMessage = `Error: ${error.message}`;
      } else if (error && error.response && error.response.status === 401) {
        errorMessage = "Invalid API key. Please check your settings.";
      }

      setMessages((prevMessages) => [...prevMessages, { text: errorMessage, isUser: false }]);
    } finally {
      setIsLoading(false);
    }
  };

  const [fadeIn, setFadeIn] = React.useState(false);

  React.useEffect(() => {
    setTimeout(() => setFadeIn(true), 100);
  }, []);

  React.useEffect(() => {
    const fetchWorksheetNames = async () => {
      try {
        const names = await getWorksheetNames();
        setWorksheetNamesState(names);
      } catch (error) {
        console.error("Error fetching worksheet names:", error);
      }
    };

    fetchWorksheetNames();
  }, []);

  React.useEffect(() => {
    const updateActiveWorksheet = async () => {
      try {
        const name = await getActiveWorksheetName();
        setActiveWorksheet(name);
      } catch (error) {
        console.error("Error fetching active worksheet:", error);
      }
    };

    updateActiveWorksheet();
    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, updateActiveWorksheet);

    return () => {
      Office.context.document.removeHandlerAsync(Office.EventType.ActiveViewChanged, updateActiveWorksheet);
    };
  }, []);

  const handleEmbedding = async (sheetName: string): Promise<number[]> => {
    try {
      const embeddingResult = await embedWorksheet(openai, sheetName);
      setMessages((prevMessages) => [
        ...prevMessages,
        { text: `Embedding for "${sheetName}" created successfully.`, isUser: false },
      ]);
      return embeddingResult;
    } catch (error) {
      console.error("Error creating embedding:", error);
      setMessages((prevMessages) => [
        ...prevMessages,
        { text: "Error creating embedding. Please try again.", isUser: false },
      ]);
      throw error;
    }
  };

  const handleEmbedAllWorksheets = async (): Promise<{ [key: string]: number[] }> => {
    try {
      const embeddingResults = await embedAllWorksheets(openai);
      setMessages((prevMessages) => [
        ...prevMessages,
        { text: "All worksheets embeddings created successfully.", isUser: false },
      ]);
      return embeddingResults;
    } catch (error) {
      console.error("Error creating embeddings for all worksheets:", error);
      setMessages((prevMessages) => [
        ...prevMessages,
        { text: "Error creating embeddings for all worksheets. Please try again.", isUser: false },
      ]);
      throw error;
    }
  };

  return (
    <div className={`${styles.chatInterface} ${fadeIn ? styles.fadeIn : ""}`}>
      <div className={styles.messageListContainer}>
        <MessageList messages={messages} worksheetNames={worksheetNames} />
      </div>
      <UserInput onSendMessage={handleSendMessage} worksheetNames={worksheetNames} />
    </div>
  );
};

export default ChatInterface;
