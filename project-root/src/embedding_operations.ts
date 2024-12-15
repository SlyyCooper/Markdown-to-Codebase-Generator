import OpenAI from "openai";
import { getCurrentSheetContent, getWorksheetNames } from "./excelOperations";

// Initialize OpenAI
export function initializeOpenAI(apiKey: string) {
  return new OpenAI({ apiKey, dangerouslyAllowBrowser: true });
}

// Get Embedding
async function getEmbedding(openai: OpenAI, text: string): Promise<number[]> {
  const response = await openai.embeddings.create({
    model: "text-embedding-3-large",
    input: text,
    encoding_format: "float",
  });
  return response.data[0].embedding;
}

export async function embedWorksheet(openai: OpenAI, sheetName: string): Promise<number[]> {
  const sheetContent = await getCurrentSheetContent({ includeMetadata: true, sheetName });
  const embedding = await getEmbedding(openai, sheetContent);
  return embedding;
}

export async function embedAllWorksheets(openai: OpenAI): Promise<{ [key: string]: number[] }> {
  const worksheetNames = await getWorksheetNames();
  const embeddings: { [key: string]: number[] } = {};

  for (const name of worksheetNames) {
    try {
      const sheetContent = await getCurrentSheetContent({ includeMetadata: true, sheetName: name });
      const embedding = await getEmbedding(openai, sheetContent);
      embeddings[name] = embedding;
    } catch (error) {
      console.error(`Error embedding worksheet ${name}:`, error);
    }
  }
  return embeddings;
}
