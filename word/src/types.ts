export type Role = "system" | "user" | "assistant" | "tool";

export interface ChatMessage {
  role: Role;
  content: string | null;
  tool_call_id?: string;
  name?: string;
  tool_calls?: ToolCall[];
}

export interface ToolCall {
  id: string;
  type: "function";
  function: {
    name: string;
    arguments: string;
  };
}

export interface OpenAITool {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface ActivityItem {
  id: string;
  status: "running" | "done" | "error";
  label: string;
  detail?: string;
}

export interface ToolCallTiming {
  name: string;
  ms: number;
  status: "ok" | "error";
}

export interface AgentRunMetrics {
  snapshotMs: number;
  modelRequestMs: number[];
  toolCallTimings: ToolCallTiming[];
  totalMs: number;
  iterations: number;
  taskMode: string;
}
