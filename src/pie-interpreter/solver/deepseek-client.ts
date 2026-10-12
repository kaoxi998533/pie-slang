import type { AiCallMeta } from "./ai-call-log";
import { chatComplete, hasChatAccess, type ChatAccess, type ChatModelConfig } from "../ai/providers/openai-compatible";
/** Shared by browser goal translation and worker hint explanations. */
export const DEEPSEEK_MODEL = "deepseek-flash";

/** The General LLM: DeepSeek over its OpenAI-compatible chat API. */
export const DEEPSEEK_CONFIG: ChatModelConfig = {
  displayName: "DeepSeek",
  logProvider: "deepseek",
  endpoint: "https://api.deepseek.com/chat/completions",
  model: DEEPSEEK_MODEL,
  maxTokens: 2048,
  extraBody: { thinking: { type: "disabled" } },
};

/**
 * How to reach the General LLM: the user's own DeepSeek key (a plain string,
 * or `apiKey`), or a proxy that holds the key server-side (`proxyUrl`). A
 * user's own key takes priority over the proxy.
 */
export type GeneralLlmAccess = string | ChatAccess;

function toChatAccess(access: GeneralLlmAccess | undefined): ChatAccess {
  return typeof access === "string" ? { apiKey: access } : access ?? {};
}

export function hasGeneralLlm(access: GeneralLlmAccess | undefined): boolean {
  return hasChatAccess(toChatAccess(access));
}

export function callDeepSeek(
  access: GeneralLlmAccess,
  prompt: string,
  operation: AiCallMeta['operation'] = 'hint',
  signal?: AbortSignal,
): Promise<string> {
  return chatComplete(DEEPSEEK_CONFIG, toChatAccess(access), prompt, { operation, signal });
}
