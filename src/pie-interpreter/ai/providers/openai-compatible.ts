// General LLM client for any OpenAI-compatible /chat/completions endpoint
// (DeepSeek today). Either calls the provider with the user's own key, or the
// proxy/ Worker's /chat route, which adds the key, model and limits itself and
// answers with the same `choices` shape.

import { loggedAiFetch, type AiCallMeta } from "../../solver/ai-call-log";
import { abortableDelay, withTimeout } from "../abort";

export interface ChatModelConfig {
  /** Name used in user-facing error messages, e.g. "DeepSeek". */
  displayName: string;
  /** Provider tag for the local AI call log. */
  logProvider: AiCallMeta["provider"];
  /** Full chat-completions URL, e.g. https://api.deepseek.com/chat/completions. */
  endpoint: string;
  model: string;
  maxTokens: number;
  /** Provider-specific request fields, merged into the body. */
  extraBody?: Record<string, unknown>;
  /** Per-attempt timeout. */
  timeoutMs?: number;
}

/** The user's own key, or a proxy that holds the key server-side. A key wins over the proxy. */
export interface ChatAccess {
  apiKey?: string | null;
  proxyUrl?: string | null;
}

export interface ChatOptions {
  operation: AiCallMeta["operation"];
  signal?: AbortSignal;
}

const DEFAULT_TIMEOUT_MS = 60_000;
const MAX_RETRIES = 2;

export function resolveChatAccess(access: ChatAccess | undefined): { key: string; proxyUrl: string } {
  return {
    key: (access?.apiKey ?? "").trim(),
    proxyUrl: (access?.proxyUrl ?? "").trim().replace(/\/+$/, ""),
  };
}

export function hasChatAccess(access: ChatAccess | undefined): boolean {
  const { key, proxyUrl } = resolveChatAccess(access);
  return !!(key || proxyUrl);
}

/**
 * Send one user prompt and return the trimmed answer. Retries transient 503s
 * with backoff. Error messages never include response bodies, which may echo
 * request data.
 */
export async function chatComplete(
  config: ChatModelConfig,
  access: ChatAccess,
  prompt: string,
  options: ChatOptions,
): Promise<string> {
  const { key, proxyUrl } = resolveChatAccess(access);
  const name = config.displayName;
  if (!key && !proxyUrl) throw new Error(`${name} API key is not set.`);

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const deadline = withTimeout(options.signal, config.timeoutMs ?? DEFAULT_TIMEOUT_MS);
    let response: Response;
    let payload: unknown;
    try {
      response = key
        ? await loggedAiFetch({ provider: config.logProvider, operation: options.operation }, config.endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` },
          body: JSON.stringify({
            model: config.model,
            messages: [{ role: "user", content: prompt }],
            stream: false,
            ...config.extraBody,
            max_tokens: config.maxTokens,
          }),
          signal: deadline.signal,
        })
        : await loggedAiFetch({ provider: config.logProvider, operation: options.operation }, `${proxyUrl}/chat`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ prompt, operation: options.operation }),
          signal: deadline.signal,
        });
      if (response.ok) payload = await response.json();
    } catch {
      if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");
      throw new Error(deadline.timedOut()
        ? `${name} request timed out. Please try again.`
        : `${name} request failed. Check your connection and try again.`);
    } finally {
      deadline.dispose();
    }

    if (response.status === 503 && attempt < MAX_RETRIES) {
      await abortableDelay(2000 * Math.pow(2, attempt), options.signal);
      continue;
    }
    if (!response.ok) {
      const advice = response.status === 401 ? ` Check your ${name} API key.`
        : response.status === 402 ? ` Check your ${name} account balance.`
        : response.status === 429 ? " Rate limit reached; try again later." : "";
      throw new Error(`${name} API returned HTTP ${response.status}.${advice}`);
    }

    const result = payload as {
      choices?: Array<{ finish_reason?: string; message?: { content?: unknown } }>;
    } | null;
    const choice = result?.choices?.[0];
    if (choice?.finish_reason && choice.finish_reason !== "stop") {
      throw new Error(`${name} did not return a complete answer. Please try again.`);
    }
    const content = choice?.message?.content;
    if (typeof content !== "string" || !content.trim()) {
      throw new Error(`${name} returned an empty response.`);
    }
    return content.trim();
  }
  throw new Error(`${name} is temporarily unavailable.`);
}
