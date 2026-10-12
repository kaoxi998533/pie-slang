// Tactic predictor backed by a General LLM (any OpenAI-compatible chat model).
// Unlike the Tactic LLM it reads the kernel errors of earlier attempts.

import { chatComplete, type ChatAccess, type ChatModelConfig } from "./openai-compatible";
import type { ContextEntry, FailedAttempt, ProofStateInput, TacticPrediction, TacticPredictor } from "../types";

export class GeneralLlmPredictor implements TacticPredictor {
  constructor(
    private readonly config: ChatModelConfig,
    private readonly access: ChatAccess,
  ) {}

  async predict(
    state: ProofStateInput,
    failed: readonly FailedAttempt[] = [],
    signal?: AbortSignal,
  ): Promise<TacticPrediction> {
    const text = await chatComplete(this.config, this.access, buildTacticPrompt(state, failed), {
      operation: "prediction",
      signal,
    });
    return { tactic: extractTactic(text) };
  }
}

export function buildTacticPrompt(state: ProofStateInput, failed: readonly FailedAttempt[] = []): string {
  let prompt = `Predict the next tactic for the following proof state.

Global context:
${formatEntries(state.globalContext)}

Local context:
${formatEntries(state.localContext)}

Goal: ${state.goal}

Respond with ONLY the tactic (e.g., "intro n", "exact (same 0)", "elim-Nat n", "symm").`;

  if (failed.length > 0) {
    prompt += `\n\nPrevious attempts that failed:\n${failed
      .map((f, i) => `  ${i + 1}. "${f.tactic}" → ${f.error}`)
      .join("\n")}\nPlease try a different tactic.`;
  }
  return prompt;
}

/** Strip code fences and surrounding quotes; keep the first non-empty line. */
export function extractTactic(text: string): string {
  const unfenced = text.trim().replace(/^```[\w]*\n?/, "").replace(/\n?```$/, "").trim();
  const firstLine = unfenced.split("\n").find((line) => line.trim()) ?? "";
  return firstLine.trim().replace(/^["'`]|["'`]$/g, "").trim();
}

function formatEntries(entries: readonly ContextEntry[]): string {
  return entries.map((e) => `  ${e.name} : ${e.type}`).join("\n") || "  (empty)";
}
