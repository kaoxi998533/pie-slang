// Builds the ProofStateInput sent to the Tactic LLM, and the plain-text proof
// state shown to the General LLM. Keep both in this one file so the request the
// model sees cannot drift between callers.

import type { ContextEntry, ProofStateInput } from "./types";

/**
 * Proof state from already-serialized proof-editor data.
 *
 * Global context: definitions and theorems that existed before the proof
 * started, deduplicated by name. Local context: goal hypotheses that are not
 * global and are not the claim being proved. Types are expected to be sugared
 * already (proofstate.ts and startSession use sugarType).
 */
export function proofStateFromEntries(
  goalType: string,
  goalContext: readonly ContextEntry[],
  globalEntries: readonly ContextEntry[],
  claimName: string,
): ProofStateInput {
  const globalNames = new Set<string>();
  const globalContext: ContextEntry[] = [];
  for (const entry of globalEntries) {
    if (!globalNames.has(entry.name)) {
      globalNames.add(entry.name);
      globalContext.push({ name: entry.name, type: entry.type });
    }
  }
  const localContext = goalContext
    .filter((c) => !globalNames.has(c.name) && c.name !== claimName)
    .map((c) => ({ name: c.name, type: c.type }));
  return { goal: goalType, globalContext, localContext };
}

/** The proof state as text, in the same layout as the Tactic LLM's training data. */
export function formatProofStateText(state: ProofStateInput): string {
  const parts: string[] = [];
  if (state.globalContext.length > 0) {
    parts.push("Definitions:\n" + state.globalContext.map(formatEntry).join("\n"));
  }
  if (state.localContext.length > 0) {
    parts.push("Local variables:\n" + state.localContext.map(formatEntry).join("\n"));
  }
  parts.push(`Goal: ${state.goal}`);
  return parts.join("\n");
}

function formatEntry(entry: ContextEntry): string {
  return `  ${entry.name} : ${entry.type}`;
}
