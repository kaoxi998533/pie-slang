/**
 * Tactic Agent — autonomous proof solver using a fine-tuned LLM.
 *
 * The agent loop:
 * 1. Start with global context + goal type
 * 2. Serialize context + goal (the Tactic LLM's training data format)
 * 3. Ask the predictor for the next tactic
 * 4. Parse the tactic string and apply it via ProofManager
 * 5. If success → check if proof complete; if not, get next goal → go to 2
 * 6. If type error → record it as a failed attempt → retry from 3
 *    (chat predictors read failed attempts; the greedy Tactic LLM ignores them)
 * 7. Repeat until proof complete or max retries exhausted
 */

import { ProofManager } from '../tactics/proof-manager';
import { Goal } from '../tactics/proofstate';
import { Tactic } from '../tactics/tactics';
import { Context } from '../utils/context';
import { Location, Syntax } from '../utils/locations';
import { Position } from '../../scheme-parser/transpiler/types/location';
import { go, stop } from '../types/utils';
import { serializeContext, serializeGoal } from '../tactics/training-data-extractor';
import { Parser, schemeParse } from '../parser/parser';
import type { FailedAttempt, ProofStateInput, TacticPredictor } from '../ai/types';

export type { TacticPredictor } from '../ai/types';

// ---------------------------------------------------------------------------
// Tactic string parser — converts LLM output (toString format) to S-expression
// format that Parser.parseToTactics can handle.
// ---------------------------------------------------------------------------

/**
 * Parse a tactic string (as produced by Tactic.toString() / training data)
 * into a Tactic object. The toString() format matches the parser format
 * directly, so we just wrap in parens and parse.
 *
 * Examples:
 *   "intro n"           → (intro n)
 *   "exact (same 0)"    → (exact (same 0))
 *   "elim-Nat n"        → (elim-Nat n)
 *   "symm"              → (symm)
 *   "cong ih (+ 1)"     → (cong ih (+ 1))
 *   "exists 5 k"        → (exists 5 k)
 */
export function parseTacticString(tacticStr: string): Tactic {
  const trimmed = tacticStr.trim();

  // Extract keyword to decide wrapping
  const firstSpace = trimmed.indexOf(' ');
  const rest = firstSpace === -1 ? '' : trimmed.substring(firstSpace);

  // Wrap as S-expression: (keyword args...)
  const sexpr = rest ? `(${trimmed})` : `(${trimmed})`;

  // Parse through the scheme parser + tactic parser
  const elements = schemeParse(sexpr);

  return Parser.parseToTactics(elements[0]);
}

// ---------------------------------------------------------------------------
// Proof state for the predictor — the Tactic LLM's training-data format
// ---------------------------------------------------------------------------

function proofStateFromGoal(goal: Goal, originalContext: Context): ProofStateInput {
  const { globalContext, localContext } = serializeContext(goal.context, originalContext);
  const goalStr = serializeGoal(goal) ?? goal.type.readBackType(goal.context).prettyPrint();
  return { goal: goalStr, globalContext, localContext };
}

// ---------------------------------------------------------------------------
// Agent loop
// ---------------------------------------------------------------------------

export interface AgentResult {
  success: boolean;
  /** List of tactics applied (in order) */
  tactics: string[];
  /** Number of LLM calls made */
  llmCalls: number;
  /** Total retries due to errors */
  retries: number;
  /** Error message if failed */
  error?: string;
  /** Step-by-step log */
  log: string[];
}

export interface AgentOptions {
  /** Max retries per goal before giving up (default: 5) */
  maxRetriesPerGoal?: number;
  /** Max total LLM calls across the whole proof (default: 50) */
  maxTotalCalls?: number;
  /** Log each step to console (default: true) */
  verbose?: boolean;
  /** Stops the agent between steps and cancels the in-flight prediction */
  signal?: AbortSignal;
}

/**
 * Run the tactic agent to automatically prove a claim.
 *
 * @param context - The global context (definitions, claims, etc.)
 * @param claimName - The name of the claim to prove
 * @param predictor - The LLM backend to use for tactic prediction
 * @param options - Agent configuration
 */
export async function runTacticAgent(
  context: Context,
  claimName: string,
  predictor: TacticPredictor,
  options: AgentOptions = {},
): Promise<AgentResult> {
  const {
    maxRetriesPerGoal = 5,
    maxTotalCalls = 50,
    verbose = true,
    signal,
  } = options;

  const result: AgentResult = {
    success: false,
    tactics: [],
    llmCalls: 0,
    retries: 0,
    log: [],
  };

  function log(msg: string) {
    result.log.push(msg);
    if (verbose) console.log(`[TacticAgent] ${msg}`);
  }

  // Initialize proof
  const pm = new ProofManager();
  const startResult = pm.startProof(claimName, context, dummyLocation());
  if (startResult instanceof stop) {
    result.error = startResult.message.toString();
    log(`Failed to start proof: ${result.error}`);
    return result;
  }
  log(`Started proof of ${claimName}`);

  // Store original context for distinguishing global/local
  const originalContext = new Map(context);

  // Main loop
  while (!pm.currentState!.isComplete()) {
    if (result.llmCalls >= maxTotalCalls) {
      result.error = `Max total LLM calls (${maxTotalCalls}) exceeded`;
      log(result.error);
      return result;
    }

    const currentGoal = (pm.currentState!.getCurrentGoal() as go<Goal>).result;
    log(`\nGoal: ${serializeGoal(currentGoal)}`);

    const state = proofStateFromGoal(currentGoal, originalContext);
    const failed: FailedAttempt[] = [];
    let goalSolved = false;

    for (let retry = 0; retry <= maxRetriesPerGoal; retry++) {
      if (result.llmCalls >= maxTotalCalls) {
        result.error = `Max total LLM calls (${maxTotalCalls}) exceeded`;
        log(result.error);
        return result;
      }

      if (signal?.aborted) {
        result.error = 'Aborted';
        log(result.error);
        return result;
      }
      result.llmCalls++;

      let tacticStr: string;
      try {
        tacticStr = (await predictor.predict(state, failed, signal)).tactic.trim();
      } catch (e: unknown) {
        if (signal?.aborted) continue;
        const msg = e instanceof Error ? e.message : String(e);
        log(`LLM call failed: ${msg}`);
        result.retries++;
        continue;
      }
      log(`LLM predicted: ${tacticStr}`);

      // Parse the tactic string
      let tactic: Tactic;
      try {
        tactic = parseTacticString(tacticStr);
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : String(e);
        log(`Parse error: ${msg}`);
        failed.push({ tactic: tacticStr, error: `parse error: ${msg}` });
        result.retries++;
        continue;
      }

      // Apply the tactic
      const applyResult = pm.applyTactic(tactic);
      if (applyResult instanceof stop) {
        const errMsg = applyResult.message.toString();
        log(`Tactic failed: ${errMsg}`);
        failed.push({ tactic: tacticStr, error: errMsg });
        result.retries++;
        continue;
      }

      // Success — tactic applied
      result.tactics.push(tacticStr);
      log(`Applied: ${tacticStr}`);
      goalSolved = true;
      break;
    }

    if (!goalSolved) {
      result.error = signal?.aborted
        ? 'Aborted'
        : `Failed to solve goal after ${maxRetriesPerGoal} retries: ${serializeGoal(currentGoal)}`;
      log(result.error);
      return result;
    }
  }

  result.success = true;
  log(`\nProof complete! Applied ${result.tactics.length} tactics in ${result.llmCalls} LLM calls.`);
  return result;
}

function dummyLocation(): Location {
  const pos = new Position(1, 0);
  const syntax = new Syntax(pos, pos, '');
  return new Location(syntax, false);
}
