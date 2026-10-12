// Progressive hints for the proof editor. One sequential fallback chain, one
// model call per step:
//
//   1. Tactic LLM predicts the next tactic (once per goal, cached), and the
//      prediction is validated by parsing it.
//   2. Valid prediction → the General LLM explains it at the requested level,
//      or a template explains it when the General LLM is unavailable.
//   3. No valid prediction → the General LLM produces the hint by itself.
//   4. Otherwise → rule-based hint.

import { schemeParse, Parser } from "../parser/parser";
import { TACTIC_REQUIREMENTS, type HintLevel, type TacticType } from "../protocol";
import { hasGeneralLlm, type GeneralLlmAccess } from "../solver/deepseek-client";
import {
  explainTactic,
  generateProgressiveHint,
  generateRuleBasedHint,
  type ExplainTacticRequest,
  type ProgressiveHint,
} from "../solver/hint-generator";
import { formatProofStateText } from "./proof-state-input";
import { TacticLlmPredictor } from "./providers/tactic-llm";
import type { ContextEntry, ProofStateInput } from "./types";

export type HintSource = "lora" | "deepseek" | "rule-based";
export type SourcedHint = ProgressiveHint & { source: HintSource };

export interface HintPipelineRequest {
  /** Cache key for the Tactic LLM prediction. */
  goalId: string;
  level: HintLevel;
  previousHint?: ProgressiveHint;
  /** Every hypothesis of the goal, for General LLM prompts. */
  context: ContextEntry[];
  /** The goal split into global and local context, for the Tactic LLM. */
  proofState: ProofStateInput;
  tacticLlm?: { url?: string | null; apiKey?: string | null };
  generalLlm: GeneralLlmAccess;
}

export interface TacticLlmPrediction {
  tactic: string;
  tacticHead: string;
  category: string;
  /** The tactic string parses as a tactic. */
  validated: boolean;
}

export type PredictTactic = (
  url: string,
  state: ProofStateInput,
  apiKey?: string | null,
) => Promise<TacticLlmPrediction | null>;

const HINT_TACTICS = (Object.keys(TACTIC_REQUIREMENTS) as TacticType[]).filter((t) => t !== "todo");

export class HintPipeline {
  private readonly predictions = new Map<string, TacticLlmPrediction>();

  constructor(private readonly predictTactic: PredictTactic = predictAndValidateTactic) {}

  /** Forget cached predictions, e.g. when a new proof session starts. */
  clear(): void {
    this.predictions.clear();
  }

  async getHint(request: HintPipelineRequest): Promise<SourcedHint> {
    const prediction = await this.getPrediction(request);

    if (prediction?.validated) {
      const explainRequest: ExplainTacticRequest = {
        predictedTactic: prediction.tactic,
        tacticCategory: prediction.category,
        goalType: request.proofState.goal,
        context: request.context,
        level: request.level,
        proofStateText: formatProofStateText(request.proofState),
      };
      // explainTactic falls back to a template itself when the General LLM
      // is unavailable or fails.
      const access = hasGeneralLlm(request.generalLlm) ? request.generalLlm : "";
      return { ...(await explainTactic(access, explainRequest)), source: "lora" };
    }

    const progressiveRequest = {
      goalType: request.proofState.goal,
      context: request.context,
      availableTactics: HINT_TACTICS,
      currentLevel: request.level,
      previousHint: request.previousHint,
    };
    if (hasGeneralLlm(request.generalLlm)) {
      try {
        return { ...(await generateProgressiveHint(request.generalLlm, progressiveRequest)), source: "deepseek" };
      } catch {
        // Fall back to rules without logging provider error contents.
      }
    }
    return { ...generateRuleBasedHint(progressiveRequest), source: "rule-based" };
  }

  /** The Tactic LLM is asked at most once per goal; later levels reuse its answer. */
  private async getPrediction(request: HintPipelineRequest): Promise<TacticLlmPrediction | undefined> {
    const cached = this.predictions.get(request.goalId);
    if (cached || !request.tacticLlm?.url) return cached;
    const prediction = await this.predictTactic(request.tacticLlm.url, request.proofState, request.tacticLlm.apiKey);
    if (prediction) this.predictions.set(request.goalId, prediction);
    return prediction ?? undefined;
  }
}

/**
 * Ask the Tactic LLM for the next tactic and check that it parses. Returns null
 * when the server fails or the output is not a usable single-step tactic.
 */
export async function predictAndValidateTactic(
  url: string,
  state: ProofStateInput,
  apiKey?: string | null,
): Promise<TacticLlmPrediction | null> {
  let prediction;
  try {
    prediction = await new TacticLlmPredictor({ url, apiKey }).predict(state);
  } catch {
    return null;
  }
  const tactic = prediction.tactic;
  // Multi-line output or a long `exact` (a whole proof term) is not a useful
  // single-step hint.
  if (tactic.includes("\n") || tactic.length > 100) return null;

  return {
    tactic,
    tacticHead: prediction.tacticHead ?? "",
    category: prediction.category ?? "unknown",
    validated: parsesAsTactic(tactic),
  };
}

/**
 * Structural validation only. Type-checking the tactic against the goal would
 * need a copy of the proof state, which ProofManager does not provide.
 */
function parsesAsTactic(tactic: string): boolean {
  try {
    const wrapped = tactic.trim().startsWith("(") ? tactic : `(${tactic})`;
    Parser.parseToTactics(schemeParse(wrapped)[0]);
    return true;
  } catch {
    return false;
  }
}
