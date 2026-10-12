// Shared types for AI model calls. Pure types: safe to import from the
// interpreter, the web worker and the main thread.

/** One `name : type` line of a proof state. */
export interface ContextEntry {
  name: string;
  type: string;
}

/**
 * A proof state as the Tactic LLM sees it. This is exactly the request body of
 * `POST /predict` on training/serve.py, training/runpod/main.py and proxy/;
 * the server turns it into the training-format prompt, so clients never build
 * Tactic LLM prompts themselves.
 */
export interface ProofStateInput {
  goal: string;
  globalContext: ContextEntry[];
  localContext: ContextEntry[];
}

/** A predicted next tactic, e.g. `{ tactic: "elim-Nat n", tacticHead: "elim-Nat" }`. */
export interface TacticPrediction {
  tactic: string;
  tacticHead?: string;
  category?: string;
}

/** A tactic that the kernel rejected, with the kernel's message. */
export interface FailedAttempt {
  tactic: string;
  error: string;
}

/**
 * Anything that proposes the next tactic for a proof state: the fine-tuned
 * Tactic LLM, an OpenAI-compatible chat model, or a test double.
 *
 * `failed` lists earlier rejected attempts on the same goal. Predictors that
 * cannot use it (the Tactic LLM decodes greedily and was not trained on error
 * feedback) ignore it.
 */
export interface TacticPredictor {
  predict(
    state: ProofStateInput,
    failed?: readonly FailedAttempt[],
    signal?: AbortSignal,
  ): Promise<TacticPrediction>;
}
