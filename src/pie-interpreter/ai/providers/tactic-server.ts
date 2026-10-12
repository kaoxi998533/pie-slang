// Tactic LLM client. One class for every place the fine-tuned model is served:
//   - training/serve.py, or the proxy/ Worker:  POST {url}/predict
//   - Runpod serverless directly (api.runpod.ai): POST {url}/runsync, then poll
//     {url}/status/{id} while a cold worker reloads the model.
// All three take a ProofStateInput and return { tactic, tactic_head, category }.

import { loggedAiFetch } from "../../solver/ai-call-log";
import { abortableDelay, withTimeout } from "../abort";
import type { FailedAttempt, ProofStateInput, TacticPrediction, TacticPredictor } from "../types";

export interface TacticServerOptions {
  /** Server base URL: a local serve.py, the proxy, or a Runpod endpoint (…/v2/<id>). */
  url: string;
  /** Bearer token; only sent to Runpod. */
  apiKey?: string | null;
  /** Upper bound for the whole prediction, including Runpod cold-start polling. */
  timeoutMs?: number;
  /** Delay between Runpod status polls. */
  pollIntervalMs?: number;
}

interface PredictOutput {
  tactic?: unknown;
  tactic_head?: unknown;
  category?: unknown;
}

// A cold 7B reload takes ~80s and runsync gives up after ~60s; 3 minutes
// covers the reload plus queueing.
const DEFAULT_TIMEOUT_MS = 180_000;
const DEFAULT_POLL_INTERVAL_MS = 2_000;

export class TacticServerPredictor implements TacticPredictor {
  private readonly url: string;

  constructor(private readonly options: TacticServerOptions) {
    this.url = options.url.trim().replace(/\/+$/, "");
  }

  /** The Tactic LLM decodes greedily and ignores `failed`: retrying the same state gives the same tactic. */
  async predict(
    state: ProofStateInput,
    _failed?: readonly FailedAttempt[],
    signal?: AbortSignal,
  ): Promise<TacticPrediction> {
    const deadline = withTimeout(signal, this.options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
    try {
      const output = this.isRunpod()
        ? await this.predictRunpod(state, deadline.signal)
        : await this.predictDirect(state, deadline.signal);
      if (typeof output.tactic !== "string" || !output.tactic.trim()) {
        throw new Error("Tactic LLM returned no tactic.");
      }
      return {
        tactic: output.tactic.trim(),
        tacticHead: typeof output.tactic_head === "string" ? output.tactic_head : undefined,
        category: typeof output.category === "string" ? output.category : undefined,
      };
    } catch (error) {
      if (deadline.timedOut()) throw new Error("Tactic LLM timed out.");
      if (signal?.aborted) throw new DOMException("Aborted", "AbortError");
      throw error;
    } finally {
      deadline.dispose();
    }
  }

  private isRunpod(): boolean {
    return this.url.includes("runpod.ai");
  }

  private async predictDirect(state: ProofStateInput, signal: AbortSignal): Promise<PredictOutput> {
    const resp = await loggedAiFetch({ provider: "lora", operation: "prediction" }, `${this.url}/predict`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(state),
      signal,
    });
    if (!resp.ok) throw new Error(`Tactic LLM returned HTTP ${resp.status}.`);
    return await resp.json() as PredictOutput;
  }

  private async predictRunpod(state: ProofStateInput, signal: AbortSignal): Promise<PredictOutput> {
    // Accept the endpoint base or a URL ending in /runsync, /run or /status.
    const base = this.url.replace(/\/(runsync|run|status)$/, "");
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (this.options.apiKey) headers["Authorization"] = `Bearer ${this.options.apiKey}`;

    const sync = await loggedAiFetch({ provider: "runpod", operation: "prediction" }, `${base}/runsync`, {
      method: "POST",
      headers,
      body: JSON.stringify({ input: state }),
      signal,
    });
    if (!sync.ok) throw new Error(`Tactic LLM returned HTTP ${sync.status}.`);

    let job = await sync.json() as { id?: string; status?: string; output?: PredictOutput };
    while (job.status !== "COMPLETED") {
      if (!job.id || ["FAILED", "CANCELLED", "TIMED_OUT"].includes(job.status ?? "")) {
        throw new Error("Tactic LLM job did not finish.");
      }
      await abortableDelay(this.options.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS, signal);
      const status = await loggedAiFetch(
        { provider: "runpod", operation: "poll" },
        `${base}/status/${encodeURIComponent(job.id)}`,
        { headers, signal },
      );
      if (!status.ok) throw new Error(`Tactic LLM returned HTTP ${status.status}.`);
      job = await status.json() as typeof job;
    }
    return job.output ?? {};
  }
}
