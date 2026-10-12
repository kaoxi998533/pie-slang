import 'jest';
import { jest } from '@jest/globals';
import { proofStateFromEntries, formatProofStateText } from '../ai/proof-state-input';
import { TacticLlmPredictor } from '../ai/providers/tactic-llm';
import { GeneralLlmPredictor, buildTacticPrompt, extractTactic } from '../ai/providers/general-llm-predictor';
import { chatComplete, type ChatModelConfig } from '../ai/providers/openai-compatible';
import type { ProofStateInput } from '../ai/types';

const state: ProofStateInput = {
  goal: '(= Nat (+ n 0) n)',
  globalContext: [{ name: '+', type: '(→ Nat Nat Nat)' }],
  localContext: [{ name: 'n', type: 'Nat' }],
};

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });

function mockFetch(...responses: Array<Response | ((url: string, init?: RequestInit) => Promise<Response>)>) {
  const fn = jest.fn(async (url: string, init?: RequestInit) => {
    const next = responses.shift();
    if (!next) throw new Error('unexpected fetch');
    return typeof next === 'function' ? next(url, init) : next;
  });
  (globalThis as { fetch: unknown }).fetch = fn;
  return fn;
}

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

describe('proofStateFromEntries', () => {
  it('dedupes globals and keeps only non-global, non-claim hypotheses as locals', () => {
    const result = proofStateFromEntries(
      '(= Nat n n)',
      [
        { name: '+', type: '(→ Nat Nat Nat)' },
        { name: 'thm', type: 'U' },
        { name: 'n', type: 'Nat' },
      ],
      [
        { name: '+', type: '(→ Nat Nat Nat)' },
        { name: '+', type: 'duplicate' },
      ],
      'thm',
    );
    expect(result).toEqual({
      goal: '(= Nat n n)',
      globalContext: [{ name: '+', type: '(→ Nat Nat Nat)' }],
      localContext: [{ name: 'n', type: 'Nat' }],
    });
  });

  it('formats the training-data text layout', () => {
    expect(formatProofStateText(state)).toBe(
      'Definitions:\n  + : (→ Nat Nat Nat)\nLocal variables:\n  n : Nat\nGoal: (= Nat (+ n 0) n)',
    );
    expect(formatProofStateText({ goal: 'Nat', globalContext: [], localContext: [] })).toBe('Goal: Nat');
  });
});

describe('TacticLlmPredictor', () => {
  it('posts the proof state to {url}/predict for serve.py and the proxy', async () => {
    const fetchMock = mockFetch(json({ tactic: ' elim-Nat n ', tactic_head: 'elim-Nat', category: 'elimination' }));
    const prediction = await new TacticLlmPredictor({ url: 'http://localhost:8000/' }).predict(state);
    expect(prediction).toEqual({ tactic: 'elim-Nat n', tacticHead: 'elim-Nat', category: 'elimination' });
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe('http://localhost:8000/predict');
    expect(JSON.parse(init?.body as string)).toEqual(state);
  });

  it('wraps the state for Runpod and polls until the job completes', async () => {
    const fetchMock = mockFetch(
      json({ id: 'job-1', status: 'IN_PROGRESS' }),
      json({ id: 'job-1', status: 'IN_QUEUE' }),
      json({ id: 'job-1', status: 'COMPLETED', output: { tactic: 'intro n' } }),
    );
    const predictor = new TacticLlmPredictor({
      url: 'https://api.runpod.ai/v2/abc/runsync',
      apiKey: 'rpa_test',
      pollIntervalMs: 1,
    });
    expect((await predictor.predict(state)).tactic).toBe('intro n');
    const urls = fetchMock.mock.calls.map(([url]) => url);
    expect(urls).toEqual([
      'https://api.runpod.ai/v2/abc/runsync',
      'https://api.runpod.ai/v2/abc/status/job-1',
      'https://api.runpod.ai/v2/abc/status/job-1',
    ]);
    const [, init] = fetchMock.mock.calls[0];
    expect(JSON.parse(init?.body as string)).toEqual({ input: state });
    expect((init?.headers as Record<string, string>).Authorization).toBe('Bearer rpa_test');
  });

  it('rejects failed Runpod jobs, HTTP errors and missing tactics', async () => {
    mockFetch(json({ id: 'job-1', status: 'FAILED' }));
    await expect(new TacticLlmPredictor({ url: 'https://api.runpod.ai/v2/abc' }).predict(state))
      .rejects.toThrow('did not finish');
    mockFetch(json({}, 502));
    await expect(new TacticLlmPredictor({ url: 'http://localhost:8000' }).predict(state))
      .rejects.toThrow('HTTP 502');
    mockFetch(json({ raw_output: '' }));
    await expect(new TacticLlmPredictor({ url: 'http://localhost:8000' }).predict(state))
      .rejects.toThrow('no tactic');
  });

  it('times out a stalled server', async () => {
    mockFetch((_url, init) => new Promise((_resolve, reject) => {
      init?.signal?.addEventListener('abort', () => reject(new Error('aborted')));
    }));
    await expect(new TacticLlmPredictor({ url: 'http://localhost:8000', timeoutMs: 5 }).predict(state))
      .rejects.toThrow('timed out');
  });

  it('stops polling when the caller aborts', async () => {
    const controller = new AbortController();
    mockFetch(async () => {
      controller.abort();
      return json({ id: 'job-1', status: 'IN_PROGRESS' });
    });
    await expect(new TacticLlmPredictor({ url: 'https://api.runpod.ai/v2/abc', pollIntervalMs: 10_000 })
      .predict(state, [], controller.signal)).rejects.toThrow('Aborted');
  });
});

const chatConfig: ChatModelConfig = {
  displayName: 'TestLLM',
  logProvider: 'deepseek',
  endpoint: 'https://llm.example/chat/completions',
  model: 'test-model',
  maxTokens: 64,
};
const answer = (content: string) => json({ choices: [{ finish_reason: 'stop', message: { content } }] });

describe('chatComplete', () => {
  it('uses the proxy /chat route when no key is given', async () => {
    const fetchMock = mockFetch(answer('ok'));
    expect(await chatComplete(chatConfig, { proxyUrl: 'https://proxy.example/' }, 'hi', { operation: 'hint' }))
      .toBe('ok');
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe('https://proxy.example/chat');
    expect(JSON.parse(init?.body as string)).toEqual({ prompt: 'hi', operation: 'hint' });
  });

  it('rejects with AbortError when the caller aborts, without retrying', async () => {
    const controller = new AbortController();
    const fetchMock = mockFetch((_url, init) => new Promise((_resolve, reject) => {
      init?.signal?.addEventListener('abort', () => reject(new Error('aborted')));
      controller.abort();
    }));
    const error = await chatComplete(chatConfig, { apiKey: 'k' }, 'hi', {
      operation: 'hint',
      signal: controller.signal,
    }).catch((e) => e as Error);
    expect((error as Error).name).toBe('AbortError');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});

describe('GeneralLlmPredictor', () => {
  it('includes failed attempts in the prompt and extracts a bare tactic', async () => {
    const fetchMock = mockFetch(answer('```\nexact (same n)\n```'));
    const predictor = new GeneralLlmPredictor(chatConfig, { apiKey: 'k' });
    const prediction = await predictor.predict(state, [{ tactic: 'intro x', error: 'Not a Π type' }]);
    expect(prediction).toEqual({ tactic: 'exact (same n)' });
    const prompt = JSON.parse(fetchMock.mock.calls[0][1]?.body as string).messages[0].content as string;
    expect(prompt).toContain('Goal: (= Nat (+ n 0) n)');
    expect(prompt).toContain('"intro x" → Not a Π type');
  });

  it('builds a prompt without a failure section when there are no failures', () => {
    expect(buildTacticPrompt(state)).not.toContain('Previous attempts');
  });

  it('extracts the first line without quotes', () => {
    expect(extractTactic('"intro n"\nbecause...')).toBe('intro n');
  });
});
