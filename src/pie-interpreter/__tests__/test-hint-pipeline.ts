import 'jest';
import { jest } from '@jest/globals';
import {
  HintPipeline,
  predictAndValidateTactic,
  type HintPipelineRequest,
  type PredictTactic,
  type TacticLlmPrediction,
} from '../ai/hint-pipeline';

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
const answer = (content: string) => json({ choices: [{ finish_reason: 'stop', message: { content } }] });

function mockFetch(...responses: Response[]) {
  const fn = jest.fn(async () => {
    const next = responses.shift();
    if (!next) throw new Error('unexpected fetch');
    return next;
  });
  (globalThis as { fetch: unknown }).fetch = fn;
  return fn;
}

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

const request = (overrides: Partial<HintPipelineRequest> = {}): HintPipelineRequest => ({
  goalId: 'goal_0',
  level: 'category',
  context: [{ name: 'n', type: 'Nat' }],
  proofState: {
    goal: '(Π ((n Nat)) (= Nat n n))',
    globalContext: [],
    localContext: [],
  },
  tacticLlm: { url: 'http://localhost:8000' },
  generalLlm: { apiKey: 'key' },
  ...overrides,
});

const prediction = (overrides: Partial<TacticLlmPrediction> = {}): TacticLlmPrediction => ({
  tactic: 'intro n',
  tacticHead: 'intro',
  category: 'introduction',
  validated: true,
  ...overrides,
});

function predictor(result: TacticLlmPrediction | null) {
  return jest.fn<PredictTactic>(async () => result);
}

describe('HintPipeline', () => {
  it('explains a valid Tactic LLM prediction with the General LLM, predicting once per goal', async () => {
    const predict = predictor(prediction());
    const fetchMock = mockFetch(
      answer(JSON.stringify({ explanation: 'Introduce the number first.', confidence: 0.9 })),
      answer(JSON.stringify({ explanation: 'Use intro.', confidence: 0.9 })),
    );
    const pipeline = new HintPipeline(predict);

    const first = await pipeline.getHint(request());
    expect(first).toMatchObject({ source: 'lora', explanationSource: 'deepseek', category: 'introduction' });

    const second = await pipeline.getHint(request({ level: 'tactic' }));
    expect(second).toMatchObject({ source: 'lora', tacticType: 'intro' });

    expect(predict).toHaveBeenCalledTimes(1);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('uses a template explanation when no General LLM is configured', async () => {
    const fetchMock = mockFetch();
    const hint = await new HintPipeline(predictor(prediction())).getHint(
      request({ level: 'full', generalLlm: {} }),
    );
    expect(hint).toMatchObject({
      source: 'lora',
      explanationSource: 'template',
      tacticType: 'intro',
      parameters: { variableName: 'n' },
    });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('asks the General LLM directly when there is no Tactic LLM', async () => {
    const predict = predictor(prediction());
    mockFetch(answer(JSON.stringify({ category: 'introduction', explanation: 'A Π goal needs intro.', confidence: 0.8 })));
    const hint = await new HintPipeline(predict).getHint(request({ tacticLlm: undefined }));
    expect(hint).toMatchObject({ source: 'deepseek', category: 'introduction' });
    expect(predict).not.toHaveBeenCalled();
  });

  it('skips an invalid prediction, caches it, and falls back to rules when the General LLM fails', async () => {
    const predict = predictor(prediction({ validated: false }));
    mockFetch(json({}, 401), json({}, 401));
    const pipeline = new HintPipeline(predict);

    expect((await pipeline.getHint(request())).source).toBe('rule-based');
    expect((await pipeline.getHint(request({ level: 'tactic' }))).source).toBe('rule-based');
    expect(predict).toHaveBeenCalledTimes(1);
  });

  it('does not cache a failed prediction, and clear() forgets cached ones', async () => {
    const predict = jest.fn<PredictTactic>()
      .mockResolvedValueOnce(null)
      .mockResolvedValue(prediction());
    const pipeline = new HintPipeline(predict);
    const noGeneral = request({ generalLlm: {} });

    expect((await pipeline.getHint(noGeneral)).source).toBe('rule-based');
    expect((await pipeline.getHint(noGeneral)).source).toBe('lora');
    pipeline.clear();
    await pipeline.getHint(noGeneral);
    expect(predict).toHaveBeenCalledTimes(3);
  });
});

describe('predictAndValidateTactic', () => {
  const state = { goal: 'Nat', globalContext: [], localContext: [] };

  it('accepts a tactic that parses', async () => {
    mockFetch(json({ tactic: 'intro n', tactic_head: 'intro', category: 'introduction' }));
    expect(await predictAndValidateTactic('http://localhost:8000', state)).toEqual(prediction());
  });

  it('marks unparsable output as not validated', async () => {
    mockFetch(json({ tactic: 'no-such-tactic n' }));
    expect(await predictAndValidateTactic('http://localhost:8000', state)).toMatchObject({
      validated: false,
      category: 'unknown',
    });
  });

  it('rejects multi-line output and server failures', async () => {
    mockFetch(json({ tactic: 'intro n\nexact n' }));
    expect(await predictAndValidateTactic('http://localhost:8000', state)).toBeNull();
    mockFetch(json({}, 500));
    expect(await predictAndValidateTactic('http://localhost:8000', state)).toBeNull();
  });
});
