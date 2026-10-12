import 'jest';
import { buildTacticFromParams, syntheticLocation } from '../tactics/tactic-builder';
import { ProofManager } from '../tactics/proof-manager';
import { evaluatePieAndGetContext } from '../main';
import { go, stop } from '../types/utils';
import type { Tactic } from '../tactics/tactics';

function built(tacticType: string, params = {}): Tactic {
  const result = buildTacticFromParams(tacticType, params);
  if ('error' in result) throw new Error(result.error);
  return result.tactic;
}

describe('buildTacticFromParams', () => {
  it.each([
    ['intro', { variableName: 'n' }, 'intro n'],
    ['exact', { expression: '(same 0)' }, 'exact'],
    ['exists', { expression: '0', variableName: 'half' }, 'exists'],
    ['split', {}, 'split-Pair'],
    ['left', {}, 'go-Left'],
    ['right', {}, 'go-Right'],
    ['elimNat', { variableName: 'n' }, 'elim-Nat n'],
    ['induction', { variableName: 'n' }, 'elim-Nat n'],
    ['elimList', { variableName: 'xs' }, 'elim-List xs'],
    ['elimEither', { variableName: 'e' }, 'elim-Either e'],
    ['elimAbsurd', { variableName: 'a' }, 'elim-Absurd a'],
  ])('builds %s', (tacticType, params, expected) => {
    expect(built(tacticType, params).toString()).toContain(expected);
  });

  it('allows intro without a variable name', () => {
    expect(built('intro').toString()).toContain('intro');
  });

  it.each([
    ['exact', 'exact tactic requires an expression parameter'],
    ['exists', 'exists tactic requires an expression parameter'],
    ['elimNat', 'elimNat tactic requires a target variable name'],
    ['induction', 'elimNat tactic requires a target variable name'],
    ['elimList', 'elimList tactic requires a target variable name'],
    ['elimEither', 'elimEither tactic requires a target variable name'],
    ['elimAbsurd', 'elimAbsurd tactic requires a target variable name'],
  ])('reports missing parameters for %s', (tacticType, error) => {
    expect(buildTacticFromParams(tacticType, {})).toEqual({ error });
  });

  it('rejects unknown tactic types, including Object.prototype keys', () => {
    expect(buildTacticFromParams('frobnicate')).toEqual({ error: 'Unknown tactic type: frobnicate' });
    expect(buildTacticFromParams('toString')).toEqual({ error: 'Unknown tactic type: toString' });
  });

  it('throws on a malformed expression, like Parser.parsePie', () => {
    expect(() => buildTacticFromParams('exact', { expression: '(same' })).toThrow();
  });

  it('produces tactics the proof manager can apply', () => {
    const { context } = evaluatePieAndGetContext(`
(claim + (→ Nat Nat Nat))
(define + (λ (n j) (iter-Nat n j (λ (n-1) (add1 n-1)))))
(claim zero+n (Π ((n Nat)) (= Nat (+ 0 n) n)))
`);
    const pm = new ProofManager();
    expect(pm.startProof('zero+n', context, syntheticLocation())).toBeInstanceOf(go);
    expect(pm.applyTactic(built('intro', { variableName: 'n' }))).toBeInstanceOf(go);
    expect(pm.applyTactic(built('exact', { expression: '(same n)' }))).toBeInstanceOf(go);
    expect(pm.currentState!.isComplete()).toBe(true);
  });

  it('lets the kernel reject an ill-typed tactic', () => {
    const { context } = evaluatePieAndGetContext(`(claim t (= Nat 0 0))`);
    const pm = new ProofManager();
    pm.startProof('t', context, syntheticLocation());
    expect(pm.applyTactic(built('elimNat', { variableName: 'nope' }))).toBeInstanceOf(stop);
  });
});
