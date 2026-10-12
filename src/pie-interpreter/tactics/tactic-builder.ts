// Builds a Tactic from the proof editor's (tacticType, params) form. Missing
// parameters are reported as an error; a malformed Pie expression throws, as
// Parser.parsePie does.

import { Parser } from '../parser/parser';
import type { TacticParams } from '../protocol';
import { Location, Syntax } from '../utils/locations';
import { Position } from '../../scheme-parser/transpiler/types/location';
import {
  Tactic,
  IntroTactic,
  ExactTactic,
  ExistsTactic,
  SpiltTactic,
  LeftTactic,
  RightTactic,
  EliminateNatTactic,
  EliminateListTactic,
  EliminateEitherTactic,
  EliminateAbsurdTactic,
} from './tactics';

export type TacticBuildResult = { tactic: Tactic } | { error: string };

type Builder = (location: Location, params: TacticParams) => TacticBuildResult;

/** Tactics that need a Pie expression; `variableName` is optional. */
function withExpression(name: string, make: (loc: Location, params: TacticParams & { expression: string }) => Tactic): Builder {
  return (location, params) => params.expression
    ? { tactic: make(location, { ...params, expression: params.expression }) }
    : { error: `${name} tactic requires an expression parameter` };
}

/** Eliminators: they need the name of the hypothesis to eliminate. */
function withTarget(name: string, make: (loc: Location, target: string) => Tactic): Builder {
  return (location, params) => params.variableName
    ? { tactic: make(location, params.variableName) }
    : { error: `${name} tactic requires a target variable name` };
}

const BUILDERS: Record<string, Builder> = {
  intro: (loc, p) => ({ tactic: new IntroTactic(loc, p.variableName) }),
  exact: withExpression('exact', (loc, p) => new ExactTactic(loc, Parser.parsePie(p.expression))),
  exists: withExpression('exists', (loc, p) => new ExistsTactic(loc, Parser.parsePie(p.expression), p.variableName)),
  split: (loc) => ({ tactic: new SpiltTactic(loc) }),
  left: (loc) => ({ tactic: new LeftTactic(loc) }),
  right: (loc) => ({ tactic: new RightTactic(loc) }),
  elimNat: withTarget('elimNat', (loc, t) => new EliminateNatTactic(loc, t)),
  induction: withTarget('elimNat', (loc, t) => new EliminateNatTactic(loc, t)),
  elimList: withTarget('elimList', (loc, t) => new EliminateListTactic(loc, t)),
  elimEither: withTarget('elimEither', (loc, t) => new EliminateEitherTactic(loc, t)),
  elimAbsurd: withTarget('elimAbsurd', (loc, t) => new EliminateAbsurdTactic(loc, t)),
};

export function buildTacticFromParams(
  tacticType: string,
  params: TacticParams = {},
  location: Location = syntheticLocation(),
): TacticBuildResult {
  const build = Object.prototype.hasOwnProperty.call(BUILDERS, tacticType) ? BUILDERS[tacticType] : undefined;
  return build ? build(location, params) : { error: `Unknown tactic type: ${tacticType}` };
}

/** Location for tactics that do not come from source text. */
export function syntheticLocation(): Location {
  const pos = new Position(1, 0);
  return new Location(new Syntax(pos, pos, ''), false);
}
