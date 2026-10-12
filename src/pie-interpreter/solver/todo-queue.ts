// TODOs collected by the typechecker for later solving. Kept apart from
// todo-solver.ts so the typechecker does not depend on any LLM code.

import type { Location } from "../utils/locations";
import type { Context } from "../utils/context";
import type { Value } from "../types/value";
import type { Renaming } from "../typechecker/utils";

export interface TodoInfo {
  location: Location;
  context: Context;
  expectedType: Value;
  renaming: Renaming;
  name?: string;
}

export const todoQueue: TodoInfo[] = [];

export function addTodo(todo: TodoInfo): void {
  todoQueue.push(todo);
}

export function clearQueue(): void {
  todoQueue.length = 0;
}
