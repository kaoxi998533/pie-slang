
export { todoQueue, addTodo, clearQueue, type TodoInfo } from "./todo-queue";

export function clearCache(): void {
  // No cached state when running in the browser.
}

export async function solveTodo(): Promise<string> {
  throw new Error("TODO solving via LLM is not available in the browser environment.");
}
