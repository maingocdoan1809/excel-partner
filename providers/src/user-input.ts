import { UserInputRequestSchema, validateUserInputResponse, CANCELLED_USER_INPUT, type UserInputRequest } from "../../protocol/src/index.js";
import type { StartSessionContext } from "./types.js";

export type RequestUserInput = StartSessionContext["requestUserInput"];

export async function askProviderQuestions(request: UserInputRequest, ask: RequestUserInput, signal?: AbortSignal) {
  if (!ask || signal?.aborted) return CANCELLED_USER_INPUT;
  const parsed = UserInputRequestSchema.parse(request);
  const response = await ask(parsed, signal);
  return signal?.aborted ? CANCELLED_USER_INPUT : validateUserInputResponse(parsed, response);
}

export async function answerCodexQuestions(params: Record<string, any>, ask: RequestUserInput, signal?: AbortSignal) {
  const request = UserInputRequestSchema.parse({ questions: params.questions?.map((q: Record<string, any>) => ({
    ...q, options: q.options ?? [], allowFreeform: !q.options?.length || q.isOther === true,
  })) });
  const response = await askProviderQuestions(request, ask, signal);
  return { answers: Object.fromEntries(request.questions.map(q => [q.id, { answers: response.cancelled ? [] : response.answers[q.id] }])) };
}

export async function answerClaudeQuestion(input: Record<string, any>, ask: RequestUserInput, signal?: AbortSignal) {
  const request = UserInputRequestSchema.parse({ questions: input.questions?.map((q: Record<string, any>, i: number) => ({
    ...q, id: String(i), allowFreeform: true,
  })) });
  const response = await askProviderQuestions(request, ask, signal);
  if (response.cancelled) return { behavior: "deny", message: "The user cancelled the question. Do not assume an answer." };
  return { behavior: "allow", updatedInput: { ...input,
    answers: Object.fromEntries(request.questions.map(q => [q.question, response.answers[q.id]!.join(", ")])),
  } };
}
