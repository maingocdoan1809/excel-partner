import { z } from "zod";

/** Normalised native provider questions; this is not an MCP tool. */
export const UserInputRequestSchema = z.object({
  questions: z.array(z.object({
    id: z.string().min(1).max(500),
    header: z.string().max(500).optional(),
    question: z.string().min(1).max(20_000),
    options: z.array(z.object({
      label: z.string().min(1).max(1000),
      description: z.string().max(4000).optional(),
    })).max(30).default([]),
    multiSelect: z.boolean().default(false),
    allowFreeform: z.boolean().default(true),
    isSecret: z.boolean().default(false),
  })).min(1).max(10),
}).refine(value => new Set(value.questions.map(q => q.id)).size === value.questions.length, "Duplicate question IDs");

export const UserInputResponseSchema = z.object({
  cancelled: z.boolean(),
  answers: z.record(z.array(z.string().min(1).max(20_000)).max(30)),
});
export type UserInputRequest = z.infer<typeof UserInputRequestSchema>;
export type UserInputResponse = z.infer<typeof UserInputResponseSchema>;
export const CANCELLED_USER_INPUT: UserInputResponse = { cancelled: true, answers: {} };

export function validateUserInputResponse(request: UserInputRequest, data: unknown): UserInputResponse {
  const response = UserInputResponseSchema.parse(data);
  if (response.cancelled) return { cancelled: true, answers: {} };
  const ids = new Set(request.questions.map(q => q.id));
  if (Object.keys(response.answers).some(id => !ids.has(id))) throw new Error("Unknown question ID");
  for (const question of request.questions) {
    const answers = response.answers[question.id];
    if (!answers?.length || (!question.multiSelect && answers.length !== 1)) throw new Error("Answer every question");
    if (new Set(answers).size !== answers.length) throw new Error("Duplicate answers");
    if (!question.allowFreeform && answers.some(answer => !question.options.some(option => option.label === answer))) {
      throw new Error("Choose a provided option");
    }
  }
  return response;
}
