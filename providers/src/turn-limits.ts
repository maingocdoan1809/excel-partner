/**
 * A turn has no fixed length: a long job (testing a whole sheet of cases, say) runs as long as the CLI keeps working,
 * and ends when the CLI is done or the person presses Stop. Only a CLI that has gone silent is cut off.
 */

/** A CLI that sends nothing at all for this long is taken to be stuck. Every event it sends starts the count again. */
export const TURN_IDLE_LIMIT_MS = 20 * 60_000;

/** Backstop for transports that can only wait for a fixed time: far beyond any real job, so it never ends one early. */
export const TURN_BACKSTOP_MS = 12 * 60 * 60_000;

/** What the person is told when a CLI went silent; "continue" resumes the same conversation. */
export const idleMessage = (provider: string): string =>
  `${provider} sent nothing for ${TURN_IDLE_LIMIT_MS / 60_000} minutes, so the reply was stopped. Send "continue" to carry on.`;
