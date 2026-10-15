/**
 * Protocol version negotiated between the task pane (client) and the bridge.
 * Bump when the wire format changes in a backwards-incompatible way.
 */
export const PROTOCOL_VERSION = 1 as const;

/** Minimum bridge version the shipped add-in is allowed to talk to. */
export const MIN_SUPPORTED_PROTOCOL_VERSION = 1 as const;
