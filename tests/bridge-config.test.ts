import { afterEach, describe, expect, it, vi } from "vitest";
import { loadConfig } from "../bridge/src/config.js";

afterEach(() => { vi.unstubAllEnvs(); });

describe("bridge transport", () => {
  it("refuses to fall back to plain HTTP when the certificate is missing", () => {
    vi.stubEnv("BRIDGE_CERT", "Z:\\missing\\localhost.crt");
    vi.stubEnv("BRIDGE_KEY", "Z:\\missing\\localhost.key");
    expect(() => loadConfig()).toThrow(/certificate is missing/);
  });

  it("serves HTTP only when development asks for it or code sets it up", () => {
    vi.stubEnv("BRIDGE_CERT", "Z:\\missing\\localhost.crt");
    vi.stubEnv("BRIDGE_TLS", "0");
    expect(loadConfig()).toMatchObject({ tls: null, publicBaseUrl: "http://127.0.0.1:8765" });
    vi.stubEnv("BRIDGE_TLS", "1");
    expect(loadConfig({ tls: null })).toMatchObject({ tls: null, publicBaseUrl: "http://127.0.0.1:8765" });
  });
});
