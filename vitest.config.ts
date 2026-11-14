import { configDefaults, defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["tests/**/*.test.ts"],
    exclude: [...configDefaults.exclude, ...(process.env.LIVE_PROVIDERS?.trim() ? [] : ["**/*.live.test.ts"])],
    environment: "node",
    setupFiles: ["tests/support/hermetic.ts"],
    testTimeout: 15000,
  },
});
