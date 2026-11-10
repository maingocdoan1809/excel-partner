import { configDefaults, defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["tests/**/*.test.ts"],
    exclude: [...configDefaults.exclude, ...(process.env.LIVE_PROVIDERS?.trim() ? [] : ["**/*.live.test.ts"])],
    environment: "node",
    testTimeout: 15000,
  },
});
