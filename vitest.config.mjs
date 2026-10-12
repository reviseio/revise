import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    // The CLI tests open and save real .docx files through the engine; on a
    // shared CI runner some take longer than vitest's 5-second default.
    testTimeout: 60_000,
  },
});
