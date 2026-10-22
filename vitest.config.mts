import { defineConfig } from "vitest/config";

// testTimeout: the cold Fluent UI + React import under jsdom takes >5 s on this OneDrive path
export default defineConfig({
  test: { environment: "jsdom", globals: false, exclude: ["node_modules/**", "supabase/**"], testTimeout: 20000 },
});
