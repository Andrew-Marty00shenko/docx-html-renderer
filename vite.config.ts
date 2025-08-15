import { defineConfig } from "vite";

export default defineConfig({
  build: {
    lib: {
      entry: "src/docx-preview.ts",
      name: "docx",
    },
    sourcemap: true,
    minify: "esbuild",
    rollupOptions: {
      external: ["jszip"],
      output: [
        {
          format: "umd",
          name: "docx",
          globals: { jszip: "JSZip" },
          entryFileNames: "docx-preview.js",
          sourcemap: true,
        },
        {
          format: "es",
          entryFileNames: "docx-preview.mjs",
          sourcemap: true,
        },
      ],
    },
  },
});
