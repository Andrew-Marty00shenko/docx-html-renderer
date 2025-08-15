import { defineConfig } from "vite";

export default defineConfig({
  build: {
    outDir: "lib",
    emptyOutDir: true,
    lib: {
      entry: "src/docx2html.ts",
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
          entryFileNames: "docx2html.js",
        },
        {
          format: "es",
          entryFileNames: "docx2html.mjs",
        },
      ],
    },
  },
});
