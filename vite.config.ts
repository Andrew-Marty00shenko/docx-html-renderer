import { defineConfig } from "vite";

export default defineConfig(({ command }) => {
  if (command === "serve") {
    return {
      server: {
        port: 5173,
        open: true,
      },
      optimizeDeps: {
        include: ["jszip"],
      },
      esbuild: {
        target: "es2020",
      },
    };
  }

  return {
    build: {
      outDir: "lib",
      emptyOutDir: true,
      lib: {
        entry: "src/index.ts",
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
            entryFileNames: "docx-html-renderer.js",
          },
          { format: "es", entryFileNames: "docx-html-renderer.mjs" },
        ],
      },
    },
  };
});
