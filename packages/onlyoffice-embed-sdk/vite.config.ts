import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";

const fromPackageRoot = (path: string) =>
  fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
  // Library assets must resolve from the installed module instead of the
  // consumer application's origin root. This keeps the emitted x2t Worker
  // usable for applications deployed below a pathname prefix.
  base: "./",
  publicDir: false,
  build: {
    target: "es2022",
    copyPublicDir: false,
    emptyOutDir: true,
    sourcemap: true,
    lib: {
      entry: {
        index: fromPackageRoot("./src/index.ts"),
        compat: fromPackageRoot("./src/compat.ts"),
        "compat-subframe": fromPackageRoot("./src/compat-subframe.ts"),
      },
      formats: ["es"],
      fileName: (_format, entryName) => `${entryName}.js`,
    },
    rollupOptions: {
      // Keep ExcelJS external so its transitive packages retain the license
      // metadata shipped by their own npm distributions.
      external: [/^exceljs(?:\/.*)?$/],
      output: {
        assetFileNames: "assets/[name]-[hash][extname]",
        chunkFileNames: "chunks/[name]-[hash].js",
      },
    },
  },
  worker: {
    format: "es",
  },
});
