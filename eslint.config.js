import js from "@eslint/js";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";

export default tseslint.config(
  {
    ignores: ["**/lib/**", "**/node_modules/**", "**/dist/**"],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  ...tseslint.configs.stylistic,
  {
    files: ["src/**/*.{ts,tsx}"],
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        sourceType: "module",
        ecmaVersion: 2023,
      },
      globals: {
        window: "readonly",
        document: "readonly",
      },
    },
    rules: {
      "no-console": "warn",
      "no-var": "off",
      "no-case-declarations": "off",
      "no-fallthrough": "off",
      "@typescript-eslint/consistent-type-definitions": "off",
      "@typescript-eslint/consistent-type-assertions": "off",

      "@typescript-eslint/no-empty-object-type": "off",
      "@typescript-eslint/no-inferrable-types": "off",
      "@typescript-eslint/prefer-for-of": "off",
      "@typescript-eslint/consistent-type-imports": [
        "error",
        {
          prefer: "type-imports",
          disallowTypeAnnotations: false,
        },
      ],
      "@typescript-eslint/no-unused-vars": [
        "warn",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
    },
  },
  prettier,
);
