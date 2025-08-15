import js from "@eslint/js";
import tseslint from "typescript-eslint";
import prettier from "eslint-config-prettier";

export default tseslint.config(
	{
		ignores: ["dist/**", "node_modules/**"]
	},
	js.configs.recommended,
	...tseslint.configs.recommendedTypeChecked,
	...tseslint.configs.stylisticTypeChecked,
	{
		files: ["**/*.{ts,tsx,js}"],
		languageOptions: {
			parser: tseslint.parser,
			parserOptions: {
				project: ["./tsconfig.json"],
				tsconfigRootDir: import.meta.dirname,
				sourceType: "module",
				ecmaVersion: 2023
			},
			globals: {
				window: "readonly",
				document: "readonly"
			}
		},
		rules: {
			"no-console": "warn"
		}
	},
	prettier
);


