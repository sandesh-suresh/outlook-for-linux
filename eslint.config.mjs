import globals from "globals";
import pluginJs from "@eslint/js";

export default [
  { ignores: ["dist/**", "out/**"] },
  { files: ["**/*.js"], languageOptions: { sourceType: "commonjs" } },
  { languageOptions: { globals: { ...globals.browser, ...globals.node } } },
  pluginJs.configs.recommended,
  {
    rules: {
      "no-var": "error",
      eqeqeq: ["error", "always"],
    },
  },
];
