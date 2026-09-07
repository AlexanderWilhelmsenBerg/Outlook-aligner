export default [
  {
    files: ["**/*.js"],
    ignores: ["dist/", "node_modules/"],
    rules: {
      "no-undef": "error",
      "no-unused-vars": "error",
    },
  },
];
