// SPDX-License-Identifier: Apache-2.0
// Builds the static Sidecar for PowerPoint sidebar. Production use: `npm run build:static -- --base <https site>/`
// (scripts/build-static.mjs), which also renders manifest.xml for that site. No dev server, no certificates.
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

module.exports = (_env, argv) => {
  const isDev = argv.mode === "development";

  return {
    devtool: isDev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts"
    },
    output: {
      clean: true,
      filename: "[name].[contenthash].js",
      path: path.resolve(__dirname, "dist")
    },
    resolve: {
      extensions: [".ts", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: "ts-loader",
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"]
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        chunks: ["taskpane"],
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html"
      }),
      new HtmlWebpackPlugin({
        chunks: [],
        filename: "commands.html",
        template: "./src/commands/commands.html"
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/omnia-tokens.css", to: "omnia-tokens.css" },
          { from: "src/taskpane/omnia-pane.css", to: "omnia-pane.css", info: { minimized: true } },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          { from: "src/taskpane/theme-bootstrap.js", to: "theme-bootstrap.js" }
        ]
      })
    ]
  };
};
