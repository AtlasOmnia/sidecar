const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

module.exports = async (_env, argv) => {
  const isDev = argv.mode === "development";
  const acceptanceEnabled = process.env.HERMES_OFFICE_ACCEPTANCE === "1";

  return {
    devtool: isDev ? "source-map" : false,
    entry: {
      omniaTheme: "./src/taskpane/omniaThemeBootstrap.ts",
      taskpane: "./src/taskpane/taskpane.ts",
      ...(acceptanceEnabled ? { workbookPlanAcceptance: "./src/taskpane/workbookPlanAcceptanceEntry.ts" } : {})
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
      new webpack.DefinePlugin({
        __HERMES_OFFICE_ACCEPTANCE__: JSON.stringify(acceptanceEnabled)
      }),
      new HtmlWebpackPlugin({
        chunks: acceptanceEnabled ? ["omniaTheme", "taskpane", "workbookPlanAcceptance"] : ["omniaTheme", "taskpane"],
        inject: false,
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
          { from: "src/taskpane/omnia-tokens.css", to: "omnia-tokens.css" },
          { from: "src/taskpane/omnia-pane.css", to: "omnia-pane.css", info: { minimized: true } },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          { from: "assets", to: "assets" }
        ]
      })
    ]
  };
};
