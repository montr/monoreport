const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MonacoWebpackPlugin = require("monaco-editor-webpack-plugin");
const appVersion = process.env.MONOREPORT_VERSION || "dev";

class PlaygroundExamplesPlugin {
  apply(compiler) {
    compiler.hooks.thisCompilation.tap("PlaygroundExamplesPlugin", compilation => {
      compilation.hooks.processAssets.tap(
        { name: "PlaygroundExamplesPlugin", stage: compiler.webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONAL },
        () => {
          const files = [
            [path.resolve(__dirname, "public/examples.json"), "examples.json"],
            ...["LICENSE", "THIRD_PARTY_NOTICES.md"].map(name => [
              path.resolve(__dirname, "node_modules/@silurus/ooxml", name), `notices/ooxml/${name}`
            ]),
            [path.resolve(__dirname, "public/notices/xlsx-viewer.txt"), "notices/xlsx-viewer.txt"],
            ...fs.readdirSync(path.resolve(__dirname, "../Samples"))
              .filter(name => name.endsWith(".template.docx") || name.endsWith(".json"))
              .map(name => [path.resolve(__dirname, "../Samples", name), `examples/${name}`])
          ];
          for (const [source, target] of files) compilation.emitAsset(target, new compiler.webpack.sources.RawSource(fs.readFileSync(source)));
        }
      );
    });
  }
}

module.exports = {
  entry: "./src/index.tsx",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "assets/[name].[contenthash].js",
    clean: true,
    publicPath: "/"
  },
  resolve: { extensions: [".tsx", ".ts", ".js"] },
  module: {
    rules: [
      { test: /\.tsx?$/, exclude: /node_modules/, use: "swc-loader" },
      { test: /\.css$/, use: ["style-loader", "css-loader"] },
      { test: /\.ttf$/, type: "asset/resource" }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({ template: "./public/index.html" }),
    new PlaygroundExamplesPlugin(),
    new MonacoWebpackPlugin({ languages: ["json"] }),
    new webpack.DefinePlugin({ __APP_VERSION__: JSON.stringify(appVersion) })
  ],
  devServer: {
    port: 3000,
    historyApiFallback: true,
    proxy: [{ context: ["/api", "/docs", "/openapi"], target: process.env.MONOREPORT_API_URL || "http://127.0.0.1:5080" }]
  },
  performance: { hints: false }
};
