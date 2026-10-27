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
              .filter(name => name.endsWith(".template.docx") || name.endsWith(".template.xlsx") || name.endsWith(".json"))
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
  // @silurus/ooxml 0.89.6's optional legacyXls loader uses a runtime import URL.
  // MonoReport never configures legacyXls; suppress only this known warning.
  // Reassess on viewer upgrades or legacy support: Feature 013 plan, section 11.
  ignoreWarnings: [{
    module: /[\\/]node_modules[\\/]@silurus[\\/]ooxml[\\/]dist[\\/]model-source-COXegBNA\.js$/,
    message: /^Critical dependency: the request of a dependency is an expression$/
  }],
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
    client: {
      overlay: {
        errors: true,
        warnings: false,
      },
    },
    port: 3000,
    historyApiFallback: true,
    proxy: [
      {
        context: ["/api", "/docs", "/openapi"],
        target: process.env.MONOREPORT_API_URL || "http://localhost:5000"
      }]
  },
  performance: { hints: false }
};
