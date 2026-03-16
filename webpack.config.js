const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, argv) => {
  const isProd = argv.mode === "production";

  const config = {
    entry: "./src/taskpane/index.tsx",
    output: {
      path: path.resolve(__dirname, "docs"),
      filename: "taskpane.bundle.js",
      publicPath: isProd ? "/word_flowkit/" : "/",
      clean: true,
    },
    resolve: {
      extensions: [".tsx", ".ts", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/index.html",
        filename: "index.html",
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets" }],
      }),
    ],
  };

  if (!isProd) {
    const devCerts = require("office-addin-dev-certs");
    config.devServer = {
      port: 3000,
      server: {
        type: "https",
        options: devCerts.getHttpsServerOptions(),
      },
      hot: true,
      headers: { "Access-Control-Allow-Origin": "*" },
    };
  }

  return config;
};
