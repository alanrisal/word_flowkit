const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = {
  entry: "./src/taskpane/index.tsx",
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "taskpane.bundle.js",
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
      patterns: [
        {
          from: "assets",
          to: "assets",
        },
      ],
    }),
  ],
  devServer: {
    port: 3000,
    server: {
      type: "https",
      options: devCerts.getHttpsServerOptions(),
    },
    hot: true,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
};
