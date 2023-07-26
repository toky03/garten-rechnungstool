const path = require("path");

module.exports = {
  entry: "./index.ts",
  target: "node18",
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: "ts-loader",
        exclude: "/node_modules/",
      },
      { test: /\.afm$/, loader: "raw-loader" },
    ],
  },
  resolve: {
    extensions: [".ts", ".js"],
  },
  output: {
    library: "lib",
    libraryTarget: "umd",
    filename: "main.js",
    path: path.resolve(__dirname, "dist"),
    globalObject: "this",
  },
};
