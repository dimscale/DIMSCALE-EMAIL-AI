/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://dimscale.github.io/DIMSCALE-EMAIL-AI/"; // GitHub Pages, see ADR-0004

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const prototype = dev && env.prototype === "true";
  if (!dev && urlProd.includes("contoso.com")) {
    throw new Error("Production build blocked: set urlProd in webpack.config.js to the real hosting URL (it is still contoso.com).");
  }
  if (!dev && (!urlProd.startsWith("https://") || !urlProd.endsWith("/"))) {
    throw new Error("Production build blocked: urlProd must start with https:// and end with /.");
  }
  if (!dev) {
    const cfgText = require("fs").readFileSync(require("path").join(__dirname, "src/taskpane/config.ts"), "utf8");
    if (cfgText.includes("00000000-0000-0000-0000-000000000000") || cfgText.includes("//PROJECT-REF.")) {
      throw new Error("Production build blocked: src/taskpane/config.ts still has placeholder values (00000000-... or //PROJECT-REF.).");
    }
  }
  const config = {
    devtool: dev ? "source-map" : false, // no source maps in the published artifact (F-011)
    entry: prototype ? {
      prototype: {
        import: "./src/taskpane/prototype/index.tsx",
      },
    } : {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
    },
    output: {
      clean: true,
      // publicPath stays default "auto" so the /DIMSCALE-EMAIL-AI/ subpath works with relative URLs
      ...(dev ? {} : { filename: "[name].[contenthash].js", chunkFilename: "[name].[contenthash].js" }),
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      ...(prototype ? [
        new HtmlWebpackPlugin({
          filename: "prototype.html",
          template: "./src/taskpane/prototype/prototype.html",
          chunks: ["prototype"],
        }),
      ] : [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      // Second ribbon button page (same bundle; the pane reads location.pathname), F-012/AC-48
      new HtmlWebpackPlugin({
        filename: "resumo.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      ]),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
