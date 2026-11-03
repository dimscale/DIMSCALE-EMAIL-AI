/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
// ADR-0007: the host comes from PANE_BASE_URL (closed list); the default keeps GitHub Pages until the cut-over.
const PROD_URLS = [
  "https://dimscale.github.io/DIMSCALE-EMAIL-AI/",
  "https://dimscale-mail-ai.pages.dev/",
  "https://dev.dimscale-mail-ai.pages.dev/",
  "https://mailai.dimscale.com/",
];
// Cloudflare Pages sets CF_PAGES_BRANCH during its builds: the dev branch always becomes the test add-in.
const cfDev = process.env.CF_PAGES_BRANCH === "dev";
const urlProd = process.env.PANE_BASE_URL || (cfDev ? PROD_URLS[2] : PROD_URLS[0]);
// A test build gives the manifest its own Id and name, so it can sit next to the production add-in.
const paneTest = process.env.PANE_TEST === "1" || cfDev;
const PROD_ID = "05c2e1c9-3e1d-406e-9a91-e9ac64854143";
const TEST_ID = "926a7aa2-46e8-4187-9111-c9873e1fdec5";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const prototype = dev && env.prototype === "true";
  if (!dev && !PROD_URLS.includes(urlProd)) {
    throw new Error(`Production build blocked: PANE_BASE_URL must be one of ${PROD_URLS.join(", ")}.`);
  }
  if (!dev) console.log(`[manifest] host ${urlProd}, test add-in: ${paneTest}`); // shows in the Cloudflare build log
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
                // Cloudflare Pages redirects *.html to the extensionless path; GitHub Pages serves both, so use /taskpane.
                let m = content.toString().replace(new RegExp(urlDev, "g"), urlProd).replace(/taskpane\.html/g, "taskpane");
                if (paneTest) {
                  m = m.replace(PROD_ID, TEST_ID).replace('<DisplayName DefaultValue="DIMSCALE Mail"/>', '<DisplayName DefaultValue="DIMSCALE Mail (teste)"/>');
                  if (!m.includes(TEST_ID) || !m.includes("(teste)")) throw new Error("PANE_TEST build: manifest Id or name was not replaced.");
                }
                return m;
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
