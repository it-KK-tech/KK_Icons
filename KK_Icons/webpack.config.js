/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require('path');
// Load env from project dir if present (dev only)
require('dotenv').config({ path: path.resolve(__dirname, '.env') });

const urlDev = "https://localhost:3000/";
const urlProd = process.env.APP_BASE_URL || "https://your-app-service.azurewebsites.net/"; // Set APP_BASE_URL during CI/build

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
      path: path.resolve(__dirname, 'dist'),
      publicPath: '/',
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
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
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: [
        {
          context: '/api/streamline',
          target: 'https://public-api.streamlinehq.com/v1',
          changeOrigin: true,
          secure: true,
          pathRewrite: {
            '^/api/streamline': ''
          },
          onProxyReq: (proxyReq, req, res) => {
            const apiKey = process.env.STREAMLINE_API_KEY || '';
            if (!apiKey) {
              // Log a friendly warning for developers running locally
              // Do not throw; allow app to run but API calls will fail with 401
              console.warn('[devServer] STREAMLINE_API_KEY is not set. Set it in a .env file for local dev.');
            }
            proxyReq.setHeader('x-api-key', apiKey);
            // Set appropriate accept header based on the request path
            if (req.url.includes('/download/svg')) {
              proxyReq.setHeader('accept', 'image/svg+xml');
            } else {
              proxyReq.setHeader('accept', 'application/json');
            }
          },
          logLevel: 'info'
        }
      ]
    },
  };

  return config;
};
