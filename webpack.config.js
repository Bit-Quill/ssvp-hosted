/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://bit-quill.github.io/ssvp-hosted/"; // GitHub Pages deployment location

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
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
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
        {
          test: /\.(woff|woff2|eot|ttf|otf)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/fonts/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        'process.env.NODE_ENV': JSON.stringify(dev ? 'development' : 'production'),
        'process.env.OAUTH_REDIRECT_URI': JSON.stringify(
          dev ? 'http://localhost:3000/auth/callback' : 'https://bit-quill.github.io/ssvp-hosted/auth/callback'
        ),
      }),
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
            from: "assets/fonts/*",
            to: "assets/fonts/[name][ext][query]",
          },
          {
            from: "src/auth/*.html",
            to: "auth/[name][ext]",
          },
          {
            from: dev ? "manifest.xml" : "manifest.production.xml",
            to: "manifest.xml",
            transform(content) {
              if (dev) {
                return content;
              } else {
                // In production, copy manifest.production.xml as manifest.xml
                return content;
              }
            },
          },
          // Copy public readme as index.html for GitHub Pages landing page
          ...(!dev ? [{
            from: "public-readme.html",
            to: "index.html",
          }] : []),
        ],
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
      setupMiddlewares: (middlewares, devServer) => {
        const path = require('path');
        const express = require('express');

        // Add JSON body parser
        devServer.app.use(express.json());
        devServer.app.use(express.urlencoded({ extended: true }));

        // OAuth auth start route
        devServer.app.get('/auth/auth-start.html', (req, res) => {
          res.sendFile(path.join(__dirname, 'src/auth/auth-start.html'));
        });

        // OAuth callback route
        devServer.app.get('/auth/callback', (req, res) => {
          // Serve the callback HTML page
          res.sendFile(path.join(__dirname, 'src/auth/callback.html'));
        });

        // Test route
        devServer.app.get('/api/test', (req, res) => {
          res.json({ status: 'ok', message: 'API is working!' });
        });

        // Auth API routes (token exchange proxy)
        try {
          const authRoutes = require('./server/routes/auth');
          devServer.app.use('/api/auth', authRoutes);
        } catch (error) {
          // Auth routes failed to load
        }

        return middlewares;
      },
    },
  };

  return config;
};
