import devCerts from "office-addin-dev-certs";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import ReactRefreshWebpackPlugin from "@pmmmwh/react-refresh-webpack-plugin";
import { BundleAnalyzerPlugin } from "webpack-bundle-analyzer";
import webpack from "webpack";
import { fileURLToPath } from "url";
import path from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

export default async (env, options) => {
  const dev = options.mode === "development";
  const isAnalyze = process.env.ANALYZE === "true";

  // package.json 里 config.ai_backend_port = 3001
  const aiBackendPort = Number(process.env.npm_package_config_ai_backend_port || 3001);

  const config = {
    devtool: dev ? "eval-source-map" : "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
      publicPath: "/",
      filename: dev ? "[name].js" : "[name].[contenthash].js",
      chunkFilename: dev ? "[name].chunk.js" : "[name].[contenthash].chunk.js",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js", ".jsx"],
      alias: {
        react: "react",
        "react-dom": "react-dom",
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: [
            {
              loader: "ts-loader",
              options: {
                transpileOnly: true,
                compilerOptions: {
                  jsx: "react-jsx",
                },
              },
            },
          ],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.css$/,
          use: [
            "style-loader",
            {
              loader: "css-loader",
              options: {
                modules: {
                  auto: true,
                  localIdentName: dev ? "[path][name]__[local]" : "[hash:base64:8]",
                },
              },
            },
          ],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
        {
          test: /\.(woff|woff2|eot|ttf|otf)$/,
          type: "asset/resource",
          generator: {
            filename: "fonts/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
        minify: !dev
          ? {
              removeComments: true,
              collapseWhitespace: true,
              removeRedundantAttributes: true,
              useShortDoctype: true,
              removeEmptyAttributes: true,
              removeStyleLinkTypeAttributes: true,
              keepClosingSlash: true,
              minifyJS: true,
              minifyCSS: true,
              minifyURLs: true,
            }
          : false,
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
        minify: !dev
          ? {
              removeComments: true,
              collapseWhitespace: true,
              removeRedundantAttributes: true,
              useShortDoctype: true,
              removeEmptyAttributes: true,
              removeStyleLinkTypeAttributes: true,
              keepClosingSlash: true,
              minifyJS: true,
              minifyCSS: true,
              minifyURLs: true,
            }
          : false,
      }),
      new webpack.DefinePlugin({
        "process.env.NODE_ENV": JSON.stringify(options.mode),
        "process.env.REACT_VERSION": JSON.stringify("19.0.0"),
      }),
      ...(dev
        ? [
            new ReactRefreshWebpackPlugin({
              overlay: { sockIntegration: "wds" },
            }),
          ]
        : []),
      ...(isAnalyze
        ? [
            new BundleAnalyzerPlugin({
              analyzerMode: "static",
              reportFilename: "bundle-report.html",
              openAnalyzer: false,
            }),
          ]
        : []),
    ],
    optimization: {
      splitChunks: {
        chunks: "all",
        minSize: 20000,
        maxSize: 244000,
        cacheGroups: {
          vendors: {
            test: /[\\/]node_modules[\\/]/,
            name: "vendors",
            chunks: "all",
            priority: -10,
            reuseExistingChunk: true,
          },
          reactVendor: {
            test: /[\\/]node_modules[\\/](react|react-dom)[\\/]/,
            name: "react-vendor",
            chunks: "all",
            priority: 20,
            reuseExistingChunk: true,
          },
          fluentUI: {
            test: /[\\/]node_modules[\\/]@fluentui[\\/]/,
            name: "fluentui-vendor",
            chunks: "all",
            priority: 15,
            reuseExistingChunk: true,
          },
          default: {
            minChunks: 2,
            priority: -20,
            reuseExistingChunk: true,
          },
        },
      },
      runtimeChunk: "single",
      minimize: !dev,
      usedExports: true,
      sideEffects: true,
    },
    performance: {
      maxAssetSize: 244 * 1024,
      maxEntrypointSize: 244 * 1024,
      hints: dev ? false : "warning",
    },
    devServer: {
      // ✅ Office Add-in 需要 https
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },

      port: Number(process.env.npm_package_config_dev_server_port || 3000),

      hot: true,
      liveReload: true,

      allowedHosts: "all",

      client: {
        overlay: { errors: true, warnings: false },
        progress: true,
      },

      static: {
        directory: path.resolve(__dirname, "dist"),
        watch: true,
      },

      historyApiFallback: true,
      compress: true,

      /**
       * ✅ 关键：同源代理到 AI 后端，彻底解决：
       * - CORS
       * - Mixed Content（https 页面调 http 后端）
       * - Office WebView 对跨域的限制
       */
      proxy: [
        {
          context: ["/api", "/chat", "/health", "/perceive", "/batch", "/agent"],
          target: `http://localhost:${aiBackendPort}`,
          changeOrigin: true,
          secure: false,
          ws: false,
          logLevel: "warn",
          onProxyReq(proxyReq, req) {
            // 方便你排查：看前端到底有没有把请求打出来
            console.log(`[proxy] ${req.method} ${req.url} -> ${proxyReq.getHeader("host")}`);
          },
        },
      ],
    },
    cache: {
      type: "filesystem",
      buildDependencies: { config: [__filename] },
    },
  };

  return config;
};
