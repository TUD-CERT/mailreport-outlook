/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const LOCAL_DEV_URL = "https://localhost:3000";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

/**
 * Performs a deep update of a template object with values from overrides. Returns the updated template.
 */
function updateTemplate(template, overrides) {
  Object.entries(overrides).forEach(([key, value]) => {
    if (value instanceof Object) {
      if (!(key in template)) template[key] = {};
      updateTemplate(template[key], value);
    } else template[key] = value;
  });
  return template;
}

class ConfigGeneratorPlugin {
  constructor(deployment, isDev) {
    this.deployment = deployment;
    this.isDev = isDev;
  }

  apply(compiler) {
    compiler.hooks.environment.tap("ConfigGeneratorPlugin", () => {
      const deploymentPath = path.join(__dirname, "configs", this.deployment);
      if (!fs.existsSync(deploymentPath)) throw new Error(`Error: Deployment ${this.deployment} not found`);
      console.log(`Deployment: ${this.deployment}`);
      const overrides = JSON.parse(fs.readFileSync(path.join(deploymentPath, "overrides.json")));
      const pluginID = `${overrides.manifest.id}@${overrides.manifest.provider_name}/${overrides.manifest.version}`;
      console.log(`Plugin ID: ${pluginID}`);

      // Generate defaults.json
      const cfgOutputPath = path.join(__dirname, "src", "defaults.json");
      console.log(`Generating ${cfgOutputPath}`);
      const templateConfig = JSON.parse(fs.readFileSync(path.join(__dirname, "templates", "defaults.tpl")));
      const resultConfig = updateTemplate(templateConfig, overrides.defaults);
      resultConfig["plugin_id"] = pluginID;
      fs.writeFileSync(cfgOutputPath, JSON.stringify(resultConfig));

      // Generate locales.json
      const resultLocales = {};
      const localesOutputPath = path.join(__dirname, "src", "locales.json");
      console.log(`Generating ${localesOutputPath}`);
      const templateLocales = path.join(__dirname, "templates", "locales");
      fs.readdirSync(templateLocales).forEach((locale) => {
        const localePath = path.join(templateLocales, locale);
        if (fs.statSync(localePath).isDirectory()) {
          const localeContent = JSON.parse(fs.readFileSync(path.join(localePath, "messages.json")));
          if ("locales" in overrides && locale in overrides.locales)
            updateTemplate(localeContent, overrides.locales[locale]);
          resultLocales[locale] = localeContent;
        }
      });
      fs.writeFileSync(localesOutputPath, JSON.stringify(resultLocales));

      // Generate manifest.xml
      const manifestOutputPath = path.join(__dirname, "manifest.xml"),
        hostedAt = this.isDev ? LOCAL_DEV_URL : overrides.manifest.hosted_at;
      console.log(`Generating ${manifestOutputPath}`);
      let manifest = fs.readFileSync(path.join(__dirname, "templates", "manifest.tpl"), { encoding: "utf8" });
      manifest = manifest
        .replaceAll("__ID__", overrides.manifest.id)
        .replaceAll("__VERSION__", overrides.manifest.version)
        .replaceAll("__PROVIDER_NAME__", overrides.manifest.provider_name)
        .replaceAll("__HOSTED_AT__", hostedAt)
        .replaceAll("__DOMAIN__", new URL(hostedAt).origin);
      for (const locale in resultLocales) {
        for (const tag in resultLocales[locale]) {
          manifest = manifest.replaceAll(`__MSG_${tag}_${locale}__`, resultLocales[locale][tag]);
        }
      }
      fs.writeFileSync(manifestOutputPath, manifest);
    });
  }
}

module.exports = async (env, options) => {
  if (!("config" in env)) throw new Error("Error: No deployment config supplied, use --env config=<config_name>");
  const isDev = options.mode === "development";

  // Deployment config image overrides
  const overrideImages = [];
  const imagesOverridePath = path.join(__dirname, "configs", env.config, "images");
  if (fs.existsSync(imagesOverridePath)) {
    overrideImages.push({
      from: `./configs/${env.config}/images/*`,
      to: "assets/[name][ext][query]",
    });
  }

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      report_fraud: "./src/report_fraud/report_fraud.ts",
      commands: "./src/commands/commands.ts",
      error: "./src/error/error.ts",
      options: "./src/options/options.ts",
      simulation_ack: "./src/simulation/simulation_ack.ts",
    },
    output: {
      clean: true,
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
            loader: "babel-loader",
            options: {
              presets: [
                [
                  "@babel/preset-env",
                  {
                    targets: {
                      ie: "11",
                      esmodules: false,
                    },
                  },
                ],
                "@babel/preset-typescript",
              ],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          ...overrideImages,
          {
            from: "./templates/images/*",
            to: "assets/[name][ext][query]",
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "report_fraud.html",
        template: "./src/report_fraud/report_fraud.html",
        chunks: ["polyfill", "report_fraud"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "error.html",
        template: "./src/error/error.html",
        chunks: ["polyfill", "error"],
      }),
      new HtmlWebpackPlugin({
        filename: "options.html",
        template: "./src/options/options.html",
        chunks: ["polyfill", "options"],
      }),
      new HtmlWebpackPlugin({
        filename: "simulation_ack.html",
        template: "./src/simulation/simulation_ack.html",
        chunks: ["polyfill", "simulation_ack"],
      }),
      new ConfigGeneratorPlugin(env.config, isDev),
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
    },
  };

  return config;
};
