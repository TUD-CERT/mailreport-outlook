/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://mailsecurity.cert.tu-dresden.de/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

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
  constructor(deployment) {
    this.deployment = deployment;
  }

  apply(compiler) {
    compiler.hooks.environment.tap("ConfigGeneratorPlugin", () => {
      const deploymentPath = path.join(__dirname, "configs", this.deployment);
      if (!fs.existsSync(deploymentPath)) throw new Error(`Error: Deployment ${this.deployment} not found`);
      console.log(`Deployment: ${this.deployment}`);
      const overrides = JSON.parse(fs.readFileSync(path.join(deploymentPath, "overrides.json")));

      // defaults.json
      const cfgOutputPath = path.join(__dirname, "src", "defaults.json");
      console.log(`Generating ${cfgOutputPath}`);
      const templateConfig = JSON.parse(fs.readFileSync(path.join(__dirname, "src", "templates", "defaults.tpl")));
      const resultConfig = updateTemplate(templateConfig, overrides.defaults);
      fs.writeFileSync(cfgOutputPath, JSON.stringify(resultConfig));

      // locales.json
      const resultLocales = {};
      const localesOutputPath = path.join(__dirname, "src", "locales.json");
      console.log(`Generating ${localesOutputPath}`);
      const templateLocales = path.join(__dirname, "src", "templates", "locales");
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
    });
  }
}

module.exports = async (env, options) => {
  if (!("config" in env)) throw new Error("Error: No deployment config supplied, use --env config=<config_name>");
  const dev = options.mode === "development";

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
      options: "./src/options/options.ts",
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
              presets: ["@babel/preset-typescript"],
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
            from: "./src/templates/images/*",
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
        filename: "options.html",
        template: "./src/options/options.html",
        chunks: ["polyfill", "options"],
      }),
      new ConfigGeneratorPlugin(env.config),
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
