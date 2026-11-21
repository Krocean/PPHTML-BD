// Dev-time build of the tool bundles that ship inside the skill (assets\tools). Needs Node; in this folder:
//   npm install
//   node build.js
// The skill itself never needs Node: the bundles run inside Edge (headless) on the user's PC.
const esbuild = require("esbuild");
const fs = require("fs");
const path = require("path");
const lab = __dirname;
const skill = path.join(__dirname, "..");
const outDir = path.join(skill, "assets", "tools");
fs.mkdirSync(outDir, { recursive: true });

const nodePaths = (process.env.NODE_PATH || "").split(path.delimiter).filter(Boolean);
const modules = name => { for (const p of [path.join(lab, "node_modules")].concat(nodePaths)) { if (fs.existsSync(path.join(p, name))) { return path.join(p, name); } } throw new Error("module not found: " + name + " (run npm install)"); };
(async () => {
  if (process.argv.indexOf("themes") >= 0) {
    // theme files from the design systems' published tokens (see src/themes.js)
    const out = path.join(lab, "out", "themes.cjs");
    await esbuild.build({ entryPoints: [path.join(__dirname, "src", "themes.js")], bundle: true, nodePaths: [path.join(lab, "node_modules")].concat(nodePaths), format: "cjs", platform: "node", target: "node18",
      loader: { ".css": "text" }, outfile: out, logLevel: "warning" });
    require("child_process").execFileSync(process.execPath, [out, path.join(skill, "assets", "kit")], { stdio: "inherit" });
    return;
  }
  await esbuild.build({
    entryPoints: [path.join(__dirname, "src", "convert.js")],
    bundle: true, nodePaths: [path.join(lab, "node_modules")].concat(nodePaths), format: "iife", globalName: "IHConvert", platform: "browser", target: "es2020", minify: true, legalComments: "none",
    define: { "process.env.NODE_ENV": '"production"', "process.env.BROWSERSLIST": "undefined", "process.env.BROWSERSLIST_ENV": "undefined", "process.env.BROWSERSLIST_CONFIG": "undefined", "process.env.BROWSERSLIST_DISABLE_CACHE": "undefined", "process.env.BROWSERSLIST_IGNORE_OLD_DATA": "undefined", "process.env.BROWSERSLIST_STATS": "undefined", "process.env.BROWSERSLIST_ROOT_PATH": "undefined", "process.env.BROWSERSLIST_DANGEROUS_EXTEND": "undefined", "process.env.AUTOPREFIXER_GRID": "undefined", "process.env.FORCE_COLOR": "undefined", "process.env.NO_COLOR": "undefined", "process.env.DEBUG": "undefined", "process.argv": "[]", "process.platform": '"browser"', "process.version": '"v0"', "process.versions.node": '"0"' },
    banner: { js: "/* ppt-html build tool: PostCSS, Autoprefixer, postcss-nesting, postcss-css-variables, postcss-flexbugs-fixes, acorn (all MIT). Runs in Edge headless at build time; never shipped inside a presentation. */" },
    outfile: path.join(outDir, "convert.js"), logLevel: "warning"
  });
  await esbuild.build({
    entryPoints: [path.join(__dirname, "src", "fonttool.js")],
    bundle: true, nodePaths: [path.join(lab, "node_modules")].concat(nodePaths), format: "iife", globalName: "IHFont", platform: "browser", target: "es2020", minify: true, legalComments: "none",
    loader: { ".wasm": "base64" },
    banner: { js: "/* ppt-html build tool: HarfBuzz subsetter (harfbuzzjs, MIT) + pako (MIT). Runs in Edge headless at build time. */" },
    outfile: path.join(outDir, "fonttool.js"), logLevel: "warning"
  });
  // the same bundle as a CommonJS module, for the dev-time font pack script
  await esbuild.build({
    entryPoints: [path.join(__dirname, "src", "fonttool.js")],
    bundle: true, nodePaths: [path.join(lab, "node_modules")].concat(nodePaths), format: "cjs", platform: "node", target: "node18", loader: { ".wasm": "base64" },
    outfile: path.join(lab, "out", "fonttool.cjs"), logLevel: "warning"
  });
  fs.copyFileSync(path.join(modules("core-js-bundle"), "minified.js"), path.join(skill, "assets", "corejs.min.txt"));
  fs.copyFileSync(path.join(modules("@babel/standalone"), "babel.min.js"), path.join(outDir, "babel.min.js"));
  for (const f of fs.readdirSync(outDir)) { console.log(f.padEnd(24), (fs.statSync(path.join(outDir, f)).size / 1024).toFixed(0) + " KB"); }
})().catch(e => { console.error(e); process.exit(1); });
