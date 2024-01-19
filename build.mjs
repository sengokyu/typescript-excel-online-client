import * as esbuild from "esbuild";
import * as path from "path";
import packagejson from "./package.json" assert { type: "json" };

const sharedBuildOptions = {
  bundle: true,
  entryPoints: [path.resolve("src", "index.ts")],
  external: Object.keys(packagejson.dependencies),
  platform: "node",
};

await esbuild.build({
  ...sharedBuildOptions,
  format: "esm",
  outfile: "./dist/index.esm.js",
});

await esbuild.build({
  ...sharedBuildOptions,
  format: "cjs",
  outfile: "./dist/index.cjs.js",
});
