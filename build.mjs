import * as esbuild from "esbuild";
import * as path from "path";

const external = [
  "@microsoft/kiota-abstractions",
  "@microsoft/msgraph-sdk",
  "@microsoft/msgraph-sdk-drives",
];

const sharedBuildOptions = {
  bundle: true,
  entryPoints: [path.resolve("src", "index.ts")],
  target: "es2022",
  sourcemap: true,
  external,
};

await esbuild.build({
  ...sharedBuildOptions,
  format: "esm",
  outdir: path.resolve("dist", "esm"),
});

await esbuild.build({
  ...sharedBuildOptions,
  format: "cjs",
  outdir: path.resolve("dist", "cjs"),
  outExtension: { ".js": ".cjs" },
});
