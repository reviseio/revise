#!/usr/bin/env node
/**
 * Builds the `revise` npm package into dist/, exactly as it is published:
 *
 *   npm run build     # dist/
 *   npm run pack      # revise-<version>.tgz from dist/
 *
 * The library (src/lib/) and the CLI (src/cli/) are bundled with esbuild;
 * @reviseio/engine, jszip, and yjs stay external dependencies. The engine is
 * a separate, closed package: the build fails if anything outside src/lib/
 * and src/cli/ ends up in the bundle.
 */
import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import { builtinModules } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build as esbuild } from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dist = path.join(root, "dist");
const typesDir = path.join(root, ".types");
const manifest = JSON.parse(
  await fs.readFile(path.join(root, "package.json"), "utf8"),
);
const { version } = manifest;
const reviseBuild = manifest.reviseBuild ?? null;

await fs.rm(dist, { recursive: true, force: true });
await fs.rm(typesDir, { recursive: true, force: true });
await fs.mkdir(dist, { recursive: true });

// ── 1. bundle ────────────────────────────────────────────────────────────
// `import text from "./x.md?raw"` (the agent skill): the file as a string,
// as Vite reads it in tests.
const markdownAsText = {
  name: "markdown-as-text",
  setup(buildApi) {
    // Strip `?raw`; the `.md` loader reads the file as text.
    buildApi.onResolve({ filter: /\.md\?raw$/ }, (args) => ({
      path: path.resolve(args.resolveDir, args.path.replace(/\?raw$/, "")),
    }));
  },
};

const result = await esbuild({
  absWorkingDir: root,
  entryPoints: {
    index: path.join(root, "src", "lib", "index.ts"),
    cli: path.join(root, "src", "cli", "bin.ts"),
  },
  outdir: dist,
  bundle: true,
  splitting: true,
  format: "esm",
  platform: "node",
  target: "node20",
  chunkNames: "chunks/[name]-[hash]",
  packages: "external",
  plugins: [markdownAsText],
  loader: { ".md": "text" },
  define: {
    "globalThis.__REVISE_SDK_VERSION__": JSON.stringify(version),
    __REVISE_SDK_VERSION__: JSON.stringify(version),
    __REVISE_SDK_BUILD__: JSON.stringify(reviseBuild),
  },
  banner: { js: "// revise: MIT License. https://revise.io/developer/docs" },
  metafile: true,
  logLevel: "warning",
});

// Only the library and the CLI may be in here: an engine module would put
// closed code into an MIT tarball.
const strays = Object.keys(result.metafile.inputs).filter(
  (input) => !/^src\/(lib|cli)\//.test(input),
);
if (strays.length > 0) {
  throw new Error(
    `revise would bundle modules outside src/lib/ and src/cli/:\n  ${strays.join("\n  ")}`,
  );
}

const cliPath = path.join(dist, "cli.js");
const cli = await fs.readFile(cliPath, "utf8");
await fs.writeFile(
  cliPath,
  `#!/usr/bin/env node\n${cli.replace(/^#!.*\n/, "")}`,
);
await fs.chmod(cliPath, 0o755);

// ── 2. declarations ──────────────────────────────────────────────────────
const tsc = (...tscArgs) =>
  new Promise((resolve) => {
    const child = spawn(
      process.execPath,
      [path.join(root, "node_modules", "typescript", "bin", "tsc"), ...tscArgs],
      { cwd: root },
    );
    let text = "";
    child.stdout.on("data", (chunk) => (text += chunk));
    child.stderr.on("data", (chunk) => (text += chunk));
    child.on("exit", (code) => resolve({ code, text }));
  });
// Emitted without checking the implementation: @reviseio/engine publishes
// declarations for its API surface, looser than the engine's own source, so
// function bodies that reach into engine modules don't type-check here.
// The public signatures are all explicit; step 4 type-checks the emitted
// declarations the way a consumer compiles them.
const emitted = await tsc(
  "-p",
  path.join(root, "tsconfig.types.json"),
  "--noCheck",
  "--outDir",
  typesDir,
  "--pretty",
  "false",
);
if (emitted.code !== 0) {
  throw new Error(`revise's declarations didn't emit:\n${emitted.text}`);
}
const libTypes = path.join(typesDir, "lib");
for (const file of await fs.readdir(libTypes)) {
  if (!file.endsWith(".d.ts")) continue;
  const text = (await fs.readFile(path.join(libTypes, file), "utf8"))
    // NodeNext consumers resolve declaration imports like runtime ones,
    // so relative specifiers need the extension of the file they name.
    .replace(/((?:from\s*|import\(\s*)["'])(\.\/[\w-]+)(["'])/g, "$1$2.js$3");
  if (/from "\.\.\//.test(text) || /import\("\.\.\//.test(text)) {
    throw new Error(
      `revise's declarations reach outside the library: lib/${file}`,
    );
  }
  await fs.writeFile(path.join(dist, file), text);
}
await fs.rm(typesDir, { recursive: true, force: true });

// ── 3. package files ─────────────────────────────────────────────────────
// The npm README is this repo's README without its "This repo" section.
const readme = await fs.readFile(path.join(root, "README.md"), "utf8");
await fs.writeFile(
  path.join(dist, "README.md"),
  `${readme.split(/^## This repo$/m)[0].trimEnd()}\n`,
);
for (const file of ["CHANGELOG.md", "LICENSE"]) {
  await fs
    .copyFile(path.join(root, file), path.join(dist, file))
    .catch((error) => {
      if (file !== "CHANGELOG.md") throw error;
    });
}
// The --json contract (revise.io/developer/docs/revise/json), as a JSON Schema.
await fs.copyFile(
  path.join(root, "src", "cli", "output.schema.json"),
  path.join(dist, "cli-output.schema.json"),
);
// revise 2.x was the editor. Its import paths fail with a pointer to
// @reviseio/editor instead of an opaque "subpath not exported".
await fs.cp(path.join(root, "legacy"), path.join(dist, "legacy"), {
  recursive: true,
});

// The published manifest: this repo's package.json with paths relative to
// dist/, and without the fields that only matter for developing here.
const chunks = await fs.readdir(path.join(dist, "chunks")).catch(() => []);
const fromDist = (value) =>
  typeof value === "string"
    ? value.replace(/^\.\/dist\//, "./")
    : Array.isArray(value)
      ? value.map(fromDist)
      : value && typeof value === "object"
        ? Object.fromEntries(
            Object.entries(value).map(([key, inner]) => [key, fromDist(inner)]),
          )
        : value;
const published = {};
for (const [key, value] of Object.entries(manifest)) {
  if (["scripts", "devDependencies", "private"].includes(key)) {
    continue;
  }
  published[key] = fromDist(value);
}
published.files = [
  "index.js",
  "cli.js",
  "legacy",
  "cli-output.schema.json",
  "*.d.ts",
  ...(chunks.length > 0 ? ["chunks"] : []),
  "README.md",
  "CHANGELOG.md",
  "LICENSE",
];
// revise is published by Revise's release pipeline, not from a build here.
published.scripts = {
  prepublishOnly:
    "node -e \"if(process.env.REVISE_SDK_VERIFIED_PUBLISH!=='1'){console.error('Refusing to publish: revise is published by the Revise release, not from this build.');process.exit(1)}\"",
};
await fs.writeFile(
  path.join(dist, "package.json"),
  `${JSON.stringify(published, null, 2)}\n`,
);

// ── 4. audit ─────────────────────────────────────────────────────────────
await audit(dist, published);
// The declarations compile for a strict NodeNext consumer, against the
// engine's published types.
const checked = await tsc(
  "--noEmit",
  "--strict",
  "--module",
  "nodenext",
  "--moduleResolution",
  "nodenext",
  "--target",
  "es2022",
  "--skipLibCheck",
  "false",
  "--types",
  "node",
  "--pretty",
  "false",
  path.join(dist, "index.d.ts"),
);
if (checked.code !== 0) {
  throw new Error(`revise's declarations don't type-check:\n${checked.text}`);
}
console.log(`Built revise ${version} in dist/`);

async function listFiles(directory, base = directory) {
  const out = [];
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) out.push(...(await listFiles(full, base)));
    else out.push(path.relative(base, full).split(path.sep).join("/"));
  }
  return out;
}

/** What the package must and must not carry. */
async function audit(directory, pkg) {
  const allowed = [
    /^package\.json$/,
    /^README\.md$/,
    /^CHANGELOG\.md$/,
    /^LICENSE$/,
    /^(index|cli)\.js$/,
    /^cli-output\.schema\.json$/,
    /^[\w-]+\.d\.ts$/,
    /^chunks\/[\w.-]+\.js$/,
    /^legacy\/(editor|server)\.(js|d\.ts)$|^legacy\/style\.css$/,
  ];
  const forbidden = [
    [/\/Users\/[A-Za-z]|\/home\/[a-z][\w-]*\//, "a local home-directory path"],
    [/\b(?:localhost|127\.0\.0\.1):\d{2,5}/, "a local development host"],
    [/\b[\w-]+\.(?:internal|corp|local)\b(?!\s*\))/, "an internal host name"],
    [/sk-ant-[\w-]{10,}|sk-[A-Za-z0-9]{32,}|AKIA[0-9A-Z]{16}/, "a credential"],
    [/-----BEGIN [A-Z ]*PRIVATE KEY-----/, "a private key"],
    [/sourceMappingURL=/, "a source map reference"],
  ];
  const declared = new Set([
    ...Object.keys(pkg.dependencies ?? {}),
    ...Object.keys(pkg.peerDependencies ?? {}),
    ...Object.keys(pkg.optionalDependencies ?? {}),
  ]);
  const problems = [];
  let javascriptBytes = 0;
  for (const file of await listFiles(directory)) {
    if (!allowed.some((pattern) => pattern.test(file))) {
      problems.push(`${file}: not on the published-file allowlist`);
      continue;
    }
    const text = await fs.readFile(path.join(directory, file), "utf8");
    for (const [pattern, what] of forbidden) {
      if (pattern.test(text)) problems.push(`${file}: contains ${what}`);
    }
    if (!file.endsWith(".js")) continue;
    javascriptBytes += text.length;
    // Every package the code imports must be declared, or a clean install
    // fails to resolve it. Template literals hold text, not imports: the
    // bundled agent skill quotes example code.
    const code = text
      .replaceAll("\\\n", "")
      .replace(/`(?:\\[\s\S]|[^`\\])*`/g, "``");
    for (const [, specifier] of code.matchAll(
      /(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s*)["']((?:@[a-z0-9][\w.-]*\/)?[a-z0-9][\w.-]*(?:\/[\w./-]*)?)["']/gi,
    )) {
      if (specifier.startsWith("node:")) continue;
      const name = specifier.startsWith("@")
        ? specifier.split("/").slice(0, 2).join("/")
        : specifier.split("/")[0];
      if (!builtinModules.includes(name) && !declared.has(name)) {
        problems.push(
          `${file}: imports ${name}, which package.json doesn't declare`,
        );
      }
    }
  }
  // The library and CLI are a few hundred KB; megabytes means the engine
  // came along.
  if (javascriptBytes > 1_500_000) {
    problems.push(
      `JavaScript totals ${Math.round(javascriptBytes / 1000)} KB: the engine looks bundled in`,
    );
  }
  if (pkg.name !== "revise") problems.push(`package name is ${pkg.name}`);
  if (pkg.license !== "MIT") problems.push("license is not MIT");
  if (!pkg.dependencies?.["@reviseio/engine"]) {
    problems.push("does not depend on @reviseio/engine");
  }
  const license = await fs
    .readFile(path.join(directory, "LICENSE"), "utf8")
    .catch(() => "");
  if (!/^MIT License/m.test(license))
    problems.push("LICENSE is not the MIT text");
  for (const hook of ["preinstall", "install", "postinstall", "prepare"]) {
    if (pkg.scripts?.[hook]) problems.push(`install hook present: ${hook}`);
  }
  if (problems.length > 0) {
    throw new Error(
      `The revise package failed its audit:\n  ${problems.join("\n  ")}`,
    );
  }
}
