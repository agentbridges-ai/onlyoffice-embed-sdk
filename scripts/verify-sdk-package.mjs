#!/usr/bin/env node

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  appendFile,
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDirectory = fileURLToPath(new URL("..", import.meta.url));
const packageDirectory = path.join(
  rootDirectory,
  "packages",
  "onlyoffice-embed-sdk",
);
const expectedPackageName = "@agentbridges-ai/onlyoffice-embed-sdk";
const maximumPackedBytes = 20 * 1024 * 1024;
const maximumUnpackedBytes = 50 * 1024 * 1024;
const maximumPackedFileCount = 250;
const keepTemporaryDirectory = process.env.SDK_VERIFY_KEEP_TEMP === "1";
const requireBun = process.env.SDK_VERIFY_REQUIRE_BUN === "1";

function printableCommand(command, args) {
  return [command, ...args]
    .map((part) => (/^[\w@./:=+-]+$/.test(part) ? part : JSON.stringify(part)))
    .join(" ");
}

function run(command, args, options = {}) {
  const cwd = options.cwd ?? rootDirectory;
  console.log(`\n> ${printableCommand(command, args)}`);

  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: {
      ...process.env,
      CI: process.env.CI ?? "1",
      FORCE_COLOR: "0",
      NO_COLOR: "1",
    },
    stdio: options.capture ? ["ignore", "pipe", "pipe"] : "inherit",
  });

  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    const diagnostic = [result.stdout, result.stderr].filter(Boolean).join("\n");
    throw new Error(
      `${printableCommand(command, args)} exited with ${result.status}${
        diagnostic ? `\n${diagnostic}` : ""
      }`,
    );
  }

  return result.stdout ?? "";
}

function commandExists(command) {
  const result = spawnSync(command, ["--version"], {
    encoding: "utf8",
    stdio: "pipe",
  });
  return !result.error && result.status === 0;
}

async function sha256(file) {
  return createHash("sha256").update(await readFile(file)).digest("hex");
}

function parsePackResult(output) {
  let result;
  try {
    result = JSON.parse(output);
  } catch (error) {
    throw new Error(`npm pack did not return JSON:\n${output}`, { cause: error });
  }

  assert.equal(result.length, 1, "npm pack must produce exactly one tarball");
  return result[0];
}

function packInto(destination) {
  const output = run(
    "npm",
    [
      "pack",
      "--ignore-scripts",
      "--silent",
      "--json",
      "--pack-destination",
      destination,
      packageDirectory,
    ],
    { capture: true },
  );
  const result = parsePackResult(output);
  return { result, tarball: path.join(destination, result.filename) };
}

function resolveExportTarget(exportValue, condition) {
  if (typeof exportValue === "string") {
    return exportValue;
  }
  assert.ok(
    exportValue && typeof exportValue === "object",
    `export must provide an object or string target for ${condition}`,
  );
  const target = exportValue[condition];
  assert.equal(
    typeof target,
    "string",
    `export must provide a ${condition} target`,
  );
  return target;
}

function toPackedPath(target) {
  return target.replace(/^\.\//, "");
}

function validateMetadata(packageJson, packedFiles, packResult) {
  assert.equal(packageJson.name, expectedPackageName);
  assert.match(
    packageJson.version,
    /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/,
    "package version must be an explicit semver",
  );
  assert.notEqual(packageJson.private, true, "SDK package must be publishable");
  assert.equal(packageJson.type, "module", "SDK package must be ESM");
  assert.equal(packageJson.publishConfig?.access, "public");
  assert.equal(
    packageJson.dependencies?.exceljs,
    "4.4.0",
    "ExcelJS must stay external so its transitive packages retain their own licenses",
  );

  assert.deepEqual(
    new Set(packageJson.files),
    new Set(["dist", "README.md", "LICENSE", "THIRD_PARTY_NOTICES.md"]),
    "package files must remain a strict publish allowlist",
  );
  assert.ok(packageJson.exports?.["."], "root export is required");
  assert.ok(packageJson.exports?.["./compat"], "compat export is required");
  assert.ok(
    packageJson.exports?.["./compat/subframe"],
    "compat subframe export is required",
  );

  const requiredTargets = [packageJson.main, packageJson.types];
  for (const exportName of [".", "./compat", "./compat/subframe"]) {
    requiredTargets.push(
      resolveExportTarget(packageJson.exports[exportName], "import"),
      resolveExportTarget(packageJson.exports[exportName], "types"),
    );
  }

  for (const target of requiredTargets) {
    assert.equal(typeof target, "string", "package entry target must be a string");
    assert.ok(
      packedFiles.has(toPackedPath(target)),
      `published entry is missing from tarball: ${target}`,
    );
  }

  const requiredRootFiles = [
    "package.json",
    "README.md",
    "LICENSE",
    "THIRD_PARTY_NOTICES.md",
  ];
  for (const file of requiredRootFiles) {
    assert.ok(packedFiles.has(file), `tarball is missing ${file}`);
  }

  for (const file of packedFiles) {
    assert.ok(
        file === "package.json" ||
        file === "README.md" ||
        file === "LICENSE" ||
        file === "THIRD_PARTY_NOTICES.md" ||
        file.startsWith("dist/"),
      `tarball contains a file outside the publish allowlist: ${file}`,
    );
    assert.doesNotMatch(
      file,
      /(^|\/)node_modules\/|^(?:public|src|tests?)(?:\/|$)|^\.env(?:\.|$)/,
      `tarball contains a forbidden path: ${file}`,
    );
    assert.doesNotMatch(
      file,
      /(?:^|\/)exceljs(?:[.-]|\/)/i,
      "ExcelJS must not be rebundled into the SDK tarball",
    );
  }

  assert.ok(
    [...packedFiles].some((file) => file.endsWith(".d.ts")),
    "tarball must contain TypeScript declarations",
  );

  const forbiddenPackages = new Set([
    "@agentbridges-ai/onlyoffice-browser",
    "onlyoffice-browser",
    "@agentbridges-ai/onlyoffice-x2t-wasm",
    "onlyoffice-x2t-wasm",
  ]);
  for (const field of [
    "dependencies",
    "devDependencies",
    "peerDependencies",
    "optionalDependencies",
  ]) {
    for (const dependency of Object.keys(packageJson[field] ?? {})) {
      assert.ok(
        !forbiddenPackages.has(dependency),
        `${field} must not depend on ${dependency}`,
      );
    }
  }

  assert.ok(
    packResult.size <= maximumPackedBytes,
    `packed tarball is too large: ${packResult.size} bytes`,
  );
  assert.ok(
    packResult.unpackedSize <= maximumUnpackedBytes,
    `unpacked package is too large: ${packResult.unpackedSize} bytes`,
  );
  assert.ok(
    packedFiles.size <= maximumPackedFileCount,
    `tarball contains too many files: ${packedFiles.size}`,
  );
}

async function listFiles(directory, baseDirectory = directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const absolutePath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await listFiles(absolutePath, baseDirectory)));
    } else if (entry.isFile()) {
      files.push(
        path.relative(baseDirectory, absolutePath).split(path.sep).join("/"),
      );
    }
  }
  return files;
}

async function validateThirdPartyNotices(packageRoot) {
  const notices = await readFile(
    path.join(packageRoot, "THIRD_PARTY_NOTICES.md"),
    "utf8",
  );
  assert.match(notices, /ExcelJS 4\.4\.0/);
  assert.match(notices, /Copyright \(c\) 2014-2019 Guyon Roche/);
  assert.match(notices, /The MIT License \(MIT\)/);
  assert.match(notices, /Brotli\.js 1\.3\.2/);
  assert.match(notices, /foliojs\/brotli\.js/);
  assert.match(notices, /Copyright \(c\) Devon Govett/);
  assert.match(notices, /Modifications by agentbridges-ai/);
  assert.match(notices, /TypeScript\s+and\s+ES modules/);
  assert.match(notices, /dictionary bootstrap/);
  assert.match(notices, /Copyright 2013 Google Inc\./);
  assert.match(notices, /Apache License\s+Version 2\.0, January 2004/);
  assert.match(notices, /END OF TERMS AND CONDITIONS/);
}

async function validateExcelJsExternal(packageRoot) {
  const distDirectory = path.join(packageRoot, "dist");
  const files = await listFiles(distDirectory);
  const emittedSources = files.filter((file) => /\.(?:js|map)$/.test(file));
  const contents = (
    await Promise.all(
      emittedSources.map((file) => readFile(path.join(distDirectory, file), "utf8")),
    )
  ).join("\n");

  assert.match(
    contents,
    /import\(["']exceljs\/dist\/exceljs\.min["']\)/,
    "SDK output must retain ExcelJS as an external dynamic import",
  );
  assert.doesNotMatch(
    contents,
    /node_modules\/exceljs|webpackUniversalModuleDefinition|JSZip\.support/,
    "SDK output appears to contain the prebundled ExcelJS browser runtime",
  );
}

async function validateWorkerOutput(
  outputDirectory,
  label,
  { rejectRootRelative = false, forbidExternalWorker = false } = {},
) {
  const files = await listFiles(outputDirectory);
  const workerFiles = files.filter((file) =>
    /(^|\/)x2t[.-]worker.*\.(?:js|mjs)$/i.test(file),
  );
  if (forbidExternalWorker) {
    assert.equal(
      workerFiles.length,
      0,
      `${label} must keep the x2t worker inline instead of emitting a second JavaScript asset`,
    );
  }

  const javascriptFiles = files.filter((file) =>
    /\.(?:js|mjs)$/.test(file),
  );
  let inlineWorkerFound = false;

  for (const javascriptFile of javascriptFiles) {
    const contents = await readFile(
      path.join(outputDirectory, javascriptFile),
      "utf8",
    );
    if (rejectRootRelative) {
      assert.doesNotMatch(
        contents,
        /new URL\(\s*["'`]\/assets\//,
        `${label} contains a site-root-relative asset URL`,
      );
    }

    const hasWorkerConstructor = /new Worker\(/.test(contents);
    const hasBlobConstructor = /new Blob\(/.test(contents);
    const hasWorkerIdentity = contents.includes("onlyoffice-x2t");
    const hasReadyProtocol =
      /postMessage\(\{\s*type\s*:\s*[`"']ready[`"']/.test(contents);
    if (
      hasWorkerConstructor &&
      hasBlobConstructor &&
      hasWorkerIdentity &&
      hasReadyProtocol
    ) {
      inlineWorkerFound = true;
    }
  }

  assert.ok(
    inlineWorkerFound,
    `${label} does not contain the inline Blob x2t Worker and ready protocol`,
  );
}

async function writeConsumerFixture(
  consumerDirectory,
  tarball,
  rootPackageJson,
  sdkPackageJson,
) {
  const sourceDirectory = path.join(consumerDirectory, "src");
  await mkdir(sourceDirectory, { recursive: true });

  await writeFile(
    path.join(consumerDirectory, "package.json"),
    `${JSON.stringify(
      {
        name: "onlyoffice-embed-sdk-external-consumer",
        version: "0.0.0",
        private: true,
        type: "module",
        dependencies: {
          [expectedPackageName]: `file:${tarball}`,
        },
        devDependencies: {
          typescript: rootPackageJson.devDependencies.typescript,
          vite: rootPackageJson.devDependencies.vite,
        },
      },
      null,
      2,
    )}\n`,
  );

  await writeFile(
    path.join(consumerDirectory, "tsconfig.json"),
    `${JSON.stringify(
      {
        compilerOptions: {
          target: "ES2022",
          module: "ESNext",
          moduleResolution: "Bundler",
          lib: ["ES2022", "DOM", "DOM.Iterable"],
          strict: true,
          noEmit: true,
          skipLibCheck: false,
        },
        include: ["src/**/*.ts"],
      },
      null,
      2,
    )}\n`,
  );

  await writeFile(
    path.join(consumerDirectory, "index.html"),
    '<!doctype html><html><body><main id="app"></main><script type="module" src="/src/browser.ts"></script></body></html>\n',
  );

  await writeFile(
    path.join(sourceDirectory, "browser.ts"),
    `import * as sdk from "${expectedPackageName}";
import * as compat from "${expectedPackageName}/compat";
import * as compatSubframe from "${expectedPackageName}/compat/subframe";
import type {
  CreateOfficeEditorOptions,
  OfficeEditorInstance,
  OfficeEditorMount,
  OfficeHostIdentity,
  OfficeHostUrlContext,
  OfficeHostUrlResolver,
} from "${expectedPackageName}/compat";
import type {
  CreateOfficeEditorOptions as SubframeCreateOfficeEditorOptions,
  OfficeEditorInstance as SubframeOfficeEditorInstance,
  OfficeEditorMount as SubframeOfficeEditorMount,
  OfficeSubframeSlot,
} from "${expectedPackageName}/compat/subframe";

declare const options: CreateOfficeEditorOptions;
declare const instance: OfficeEditorInstance;
declare const mount: OfficeEditorMount;
declare const identity: OfficeHostIdentity;
declare const context: OfficeHostUrlContext;
declare const resolver: OfficeHostUrlResolver;
void [options, instance, mount, identity, context, resolver];
declare const subframeOptions: SubframeCreateOfficeEditorOptions;
declare const subframeInstance: SubframeOfficeEditorInstance;
declare const subframeMount: SubframeOfficeEditorMount;
declare const subframeSlot: OfficeSubframeSlot;
void [subframeOptions, subframeInstance, subframeMount, subframeSlot];

const requiredRootExports = [
  "OnlyOfficeManager",
  "onlyOfficeManagerFactory",
  "registerOnlyOfficeStaticResource",
] as const;
const requiredCompatExports = [
  "createOfficeEditor",
  "mountOfficeEditor",
  "createOfficeRuntimeResourceManager",
  "ONLYOFFICE_EMBED_SDK_VERSION",
] as const;
const requiredCompatSubframeExports = [
  "createOfficeEditor",
  "mountOfficeEditor",
  "getOfficeSubframeOrigin",
  "COMPAT_SUBFRAME_PROTOCOL_SOURCE",
  "HOSTED_COMPAT_SUBFRAME_IDENTITY",
  "ONLYOFFICE_EMBED_HOST_MANIFEST",
] as const;

for (const name of requiredRootExports) {
  if (typeof sdk[name] === "undefined") throw new Error(\`Missing root export: \${name}\`);
}
for (const name of requiredCompatExports) {
  if (typeof compat[name] === "undefined") throw new Error(\`Missing compat export: \${name}\`);
}
for (const name of requiredCompatSubframeExports) {
  if (typeof compatSubframe[name] === "undefined") throw new Error(\`Missing compat subframe export: \${name}\`);
}
if (compat.ONLYOFFICE_EMBED_SDK_VERSION !== ${JSON.stringify(sdkPackageJson.version)}) {
  throw new Error(\`SDK identity version does not match the installed package: \${compat.ONLYOFFICE_EMBED_SDK_VERSION}\`);
}

document.querySelector("#app")!.textContent = "onlyoffice-embed-sdk consumer ready";
`,
  );

  await writeFile(
    path.join(sourceDirectory, "ssr.mjs"),
    `import * as sdk from "${expectedPackageName}";
import * as compat from "${expectedPackageName}/compat";
import * as compatSubframe from "${expectedPackageName}/compat/subframe";

for (const [namespace, values, names] of [
  ["root", sdk, ["OnlyOfficeManager", "onlyOfficeManagerFactory", "registerOnlyOfficeStaticResource"]],
  ["compat", compat, ["createOfficeEditor", "mountOfficeEditor", "createOfficeRuntimeResourceManager", "ONLYOFFICE_EMBED_SDK_VERSION"]],
  ["compat subframe", compatSubframe, ["createOfficeEditor", "mountOfficeEditor", "getOfficeSubframeOrigin", "COMPAT_SUBFRAME_PROTOCOL_SOURCE", "HOSTED_COMPAT_SUBFRAME_IDENTITY", "ONLYOFFICE_EMBED_HOST_MANIFEST"]],
]) {
  for (const name of names) {
    if (typeof values[name] === "undefined") throw new Error(\`Missing \${namespace} export: \${name}\`);
  }
}
if (compat.ONLYOFFICE_EMBED_SDK_VERSION !== ${JSON.stringify(sdkPackageJson.version)}) {
  throw new Error(\`SDK identity version does not match the installed package: \${compat.ONLYOFFICE_EMBED_SDK_VERSION}\`);
}

console.log("onlyoffice-embed-sdk SSR import ready");
`,
  );

  await writeFile(
    path.join(consumerDirectory, "vite.browser.config.mjs"),
    `import { defineConfig } from "vite";

export default defineConfig({
  build: {
    outDir: "dist-vite-browser",
    emptyOutDir: true,
  },
});
`,
  );

  await writeFile(
    path.join(consumerDirectory, "vite.ssr.config.mjs"),
    `import { defineConfig } from "vite";

export default defineConfig({
  build: {
    ssr: "src/ssr.mjs",
    outDir: "dist-vite-ssr",
    emptyOutDir: true,
    rollupOptions: {
      output: { entryFileNames: "entry.mjs" },
    },
  },
  ssr: {
    noExternal: ["${expectedPackageName}"],
  },
});
`,
  );
}

async function verifyExternalConsumer(consumerDirectory) {
  run(
    "npm",
    [
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      "--package-lock=false",
    ],
    { cwd: consumerDirectory },
  );
  run("npm", ["exec", "--", "tsc", "--noEmit"], {
    cwd: consumerDirectory,
  });
  run("node", ["src/ssr.mjs"], { cwd: consumerDirectory });
  run("npm", ["exec", "--", "vite", "build", "--config", "vite.browser.config.mjs"], {
    cwd: consumerDirectory,
  });
  await validateWorkerOutput(
    path.join(consumerDirectory, "dist-vite-browser"),
    "external Vite browser build",
  );
  run("npm", ["exec", "--", "vite", "build", "--config", "vite.ssr.config.mjs"], {
    cwd: consumerDirectory,
  });
  run("node", ["dist-vite-ssr/entry.mjs"], { cwd: consumerDirectory });

  if (!commandExists("bun")) {
    const message = "Bun is unavailable; Bun import/build checks were skipped";
    if (requireBun) {
      throw new Error(message);
    }
    console.warn(`\nSKIP: ${message}`);
    return;
  }

  run("bun", ["run", "src/ssr.mjs"], { cwd: consumerDirectory });
  run(
    "bun",
    [
      "build",
      "src/browser.ts",
      "--target=browser",
      "--outdir=dist-bun-browser",
    ],
    { cwd: consumerDirectory },
  );
  await validateWorkerOutput(
    path.join(consumerDirectory, "dist-bun-browser"),
    "external Bun browser build",
  );
}

async function main() {
  const packageJson = JSON.parse(
    await readFile(path.join(packageDirectory, "package.json"), "utf8"),
  );
  const rootPackageJson = JSON.parse(
    await readFile(path.join(rootDirectory, "package.json"), "utf8"),
  );

  assert.equal(packageJson.name, expectedPackageName);
  assert.notEqual(
    rootPackageJson.name,
    packageJson.name,
    "the private application root must not masquerade as the SDK package",
  );
  assert.equal(rootPackageJson.private, true, "the application root must be private");

  run("pnpm", ["--filter", expectedPackageName, "run", "typecheck"]);
  run("pnpm", ["--filter", expectedPackageName, "run", "build"]);

  const temporaryDirectory = await mkdtemp(
    path.join(os.tmpdir(), "onlyoffice-embed-sdk-package-"),
  );
  const relativeToRepository = path.relative(rootDirectory, temporaryDirectory);
  assert.ok(
    relativeToRepository.startsWith("..") && !path.isAbsolute(relativeToRepository),
    "consumer fixture must be created outside the repository",
  );

  try {
    const firstPackDirectory = path.join(temporaryDirectory, "pack-one");
    const secondPackDirectory = path.join(temporaryDirectory, "pack-two");
    const extractedDirectory = path.join(temporaryDirectory, "extracted");
    const consumerDirectory = path.join(temporaryDirectory, "consumer");
    await Promise.all([
      mkdir(firstPackDirectory),
      mkdir(secondPackDirectory),
      mkdir(extractedDirectory),
      mkdir(consumerDirectory),
    ]);

    const firstPack = packInto(firstPackDirectory);
    await rm(path.join(packageDirectory, "dist"), {
      recursive: true,
      force: true,
    });
    run("pnpm", ["--filter", expectedPackageName, "run", "build"]);
    const secondPack = packInto(secondPackDirectory);
    const [firstHash, secondHash] = await Promise.all([
      sha256(firstPack.tarball),
      sha256(secondPack.tarball),
    ]);
    assert.equal(
      firstHash,
      secondHash,
      "two clean SDK build-and-pack runs must be byte-for-byte reproducible",
    );

    const packedFiles = new Set(
      firstPack.result.files.map(({ path: file }) => file),
    );
    validateMetadata(packageJson, packedFiles, firstPack.result);

    run("tar", ["-xzf", firstPack.tarball, "-C", extractedDirectory]);
    const publishedPackageJson = JSON.parse(
      await readFile(
        path.join(extractedDirectory, "package", "package.json"),
        "utf8",
      ),
    );
    validateMetadata(publishedPackageJson, packedFiles, firstPack.result);
    await validateThirdPartyNotices(
      path.join(extractedDirectory, "package"),
    );
    await validateExcelJsExternal(path.join(extractedDirectory, "package"));
    await validateWorkerOutput(
      path.join(extractedDirectory, "package", "dist"),
      "published SDK",
      { rejectRootRelative: true, forbidExternalWorker: true },
    );

    await writeConsumerFixture(
      consumerDirectory,
      firstPack.tarball,
      rootPackageJson,
      publishedPackageJson,
    );
    await verifyExternalConsumer(consumerDirectory);

    let publishedTarball = firstPack.tarball;
    if (process.env.SDK_VERIFY_TARBALL_DIR) {
      const outputDirectory = path.resolve(
        process.env.SDK_VERIFY_TARBALL_DIR,
      );
      await mkdir(outputDirectory, { recursive: true });
      publishedTarball = path.join(outputDirectory, firstPack.result.filename);
      await copyFile(firstPack.tarball, publishedTarball);
    }
    if (process.env.GITHUB_OUTPUT) {
      await appendFile(
        process.env.GITHUB_OUTPUT,
        `tarball=${publishedTarball}\nsha256=${firstHash}\n`,
      );
    }

    console.log("\nSDK package verification passed");
    console.log(`  package: ${packageJson.name}@${packageJson.version}`);
    console.log(`  tarball: ${publishedTarball}`);
    console.log(`  sha256: ${firstHash}`);
    console.log(
      `  size: ${firstPack.result.size} bytes packed / ${firstPack.result.unpackedSize} bytes unpacked`,
    );
  } finally {
    if (keepTemporaryDirectory) {
      console.log(`\nTemporary verification directory kept at ${temporaryDirectory}`);
    } else {
      await rm(temporaryDirectory, { recursive: true, force: true });
    }
  }
}

main().catch((error) => {
  console.error("\nSDK package verification failed");
  console.error(error);
  process.exitCode = 1;
});
