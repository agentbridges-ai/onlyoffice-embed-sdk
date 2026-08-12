#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { brotliDecompressSync } from "node:zlib";
import { fileURLToPath } from "node:url";

const repositoryRoot = fileURLToPath(new URL("..", import.meta.url));
const lockPath = path.join(repositoryRoot, "config", "x2t-release.lock.json");

function sha256(data) {
  return createHash("sha256").update(data).digest("hex");
}

function sha512(data) {
  return createHash("sha512").update(data).digest("hex");
}

function assertDigest(value, algorithm, label) {
  const length = algorithm === "sha256" ? 64 : 128;
  assert.equal(typeof value, "string", `${label} must be a string`);
  assert.match(value, new RegExp(`^[a-f0-9]{${length}}$`), `${label} is invalid`);
}

function assertFileIdentity(identity, label) {
  assert.ok(identity && typeof identity === "object", `${label} is missing`);
  assert.ok(Number.isSafeInteger(identity.size) && identity.size > 0, `${label}.size is invalid`);
  assertDigest(identity.sha256, "sha256", `${label}.sha256`);
  assertDigest(identity.sha512, "sha512", `${label}.sha512`);
}

async function verifyFile(filePath, identity, label) {
  const contents = await readFile(filePath);
  assert.equal(contents.byteLength, identity.size, `${label} size changed`);
  assert.equal(sha256(contents), identity.sha256, `${label} SHA-256 changed`);
  assert.equal(sha512(contents), identity.sha512, `${label} SHA-512 changed`);
  return contents;
}

async function main() {
  const lockBytes = await readFile(lockPath);
  const lock = JSON.parse(lockBytes.toString("utf8"));

  assert.equal(lock.schemaVersion, 1, "unsupported x2t release lock schema");
  assert.equal(lock.repository, "agentbridges-ai/onlyoffice-x2t-wasm");
  assert.match(lock.tag, /^v\d+\.\d+\.\d+\+\d+$/);
  assert.match(lock.sourceCommit, /^[a-f0-9]{40}$/);
  assert.equal(lock.toolchain?.onlyofficeCore, "v9.3.0.140");
  assert.equal(lock.toolchain?.emscripten, "4.0.11");
  assert.deepEqual(lock.toolchain?.linkerFlags, [
    "-sALLOW_MEMORY_GROWTH",
    "-sEMULATE_FUNCTION_POINTER_CASTS=1",
    "-sEXPORTED_FUNCTIONS=_main1",
  ]);
  assert.equal(lock.parameters?.isNoBase64, false);
  assert.equal(lock.parameters?.usesRealInputFileName, true);
  assert.equal(lock.parameters?.tempDirectory, "/tmp/x2t-conversion");
  assert.equal(lock.parameters?.docFormatFrom, 66);
  assert.equal(lock.parameters?.docFormatTo, 8193);
  assert.equal(lock.parameters?.spreadsheetFormatFrom, 257);
  assert.equal(lock.parameters?.spreadsheetFormatTo, 8194);
  assert.equal(
    lock.releaseUrl,
    `https://github.com/${lock.repository}/releases/tag/${lock.tag}`,
  );
  assert.ok(
    typeof lock.signerWorkflow === "string" &&
      lock.signerWorkflow.startsWith(`${lock.repository}/.github/workflows/`) &&
      lock.signerWorkflow.endsWith("@refs/tags/" + lock.tag),
    "x2t signer workflow is not bound to the immutable release tag",
  );

  assertFileIdentity(lock.archive, "archive");
  assertFileIdentity(lock.regressionLog, "regressionLog");
  assertFileIdentity(lock.files?.script, "files.script");
  assertFileIdentity(lock.files?.scriptBrotli, "files.scriptBrotli");
  assertFileIdentity(lock.files?.wasmRaw, "files.wasmRaw");
  assertFileIdentity(lock.files?.wasmBrotli, "files.wasmBrotli");

  const expectedRegressions = {
    legacyDoc: {
      inputSha256: "d85e44ae5368ccbbe57ded8533ced05a250c30cfa15da10f19fdaf63f080238c",
      outputSha256: "074a9b350ff6a6e1ee32866c03416a0682c05635cdb8f3f60b6e4a02eaad9a2a",
      outputSize: 132030,
      header: "DOCY;v5;",
    },
    pivotSlicer: {
      inputSha256: "ffecc0a33c9e41b392fbee30127a97f3e5c3577c717be103471460bd07c2ec58",
      outputSha256: "c40fb3f4f67311426110d4786eb4684981aec9ed05b4f13c6367c8470de4d89e",
      outputSize: 85138,
      header: "XLSY;v2;",
    },
  };
  assert.deepEqual(lock.regressions, expectedRegressions);
  for (const [name, regression] of Object.entries(lock.regressions)) {
    assertDigest(regression.inputSha256, "sha256", `${name}.inputSha256`);
    assertDigest(regression.outputSha256, "sha256", `${name}.outputSha256`);
    assert.ok(Number.isSafeInteger(regression.outputSize) && regression.outputSize > 0);
    assert.ok(typeof regression.header === "string" && regression.header.length > 0);
  }

  const releaseDirectory = path.resolve(repositoryRoot, lock.deployment.directory);
  assert.equal(
    releaseDirectory,
    path.join(
      repositoryRoot,
      "public",
      "packages",
      "onlyoffice",
      "x2t",
      lock.tag,
    ),
    "x2t deployment directory must be immutable and tag-versioned",
  );

  const deployedLock = await readFile(path.join(releaseDirectory, "release-lock.json"));
  assert.ok(deployedLock.equals(lockBytes), "deployed x2t release lock differs from config lock");

  const script = await verifyFile(
    path.join(releaseDirectory, "x2t.js"),
    lock.files.script,
    "deployed x2t.js",
  );
  const compressedWasm = await verifyFile(
    path.join(releaseDirectory, "x2t.wasm"),
    lock.files.wasmBrotli,
    "deployed Brotli x2t.wasm",
  );
  const rawWasm = brotliDecompressSync(compressedWasm);
  assert.equal(rawWasm.byteLength, lock.files.wasmRaw.size, "raw x2t.wasm size changed");
  assert.equal(sha256(rawWasm), lock.files.wasmRaw.sha256, "raw x2t.wasm SHA-256 changed");
  assert.equal(sha512(rawWasm), lock.files.wasmRaw.sha512, "raw x2t.wasm SHA-512 changed");
  assert.match(script.toString("utf8", 0, 256), /Module|WebAssembly|emscripten/i);

  console.log("x2t release verification passed");
  console.log(`  source: ${lock.repository}@${lock.sourceCommit}`);
  console.log(`  release: ${lock.tag}`);
  console.log(`  x2t.js sha256: ${lock.files.script.sha256}`);
  console.log(`  x2t.js.br sha256: ${lock.files.scriptBrotli.sha256}`);
  console.log(`  x2t.wasm.br sha256: ${lock.files.wasmBrotli.sha256}`);
  console.log(`  x2t.wasm raw sha256: ${lock.files.wasmRaw.sha256}`);
}

main().catch((error) => {
  console.error("x2t release verification failed");
  console.error(error);
  process.exitCode = 1;
});
