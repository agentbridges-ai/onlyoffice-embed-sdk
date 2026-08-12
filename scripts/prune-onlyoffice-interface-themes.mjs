import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const runtimeRoot = path.resolve(
  process.argv[2] || "public/packages/onlyoffice/9.4.0-develop",
);
const editors = [
  "documenteditor",
  "spreadsheeteditor",
  "presentationeditor",
];
const forbiddenThemeClass =
  /(^|[^a-z0-9_-])\.theme-(?:system|light|classic-light|dark|contrast-dark)(?![a-z0-9_-])/i;
const forbiddenThemeIds = [
  "theme-system",
  "theme-light",
  "theme-classic-light",
  "theme-dark",
  "theme-contrast-dark",
];
const modernThemeRegistry =
  'var s={"theme-white":{text:e.txtThemeModernLight||"White",type:"light",source:"static",icons:{cls:"mod2"}},"theme-night":{text:e.txtThemeModernDark||"Night",type:"dark",source:"static",icons:{cls:"mod2"}}}';

function splitSelectorList(selector) {
  const selectors = [];
  let start = 0;
  let depth = 0;
  let quote = null;
  let escaped = false;

  for (let index = 0; index < selector.length; index += 1) {
    const character = selector[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (quote) {
      if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === "(" || character === "[") depth += 1;
    if (character === ")" || character === "]") depth -= 1;
    if (character === "," && depth === 0) {
      selectors.push(selector.slice(start, index));
      start = index + 1;
    }
  }
  selectors.push(selector.slice(start));
  return selectors;
}

function findStructuralCharacter(source, start, target) {
  let quote = null;
  let escaped = false;
  let comment = false;
  let parentheses = 0;
  let brackets = 0;

  for (let index = start; index < source.length; index += 1) {
    const character = source[index];
    const next = source[index + 1];
    if (comment) {
      if (character === "*" && next === "/") {
        comment = false;
        index += 1;
      }
      continue;
    }
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (quote) {
      if (character === quote) quote = null;
      continue;
    }
    if (character === "/" && next === "*") {
      comment = true;
      index += 1;
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === "(") parentheses += 1;
    if (character === ")") parentheses -= 1;
    if (character === "[") brackets += 1;
    if (character === "]") brackets -= 1;
    if (parentheses === 0 && brackets === 0 && character === target) {
      return index;
    }
  }
  return -1;
}

function findMatchingBrace(source, openBrace) {
  let depth = 1;
  let quote = null;
  let escaped = false;
  let comment = false;

  for (let index = openBrace + 1; index < source.length; index += 1) {
    const character = source[index];
    const next = source[index + 1];
    if (comment) {
      if (character === "*" && next === "/") {
        comment = false;
        index += 1;
      }
      continue;
    }
    if (escaped) {
      escaped = false;
      continue;
    }
    if (character === "\\") {
      escaped = true;
      continue;
    }
    if (quote) {
      if (character === quote) quote = null;
      continue;
    }
    if (character === "/" && next === "*") {
      comment = true;
      index += 1;
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === "{") depth += 1;
    if (character === "}") depth -= 1;
    if (depth === 0) return index;
  }
  throw new Error("Unterminated CSS block");
}

const containerAtRule =
  /^\s*@(media|supports|layer|container|document|scope|starting-style)\b/i;

function pruneCssBlocks(source, stats) {
  let cursor = 0;
  let result = "";

  while (cursor < source.length) {
    const openBrace = findStructuralCharacter(source, cursor, "{");
    if (openBrace < 0) {
      result += source.slice(cursor);
      break;
    }
    const closeBrace = findMatchingBrace(source, openBrace);
    const prelude = source.slice(cursor, openBrace);
    const body = source.slice(openBrace + 1, closeBrace);

    if (prelude.trimStart().startsWith("@")) {
      result += prelude;
      result += "{";
      result += containerAtRule.test(prelude)
        ? pruneCssBlocks(body, stats)
        : body;
      result += "}";
    } else {
      const selectors = splitSelectorList(prelude);
      const retained = selectors.filter(
        (selector) => !forbiddenThemeClass.test(selector),
      );
      stats.removedSelectors += selectors.length - retained.length;
      if (retained.length === 0) {
        stats.removedRules += 1;
      } else {
        result += retained.join(",");
        result += "{";
        result += body;
        result += "}";
      }
    }
    cursor = closeBrace + 1;
  }
  return result;
}

function pruneThemeRegistry(source, filePath) {
  const moduleStart = source.indexOf(
    'define("common/main/lib/controller/Themes"',
  );
  if (moduleStart < 0) {
    throw new Error(`Theme controller not found in ${filePath}`);
  }
  const registryStart = source.indexOf('var s={"theme-system":', moduleStart);
  const modernRegistryStart = source.indexOf(modernThemeRegistry, moduleStart);
  if (registryStart < 0 && modernRegistryStart >= 0) {
    return source;
  }
  const registryEnd = source.indexOf("},a={},r=function", registryStart);
  if (registryStart < 0 || registryEnd < 0) {
    throw new Error(`Static theme registry shape changed in ${filePath}`);
  }

  const result =
    source.slice(0, registryStart) +
    modernThemeRegistry +
    source.slice(registryEnd + 1);
  const controllerEnd = result.indexOf(
    'define("common/main/lib/util/utils"',
    moduleStart,
  );
  const controller = result.slice(
    moduleStart,
    controllerEnd < 0 ? undefined : controllerEnd,
  );
  for (const id of forbiddenThemeIds) {
    if (controller.includes(`"${id}":`)) {
      throw new Error(`Forbidden theme ${id} remains registered in ${filePath}`);
    }
  }
  for (const id of ["theme-white", "theme-night"]) {
    if (!controller.includes(`"${id}":`)) {
      throw new Error(`Required theme ${id} is missing from ${filePath}`);
    }
  }
  return result;
}

function pruneThemeCss(source, filePath) {
  const stats = { removedSelectors: 0, removedRules: 0 };
  const result = pruneCssBlocks(source, stats);
  if (forbiddenThemeClass.test(result)) {
    throw new Error(`Forbidden theme selector remains in ${filePath}`);
  }
  const requiredModernThemeRules = [
    [":root .theme-white{", "--toolbar-header-document:#f3f3f3"],
    [":root .theme-night{", "--toolbar-header-document:#222222"],
  ];
  for (const [selector, requiredVariable] of requiredModernThemeRules) {
    const selectorStart = result.indexOf(selector);
    const ruleEnd = result.indexOf("}", selectorStart);
    if (
      selectorStart < 0 ||
      ruleEnd < selectorStart ||
      !result.slice(selectorStart, ruleEnd).includes(requiredVariable)
    ) {
      throw new Error(
        `Required modern theme rule ${selector} is incomplete in ${filePath}`,
      );
    }
  }
  return { result, ...stats };
}

let bytesBefore = 0;
let bytesAfter = 0;
let totalRemovedSelectors = 0;
let totalRemovedRules = 0;

for (const editor of editors) {
  const appPath = path.join(
    runtimeRoot,
    "web-apps/apps",
    editor,
    "main/app.js",
  );
  const cssPath = path.join(
    runtimeRoot,
    "web-apps/apps",
    editor,
    "main/resources/css/app.css",
  );
  const appSource = await readFile(appPath, "utf8");
  const cssSource = await readFile(cssPath, "utf8");
  const appResult = pruneThemeRegistry(appSource, appPath);
  const cssResult = pruneThemeCss(cssSource, cssPath);

  bytesBefore += Buffer.byteLength(appSource) + Buffer.byteLength(cssSource);
  bytesAfter += Buffer.byteLength(appResult) + Buffer.byteLength(cssResult.result);
  totalRemovedSelectors += cssResult.removedSelectors;
  totalRemovedRules += cssResult.removedRules;
  await writeFile(appPath, appResult);
  await writeFile(cssPath, cssResult.result);
}

console.log("ONLYOFFICE interface theme pruning passed");
console.log(`  removed selectors: ${totalRemovedSelectors}`);
console.log(`  removed rules: ${totalRemovedRules}`);
console.log(`  removed bytes: ${bytesBefore - bytesAfter}`);
