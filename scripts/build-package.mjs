import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const sourceCssPath = join(rootDir, "app", "globals.css");
const outputCssPath = join(rootDir, "dist", "styles.css");
const buildInfoPath = join(rootDir, "dist", ".tsbuildinfo");

const sourceCss = await readFile(sourceCssPath, "utf8");

const packageCss = sourceCss
  .replace(/^@import "tailwindcss";\n\n/, "")
  .replace(/@theme inline \{[\s\S]*?\}\n\n/, "");

await mkdir(dirname(outputCssPath), { recursive: true });
await writeFile(outputCssPath, packageCss);
await rm(buildInfoPath, { force: true });
