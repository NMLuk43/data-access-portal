import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");
const inputPath = path.join(rootDir, "src", "app.js");
const outputPath = path.join(rootDir, "dist", "app.js");

const source = await readFile(inputPath, "utf8");

await mkdir(path.dirname(outputPath), { recursive: true });

const banner = "/* Built from src/app.js */\n";
await writeFile(outputPath, `${banner}${source}`, "utf8");

console.log(`Built ${path.relative(rootDir, outputPath)}`);
