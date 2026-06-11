// Generates office-script/transform.ts from src/lib/transform-core.ts by
// stripping `export` keywords (Office Scripts forbid modules) and appending
// the main() entry point that Power Automate's "Run script" action calls.

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { resolve } from "node:path";

const root = process.cwd();
const sourcePath = resolve(root, "src", "lib", "transform-core.ts");
const outDir = resolve(root, "office-script");
const outPath = resolve(outDir, "transform.ts");

const core = readFileSync(sourcePath, "utf8").replace(/^export /gm, "");

const banner = `// GENERATED FILE - do not edit directly.
// Source: src/lib/transform-core.ts. Regenerate with: npm run build:script
//
// Deploy: open https://www.office.com -> Excel -> open NoteSync/runner.xlsx ->
// Automate tab -> New Script -> replace the editor contents with this entire
// file -> rename the script to "NoteSync Transform" -> Save script.

`;

const entryPoint = `
/**
 * Entry point invoked by Power Automate (Excel Online "Run script" action).
 * The workbook is unused; it exists only because Office Scripts require one.
 *
 * Parameter order must match the "Run script" action configuration:
 * vtt, subject, dateStr, notesJson, actionItemsJson.
 */
function main(
  workbook: ExcelScript.Workbook,
  vtt: string,
  subject: string,
  dateStr: string,
  notesJson: string,
  actionItemsJson: string
): string {
  return buildPageHtml(vtt, subject, dateStr, notesJson, actionItemsJson);
}
`;

mkdirSync(outDir, { recursive: true });
writeFileSync(outPath, banner + core + entryPoint, "utf8");
console.log(`Wrote ${outPath}`);
