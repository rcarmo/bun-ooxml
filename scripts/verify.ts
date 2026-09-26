import { join } from "node:path";
import { inventoryFeatures } from "./gherkin.ts";
const root = join(import.meta.dir, "..");
const full = Bun.argv.includes("--full");
const manifest = await Bun.file(join(root,"references/manifest.json")).json();
const ledger = await Bun.file(join(root,"docs/contracts/parity-ledger.json")).json();
const errors: string[] = [];
if (manifest.sources.length !== 5 || manifest.files.length === 0) errors.push("Missing source inventory");
const seen = new Set<string>();
const corpus = new Map<string, number>();
for (const row of manifest.files) {
  if (seen.has(row.path) || row.path.startsWith("/") || row.path.includes("..")) { errors.push(`Unsafe/duplicate path ${row.path}`); continue; }
  seen.add(row.path);
  const file = Bun.file(join(root,row.path));
  if (!await file.exists()) { errors.push(`Missing ${row.path}`); continue; }
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.length !== row.bytes || new Bun.CryptoHasher("sha256").update(bytes).digest("hex") !== row.sha256) errors.push(`Source drift: ${row.path}`);
  if (row.role === "fixture" && /\.(docx|pptx|xlsx)$/.test(row.path)) corpus.set(row.source,(corpus.get(row.source) ?? 0)+1);
}
for (const id of ["go-ooxml","python-office-mcp-server"]) if (corpus.get(id)!==37) errors.push(`Expected 37 OOXML fixtures for ${id}, got ${corpus.get(id)}`);
for (const directory of ["upstream", "fixtures"]) {
  for await (const path of new Bun.Glob("**/*").scan({cwd: join(root,directory),onlyFiles:true})) {
    const relative = `${directory}/${path}`;
    if (relative !== "references/manifest.json" && !seen.has(relative)) errors.push(`Unpinned frozen file: ${relative}`);
  }
}
const sourceRows = manifest.files.filter((r: any) => r.role === "source" || r.role === "test");
const ledgerMap = new Map<string, any>();
for (const row of ledger.rows) {
  if (ledgerMap.has(row.path)) errors.push(`Duplicate ledger row ${row.path}`);
  ledgerMap.set(row.path,row);
  if (!["gap","partial","mapped","unsupported"].includes(row.state)) errors.push(`Invalid ledger state ${row.path}`);
  if (row.state === "mapped" && !row.mappings.length) errors.push(`Mapping lacks evidence ${row.path}`);
  // Mapping credit requires reviewed step/assertion evidence. Initial inventory has none.
  if (row.state === "mapped" && ledger.formalMapping !== true) errors.push(`Premature mapping claim ${row.path}`);
}
if (ledger.rows.length !== sourceRows.length) errors.push("behaviour ledger denominator drift");
for (const row of sourceRows) {
  const mapped = ledgerMap.get(row.path);
  if (!mapped || mapped.sha256 !== row.sha256 || mapped.role !== row.role) errors.push(`Missing/stale ledger row ${row.path}`);
}
// Bun's node:* modules are built into Bun; foreign runtime/process/FFI bridges are forbidden.
for await (const file of new Bun.Glob("**/*.ts").scan({cwd:join(root,"src")})) {
  const source = await Bun.file(join(root,"src",file)).text();
  if (/\bBun\.(?:spawn|spawnSync|dlopen)\s*\(|(?:from\s*|import\s*\()["'](?:node:)?(?:child_process|ffi)["']/.test(source)) errors.push(`Foreign runtime bridge: ${file}`);
}
const pkg = await Bun.file(join(root,"package.json")).json();
if (Object.keys(pkg.dependencies ?? {}).length) errors.push("Runtime dependencies require explicit native-runtime review");
const inventory = await inventoryFeatures(root);
const gaps = ledger.rows.filter((r:any)=>r.state!=="mapped").length;
if (full && (gaps || !ledger.formalMapping)) errors.push(`Full parity incomplete: ${gaps} source/test gaps`);
if (errors.length) throw new Error(errors.join("\n"));
console.log(`Verified ${manifest.files.length} frozen files; 74 rcarmo OOXML fixtures; ${sourceRows.length} source/test rows (${gaps} gaps); ${inventory.counts.scenarios.total} scenarios.`);
