import { open, rename, unlink, lstat } from "node:fs/promises";
import { dirname, join, basename, posix } from "node:path";
import { OoxmlError } from "../errors.ts";
import { parseXml, elements } from "../xml/index.ts";
import { readZip, writeZip, type ZipLimits } from "./zip.ts";

const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const encoder = new TextEncoder();
export interface Relationship { id: string; type: string; target: string; external: boolean; resolved?: string }
export interface PackageDiff { added: string[]; changed: string[]; removed: string[] }

/** Package custody: original member bytes survive until an explicit set/delete.
 * Reading returns copies; XML inspection must never mark a part dirty. Unknown
 * parts remain opaque. A no-op save returns the original archive, not a re-ZIP.
 */
export class OpcPackage {
  private parts: Map<string, Uint8Array>;
  private readonly baseline: Map<string, Uint8Array>;
  private readonly original?: Uint8Array;
  private constructor(parts: ReadonlyMap<string, Uint8Array>, original?: Uint8Array) {
    this.parts = cloneParts(parts);
    this.baseline = cloneParts(parts);
    this.original = original?.slice();
    this.validate();
  }

  static async open(input: string | Uint8Array, limits: ZipLimits = {}): Promise<OpcPackage> {
    if (typeof input === "string") {
      const file = Bun.file(input);
      if (file.size > (limits.maxArchiveBytes ?? 256 * 1024 * 1024)) throw new OoxmlError("zip-archive-too-large", "Input exceeds archive bytes limit");
      const bytes = new Uint8Array(await file.arrayBuffer());
      return new OpcPackage(readZip(bytes, limits), bytes);
    }
    return new OpcPackage(readZip(input, limits), input);
  }

  static fromParts(parts: ReadonlyMap<string, Uint8Array>): OpcPackage {
    // Run the same ZIP name validation as externally loaded packages.
    return new OpcPackage(readZip(writeZip(parts)));
  }

  names(): string[] { return [...this.parts.keys()].sort(); }
  get(name: string): Uint8Array | undefined { return this.parts.get(name)?.slice(); }
  text(name: string): string {
    const data = this.parts.get(name);
    if (!data) throw new OoxmlError("opc-part-missing", `Missing package part: ${name}`);
    return decodeXml(data);
  }
  set(name: string, value: Uint8Array | string): void {
    checkName(name);
    const collision = this.names().find(n => n !== name && asciiLower(n) === asciiLower(name));
    if (collision) throw new OoxmlError("zip-case-collision", `${name} collides with ${collision}`);
    this.parts.set(name, typeof value === "string" ? encoder.encode(value) : value.slice());
  }
  delete(name: string): void { checkName(name); this.parts.delete(name); }

  /** Relationships are resolved relative to their owner, never fetched. External
   * targets remain data; a library reader must not dereference URLs or files.
   */
  relationships(part = ""): Relationship[] {
    const path = relationshipPath(part);
    if (!this.parts.has(path)) return [];
    const xml = parseXml(this.text(path));
    if (xml.root.localName !== "Relationships" || xml.root.namespaceURI !== REL) throw new OoxmlError("opc-relationships-invalid", `Invalid relationships root: ${path}`);
    const ids = new Set<string>();
    return xml.root.children.map(node => {
      const { Id: id, Type: type, Target: target, TargetMode: mode } = node.attributes;
      if (node.localName !== "Relationship" || node.namespaceURI !== REL || !id || !type || !target || (mode !== undefined && mode !== "Internal" && mode !== "External")) throw new OoxmlError("opc-relationships-invalid", `Malformed relationship in ${path}`);
      if (ids.has(id)) throw new OoxmlError("opc-relationship-duplicate", `Duplicate relationship id ${id} in ${path}`);
      ids.add(id);
      const external = mode === "External";
      return { id, type, target, external, ...(external ? {} : { resolved: resolveTarget(part, target) }) };
    });
  }

  related(part: string, typeSuffix: string): string | undefined {
    const matches = this.relationships(part).filter(r => !r.external && r.type.endsWith("/" + typeSuffix));
    if (matches.length > 1) throw new OoxmlError("opc-relationship-ambiguous", `Several ${typeSuffix} relationships on ${part}`);
    return matches[0]?.resolved;
  }
  mainPart(): string {
    const main = this.related("", "officeDocument");
    if (!main || !this.parts.has(main)) throw new OoxmlError("opc-main-part-missing", "No internal officeDocument relationship");
    return main;
  }
  diff(): PackageDiff {
    return {
      added: this.names().filter(n => !this.baseline.has(n)),
      changed: this.names().filter(n => this.baseline.has(n) && !sameBytes(this.parts.get(n)!, this.baseline.get(n)!)),
      removed: [...this.baseline.keys()].filter(n => !this.parts.has(n)).sort(),
    };
  }
  /** Transactions restore every part when a synchronous edit refuses. Async
   * callbacks are refused: their delayed work cannot be rolled back reliably.
   */
  transaction<T>(operation: () => T): T {
    if (operation.constructor.name === "AsyncFunction") throw new OoxmlError("opc-async-transaction", "Transactions require synchronous edits");
    const before = cloneParts(this.parts);
    try {
      const result = operation();
      if (result && typeof (result as { then?: unknown }).then === "function") throw new OoxmlError("opc-async-transaction", "Transactions require synchronous edits");
      return result;
    } catch (error) { this.parts = before; throw error; }
  }
  toBytes(): Uint8Array {
    this.validate();
    const diff = this.diff();
    if (this.original && !diff.added.length && !diff.changed.length && !diff.removed.length) return this.original.slice();
    return writeZip(this.parts);
  }

  /** Build and validate before touching the destination. A sibling file permits
   * atomic rename on the same filesystem; fsync makes the archive durable first.
   * Symlink destinations refuse rather than changing an unexpected target.
   */
  async save(path: string): Promise<void> {
    const bytes = this.toBytes();
    try { if ((await lstat(path)).isSymbolicLink()) throw new OoxmlError("opc-symlink-destination", "Symlink save destinations are refused"); }
    catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
    const temp = join(dirname(path), `.${basename(path)}.${crypto.randomUUID()}.tmp`);
    const file = await open(temp, "wx", 0o600);
    try {
      await file.writeFile(bytes);
      await file.sync();
      await file.close();
      await rename(temp, path);
    } catch (error) {
      await file.close().catch(() => {});
      await unlink(temp).catch(() => {});
      throw error;
    }
  }

  private validate(): void {
    if (!this.parts.has("[Content_Types].xml")) throw new OoxmlError("opc-content-types-missing", "Missing [Content_Types].xml");
    const content = parseXml(this.text("[Content_Types].xml"));
    if (content.root.localName !== "Types" || content.root.namespaceURI !== CT) throw new OoxmlError("opc-content-types-invalid", "Invalid content types root");
    const overrides = new Map<string, string>(), defaults = new Map<string, string>();
    for (const node of content.root.children) {
      const type = node.attributes.ContentType;
      if (node.namespaceURI !== CT || !type) throw new OoxmlError("opc-content-types-invalid", "Malformed content type");
      if (node.localName === "Default") {
        const ext = node.attributes.Extension?.toLowerCase();
        if (!ext || defaults.has(ext)) throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing default extension");
        defaults.set(ext, type);
      } else if (node.localName === "Override") {
        const name = node.attributes.PartName;
        if (!name?.startsWith("/") || overrides.has(name.slice(1))) throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing part override");
        checkName(name.slice(1)); overrides.set(name.slice(1), type);
      } else throw new OoxmlError("opc-content-types-invalid", "Unknown content type element");
    }
    for (const name of this.parts.keys()) {
      checkName(name);
      if (name === "[Content_Types].xml") continue;
      const type = overrides.get(name) ?? defaults.get(name.split(".").at(-1)!.toLowerCase());
      if (!type) throw new OoxmlError("opc-content-type-missing", `No content type for ${name}`);
      if (name.endsWith(".rels")) {
        const owner = relationshipOwner(name);
        if (owner && !this.parts.has(owner)) throw new OoxmlError("opc-relationship-owner-missing", `Missing owner of ${name}`);
        for (const rel of this.relationships(owner)) if (!rel.external && !this.parts.has(rel.resolved!)) throw new OoxmlError("opc-relationship-target-missing", `Missing target ${rel.resolved} from ${name}`);
      }
    }
    this.mainPart();
  }
}

export function relationshipPath(part: string): string {
  return part ? posix.join(posix.dirname(part), "_rels", posix.basename(part) + ".rels") : "_rels/.rels";
}
function relationshipOwner(path: string): string {
  if (path === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(path);
  if (!match) throw new OoxmlError("opc-relationships-invalid", `Noncanonical relationship part: ${path}`);
  return (match[1] ?? "") + match[2];
}
function resolveTarget(part: string, target: string): string {
  if (/[\\\u0000-\u0020]/.test(target) || /^(?:[A-Za-z][\w+.-]*:|\/\/)/.test(target)) throw new OoxmlError("opc-target-invalid", `Invalid internal relationship target ${target}`);
  let path: string;
  try { path = decodeURIComponent(target.split("#")[0]!); } catch { throw new OoxmlError("opc-target-invalid", "Invalid relationship URI encoding"); }
  if (/%2f|%5c/i.test(target) || path.includes("?")) throw new OoxmlError("opc-target-invalid", "Encoded separators/query in internal relationship");
  path = path.startsWith("/") ? posix.normalize(path.slice(1)) : posix.normalize(posix.join(posix.dirname(part || "_"), path));
  checkName(path);
  return path;
}
function checkName(name: string): void {
  if (!name || name.startsWith("/") || /[\\\u0000-\u001f?#]/.test(name) || name.split("/").some(p => !p || p === "." || p === "..")) throw new OoxmlError("opc-part-name-invalid", `Noncanonical part name: ${name}`);
}
export function sameBytes(a: Uint8Array, b: Uint8Array): boolean { return a.length === b.length && a.every((v,i) => v === b[i]); }
function cloneParts(parts: ReadonlyMap<string, Uint8Array>): Map<string, Uint8Array> { return new Map([...parts].map(([n,b]) => [n,b.slice()])); }
function asciiLower(s: string): string { return s.replace(/[A-Z]/g, c => c.toLowerCase()); }
function decodeXml(bytes: Uint8Array): string {
  try {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le", { fatal: true }).decode(bytes);
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be", { fatal: true }).decode(bytes);
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch { throw new OoxmlError("opc-xml-encoding", "Invalid XML encoding"); }
}
