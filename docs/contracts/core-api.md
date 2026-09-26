# Shared API for parallel implementation

`src/errors.ts`: `OoxmlError` extends Error, readonly `code: string`, constructor
`(code: string, message: string)`. Errors use stable codes.

`src/opc/zip.ts` (sync, Bun compression only):
* `type ZipLimits = { maxEntries?: number; maxEntryBytes?: number; maxTotalBytes?: number; maxCompressionRatio?: number; maxArchiveBytes?: number }`
* `readZip(bytes: Uint8Array, limits?: ZipLimits): Map<string, Uint8Array>`
* `writeZip(parts: ReadonlyMap<string, Uint8Array>, options?: { forceZip64?: boolean }): Uint8Array`
* `crc32(bytes: Uint8Array): number`

`src/opc/package.ts`:
* `OpcPackage.open(input: string | Uint8Array, limits?: ZipLimits): Promise<OpcPackage>`
* `OpcPackage.fromParts(parts: ReadonlyMap<string, Uint8Array>): OpcPackage`
* `names(): string[]`, `get(name): Uint8Array | undefined`, `text(name): string`
* `set(name, value: Uint8Array | string): void`, `delete(name): void`
* `relationships(part?: string): Relationship[]` default root
* `related(part: string, typeSuffix: string): string | undefined`
* `mainPart(): string`
* `diff(): { added: string[]; changed: string[]; removed: string[] }`
* `toBytes(): Uint8Array`, `save(path: string): Promise<void>`
* `transaction<T>(operation: () => T): T` synchronous rollback on throw
* `Relationship = { id: string; type: string; target: string; external: boolean; resolved?: string }`

`src/xml/index.ts` (strict XML, preservation via source offsets):
* `parseXml(text: string): XmlDocument`
* `XmlDocument.elements: XmlElement[]` preorder; `root: XmlElement`
* `XmlElement`: `name`, `localName`, `namespaceURI`, `attributes: Record<string,string>`,
  `children: XmlElement[]`, `parent?: XmlElement`, `text: string` decoded aggregate,
  `start`, `openEnd`, `closeStart`, `end` UTF-16 offsets; `selfClosing: boolean`
* `elements(xml: XmlDocument | XmlElement, localName: string, namespaceURI?: string): XmlElement[]`
* `escapeText(value: string): string`, `escapeAttribute(value: string): string`
* `applyEdits(xml: string, edits: {start: number; end: number; value: string}[]): string`
  refuses invalid/overlapping offsets before returning any changed text.

All getters return detached bytes so caller mutation cannot bypass dirty tracking.
No-op package serialization returns the original whole archive bytes.
