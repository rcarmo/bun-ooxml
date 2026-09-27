# Append a plain Word run

`Paragraph.appendRun(text, formatting?)` appends one independent plain-text run
and returns the fresh paragraph handle. It uses the existing
[direct run formatting](run-formatting.md) patch, including font names, sizes and
bounded appearance fields.

```ts
let paragraph = document.addParagraph('');
paragraph = paragraph.appendRun('Bold ', { bold: true });
paragraph = paragraph.appendRun('Italic ', { italic: true });
paragraph = paragraph.appendRun('Colored', {
  color: 'FF0000', fontSizePt: 14, fontName: 'Arial',
});
// paragraph.text is 'Bold Italic Colored'
```

Appending never borrows direct properties from the preceding run. Missing or null
formatting fields leave those properties absent on the new run. Document and
paragraph styles can still affect its eventual appearance; this operation does
not compute inheritance.

Each successful call adds a run, including an empty-string call. It expires held
paragraph, span and table-cell snapshots. Table handles remain usable because the
grid is unchanged; reacquire their cells. The returned paragraph can be used for
the next append. This API does not expose a mutable run handle or split/reformat
an existing substring.

Text is limited to 1 Mi UTF-16 code units and cannot contain tabs or line breaks.
Invalid XML characters refuse. Spaces at either edge receive `xml:space="preserve"`;
metacharacters are escaped. These are editor limits. Broader Word run content,
such as break and tab elements, is not authored here.

## Preservation and refusals

The target must be a fresh, supported paragraph of plain direct runs with ordered
paragraph properties. Fields, tracked content, unsupported children, revisions,
lexical barriers and protected documents refuse. The current package's document
XML must match the held snapshot. New formatting passes the same type, namespace,
metadata and effect-conflict checks as `setRunFormatting()`.

Existing paragraph properties, runs, sibling paragraphs and unrelated package
parts retain their bytes. Self-closing paragraphs expand while retaining their
opening attributes and namespace identity. The prospective document is parsed,
all paragraph text is checked against the single intended append, and table
dimensions are checked before committing. Package write and serialization run
inside a transaction; failure preserves bytes and prior handles. Existing UTF-16
encoding survives serialization.

ECMA-376 Part 1, fifth edition (October 2016), §17.3.2.25 defines `r`; §17.3.3.31
defines `t` and its whitespace handling. The complete PDF is in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
Two shared scenarios check three-run concatenation and selected formatting after
path save/reopen. Native tests also check custody, stale handles, table paragraphs
and rollback. One sample passes Open XML SDK's Office2019 schema profile; visual
rendering and general run-editing parity have not been tested.
