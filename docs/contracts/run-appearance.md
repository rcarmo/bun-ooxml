# Direct run colour and appearance

`Paragraph.setRunFormatting()` accepts four scalar appearance fields in addition
to the [Boolean and font-size overrides](run-formatting.md):

| Field | Accepted values |
|---|---|
| `color` | Six RGB hex digits, optionally prefixed with `#`, or `auto` |
| `underline` | `none`, `single`, `double`, `thick`, `dotted`, `dash`, `wave` |
| `highlight` | The 17 `ST_HighlightColor` names, including `none` |
| `verticalAlign` | `baseline`, `superscript`, `subscript` |

RGB letter case is preserved; the optional hash is removed before writing.
Highlight names are case-sensitive: `black`, `blue`, `cyan`, `green`, `magenta`,
`red`, `yellow`, `white`, `darkBlue`, `darkCyan`, `darkGreen`, `darkMagenta`,
`darkRed`, `darkYellow`, `darkGray`, `lightGray`, `none`.

```ts
const paragraph = document.paragraphs[0]!;
paragraph.setRunFormatting({ color: '#336699', underline: 'single' });
const appearance = document.paragraphs[0]!.directRunAppearance();
// One detached object per supported direct run:
// { color: '336699', underline: 'single', highlight: null, verticalAlign: null }
```

`directRunAppearance()` reads all four direct properties in run order. Absence is
`null`; it differs from explicit `auto`, `none` and `baseline`. Setting a field to
null removes its element; undefined leaves it alone. The existing paragraph-wide
formatting rules still apply: a real change expires paragraph handles, while an
identical request retains the original archive bytes and handles.

## Refusals and preservation

Selected property elements must contain only a WordprocessingML `val` attribute
and optional namespace declarations. A themed colour or coloured/themed underline
refuses even on removal or an apparently identical assignment. The getter also
refuses these cases instead of returning a value that theme metadata can override.
Editing an unrelated property, such as bold, preserves those unselected elements
unchanged. Missing, malformed, wrong-namespace or out-of-profile scalar values
refuse when read or selected for editing.

All runs are checked before a package mutation. A late refusal or failed
serialization leaves package bytes and model handles intact. Tests separately
cover UTF-16 path save/reopen, table-grid preservation, stale source detection and
protected documents. Only selected property elements are rewritten; other run
properties and unrelated package payloads retain their bytes.

## Specification and limits

ECMA-376 Part 1, fifth edition (October 2016), defines run `color` in §17.3.2.6,
`highlight` in §17.3.2.15, underline `u` in §17.3.2.40 and `vertAlign` in
§17.3.2.42. Highlight names use §17.18.40; underline patterns use §17.18.99.
The complete PDF is in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).

The underline list is an editor subset. Theme colours, decorated underlines,
font selection, inherited defaults and rendering are outside this API. Vertical
alignment reports the direct enum, not a calculated text position. The shared
cases compare three colour values, six underline styles, five highlight values
and one pair of vertical-position getters in memory. Native tests separately
check saved values and custody. One generated sample passes Open XML SDK's
Office2019 schema profile, without visual or font-layout comparisons.
