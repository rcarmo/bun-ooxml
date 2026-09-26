# Conservative XML comparison

`xmlEquivalent(leftBytes, rightBytes)` compares supported XML 1.0 documents for
bounded package-preservation decisions. It returns a Boolean without changing
either input. It does not emit canonical XML and must not be used for signatures,
schema validity, arbitrary XML vocabularies or a general document-equivalence
claim.

```ts
import { xmlEquivalent } from 'bun-ooxml';

const equivalent = xmlEquivalent(
  new TextEncoder().encode('<a xmlns="urn:example"/>'),
  new TextEncoder().encode('<p:a xmlns:p="urn:example"></p:a>'),
);
```

Element and attribute names compare by namespace URI and local name. Attribute
order and unused namespace declarations are ignored. Decoded attribute values
and ordered character data remain exact after XML line-ending and attribute
normalisation. CDATA boundaries and entity spellings do not distinguish equal
character data. Child order, text around children, comments and processing
instruction targets, separator spacing and data remain significant, including markup before and after
the root. The standalone declaration is retained; byte encoding and an optional
XML 1.0 declaration do not distinguish an otherwise equal document.

Only a document-root OPC `Relationships` collection of empty `Relationship`
elements with nonempty unique IDs permits child reordering. Mixed text, comments,
other children and ordinary XML retain order. This comparison does not validate
the relationship schema, targets or graph. Content-type collections are not
reordered by this implementation.

Colon-bearing attribute values conservatively retain their full in-scope namespace
bindings. The same applies to QName-like element text, `xsi:type`,
`mc:Ignorable`, `mc:MustUnderstand` and `mc:Choice/@Requires`.
Unbound attribute-referenced prefixes refuse. This can reject equivalent prefix changes
or ordinary URL values. Unprefixed schema-defined QName text or other typed values cannot be identified
without that schema and are outside this API's semantic guarantee.

DTD-bearing, malformed, unsupported encoding/version and parser-limit inputs
return false even when the bytes are identical. UTF-8 and BOM-prefixed UTF-16LE/BE
are supported. The syntax-checked event path avoids building descendant text for
every ancestor; it still allocates bounded parse and comparison models rather
than streaming arbitrarily large inputs.

The ten shared comparison cases are bound directly. Additional tests check tail
text, comments, PI data, QName-sensitive bindings, encodings and depth refusal.
Python's comparator also reorders content-type children and differs in namespace
and malformed-collection policies. These tests establish their exact stated
outcomes, not cross-runtime equivalence for every XML input. Package semantic
diff remains a separate, unbound operation; `diffPackages()` keeps its existing
byte/content-type semantics.
