@implemented @bun
Feature: Strict XML parsing and offset-safe edits
  XML parts are parsed natively with preserved UTF-16 source offsets so OPC code
  can inspect and patch OOXML parts without losing untouched bytes.

  @id-xml-parse-offsets
  Scenario: Parse namespaces, mixed content and preserved offsets
    Given an XML document with a declaration, comments, processing instructions and namespaces
    When the document is parsed
    Then the root element and descendants expose decoded text, decoded attributes and namespace URIs
    And each element exposes UTF-16 source offsets, parent links, child links, root links and self-closing state

  @id-xml-normalise-line-endings
  Scenario: Decode XML line endings without changing source offsets
    Given XML text and attributes containing raw CRLF and character references
    When that XML is parsed without rewriting the source
    Then decoded text normalises raw line endings but preserves referenced carriage returns
    And decoded attributes normalise literal whitespace while preserving referenced whitespace
    And element offsets still address the original source string

  @id-xml-parse-refusals
  Scenario: Refuse malformed or unsafe XML constructs
    Given XML containing a malformed declaration or processing instruction
    And XML containing invalid comment termination or missing attribute whitespace
    And XML containing reserved namespace misuse, a DTD or an undeclared entity
    And XML containing a duplicate attribute, an unbound prefix or a mismatched tag
    When the document is parsed
    Then parsing is refused with a stable XML error code

  @id-xml-parse-bounds
  Scenario: Bound untrusted XML resources
    Given XML whose nesting depth, node count or input length exceeds the configured parser limits
    When the document is parsed
    Then parsing is refused before returning a partial tree

  @id-xml-apply-edits
  Scenario: Apply only disjoint edits that preserve full-document safety
    Given a well-formed XML document and source offsets for text or element content
    When disjoint edits are applied with escaped replacement text or XML fragments
    Then the resulting XML stays well formed and DTD free
    But overlapping edits or edits that leave malformed or DTD-bearing XML are refused before returning changed text
