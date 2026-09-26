@implemented @bun
Feature: OPC graph edits retain package custody
  Part and relationship edits validate before commit and never discard unrelated members.

  @id-opc-add-related-part
  Scenario: Add an opaque part with an exact content type and relationship
    Given the shared DOCX package for graph editing
    When an opaque part and internal root relationship are added
    Then the saved package reopens with the new part and its exact content type
    And unrelated original payloads retain their exact bytes

  @id-opc-graph-rollback
  Scenario: A referenced part cannot be removed alone
    Given the shared DOCX package with an additional related opaque part
    When removal of the referenced part is attempted
    Then graph editing refuses with code "opc-part-referenced"
    And the package bytes are unchanged after the refusal

  @id-opc-remove-related-part
  Scenario: Explicitly detach then remove an opaque part
    Given the shared DOCX package with an additional related opaque part
    When its unreferenced root relationship and part are removed
    Then the saved package no longer contains the part or its content-type override
    And all retained relationship targets resolve

  @id-opc-diff-content-type
  Scenario: A type-only change appears in the package diff
    Given the shared DOCX package for graph editing
    When only a part content type is changed
    Then the package diff identifies the part despite identical payload hashes
    And the diff reports no unrelated additions or removals
