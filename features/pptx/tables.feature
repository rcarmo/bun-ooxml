@implemented @bun
Feature: PPTX native rectangular tables
  Native PPTX table authoring preserves geometry, conservative cell formatting,
  and refusal atomicity for stale and unsupported table topologies.

  @id-pptx-table-roundtrip-geometry
  Scenario: Save and reopen an authored table while keeping exact geometry sums
    Given PPTX table round-trip scenario is prepared from a new presentation
    When PPTX authors a rectangular table and saves then reopens it
    Then PPTX preserves the table size, exact grid sums, and cell text after reopen

  @id-pptx-table-formatting
  Scenario: Preserve simple table cell formatting across update and reopen
    Given PPTX styled table fixture is prepared
    When PPTX replaces a styled table cell and saves then reopens the presentation
    Then PPTX preserves table cell formatting across the update and reopen

  @id-pptx-table-stale-handle
  Scenario: Refuse stale table cell handles atomically after slide mutation
    Given PPTX stale table handle scenario is prepared from a new presentation
    When PPTX mutates the slide and retries a stale table cell handle
    Then PPTX refuses the stale table handle without mutating the package

  @id-pptx-table-atomic-refusals
  Scenario Outline: Refuse unsupported table topology <case> atomically
    Given PPTX table refusal scenario "<case>" is prepared
    When PPTX attempts the table refusal "<case>"
    Then PPTX refusal "<code>" is returned atomically for table refusal "<case>"

    Examples:
      | case            | code                           |
      | merged-cell     | PPTX_TABLE_MERGE_UNSUPPORTED   |
      | malformed-merge | PPTX_TABLE_STRUCTURE_UNSUPPORTED |
