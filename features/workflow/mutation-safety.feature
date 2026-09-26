@implemented @bun
Feature: Shared OOXML mutation safety
  Acceptance adapters translate these document operations to native APIs or MCP tools.
  Planned scenarios are inventory only and cannot count as executed passes.

  @id-pptx-dry-run-no-mutation
  Scenario Outline: Previewing a slide edit with <destination> leaves files unchanged
    Given fixture "title-and-subtitle.pptx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes, destination bytes and document-directory entries are recorded
    When previewing this batch without committing:
      | target        | value_json             |
      | slide:1/title | "Changed by dry run"   |
    Then committed change count is 0
    And source bytes, destination bytes and document-directory entries equal the recorded state
    And the reopened source slide 1 title is "Original title"
    Examples:
      | destination       |
      | source            |
      | distinct-absent   |
      | distinct-existing |

  @id-docx-dry-run-no-mutation
  Scenario Outline: Previewing a Word replacement with <destination> leaves files unchanged
    Given fixture "present-placeholder.docx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes, destination bytes and document-directory entries are recorded
    When previewing this batch without committing:
      | target    | value_json |
      | <Present> | "changed"  |
    Then committed change count is 0
    And source bytes, destination bytes and document-directory entries equal the recorded state
    And the reopened source current body text is "<Present>"
    Examples:
      | destination       |
      | source            |
      | distinct-absent   |
      | distinct-existing |

  @id-xlsx-dry-run-no-mutation
  Scenario Outline: Previewing a cell edit with <destination> leaves files unchanged
    Given fixture "default-style.xlsx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes, destination bytes and document-directory entries are recorded
    When previewing this batch without committing:
      | target | value_json |
      | A1     | "after"    |
    Then committed change count is 0
    And source bytes, destination bytes and document-directory entries equal the recorded state
    And the reopened source active sheet cell A1 is "before"
    Examples:
      | destination       |
      | source            |
      | distinct-absent   |
      | distinct-existing |

  @id-xlsx-strict-batch-atomicity
  Scenario Outline: A missing worksheet prevents the strict batch with <destination>
    Given fixture "default-style.xlsx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes, destination bytes and document-directory entries are recorded
    When committing this batch with all-targets-required policy:
      | target     | value_json |
      | A1         | "after"    |
      | Missing!B1 | 123        |
    Then the operation is refused for unmatched target "Missing!B1"
    And committed change count is 0
    And source bytes, destination bytes and document-directory entries equal the recorded state
    And the reopened source active sheet cell A1 is "before"
    Examples:
      | destination       |
      | source            |
      | distinct-absent   |
      | distinct-existing |

  @id-docx-strict-batch-per-target-results
  Scenario Outline: A matched placeholder cannot conceal an unmatched placeholder with <destination>
    Given fixture "present-placeholder.docx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes, destination bytes and document-directory entries are recorded
    When committing this batch with all-targets-required policy:
      | target    | value_json       |
      | <Present> | "changed"        |
      | <Missing> | "never inserted" |
    Then the operation is refused for unmatched target "<Missing>"
    And "<Missing>" is never reported applied
    And committed change count is 0
    And source bytes, destination bytes and document-directory entries equal the recorded state
    And the reopened source current body text is "<Present>"
    Examples:
      | destination       |
      | source            |
      | distinct-absent   |
      | distinct-existing |

  @id-pptx-batch-output-accumulates
  Scenario Outline: The committed <destination> contains every slide edit
    Given fixture "title-and-subtitle.pptx" verified against the fixture manifest
    And destination state is "<destination>"
    And source bytes are recorded
    When committing this batch to the distinct destination:
      | target           | value_json         |
      | slide:1/title    | "Changed title"    |
      | slide:1/subtitle | "Changed subtitle" |
    Then committed change count is 2
    And source bytes equal the recorded state
    And the reopened destination slide 1 title is "Changed title"
    And the reopened destination slide 1 subtitle is "Changed subtitle"
    And all destination relationship and content-type references resolve
    And destination member payloads outside the manifest change allowance are byte-identical
    Examples:
      | destination       |
      | distinct-absent   |
      | distinct-existing |

  @id-xlsx-style-dependency-closure
  Scenario: A multiline cell edit includes its required style definition
    Given fixture "default-style.xlsx" verified against the fixture manifest
    And destination state is "distinct-absent"
    And source bytes are recorded
    When committing this batch with multiline wrap enabled:
      | target | value_json       |
      | A1     | "first\\nsecond" |
    Then committed change count is 1
    And source bytes equal the recorded state
    And the reopened destination active sheet cell A1 has value_json "first\nsecond"
    And that cell resolves to a style with wrapText enabled
    And every cell style index is below the output cellXfs count
    And all destination relationship and content-type references resolve
    And destination member payloads outside the manifest change allowance are byte-identical

  @id-xlsx-cross-sheet-cache-invalidation
  Scenario: An input edit invalidates a cached answer on another sheet
    Given fixture "cross-sheet-cache.xlsx" verified against the fixture manifest
    And destination state is "distinct-absent"
    And source bytes are recorded
    And calculation policy is "invalidate-without-recalculation"
    When committing this batch to the distinct destination:
      | target   | value_json |
      | Input!A1 | 10         |
    Then committed change count is 1
    And source bytes equal the recorded state
    And the reopened destination Input!A1 is numeric 10
    And the reopened destination Calc!A1 formula is "=Input!A1*2"
    And the destination Calc!A1 cached value is absent or empty
    And calculation state is "recalculation-required"
    And a data-only read never returns the old cached value 2 as current
    And all destination relationship and content-type references resolve
    And destination member payloads outside the manifest change allowance are byte-identical
