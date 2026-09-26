@implemented @bun
Feature: DOCX native rectangular tables
  Native DOCX table authoring and rectangular cell text updates use direct XML,
  preserve unrelated package parts, and refuse unsafe table structures.

  @id-docx-table-create-roundtrip
  Scenario: DOCX table authoring round-trips escaped text, boundary spaces, and normalized newlines
    Given DOCX table source "new-document" is prepared
    When DOCX table 2x2 is appended
    And DOCX table 1 cell (0,0) text is set to "  <Alpha & Beta>  "
    And DOCX table 1 cell (1,1) text is set to "line 1\r\nline 2"
    And DOCX table document is saved and reopened
    Then DOCX table 1 has 2 rows and 2 columns
    And DOCX table 1 cell (0,0) text equals "  <Alpha & Beta>  "
    And DOCX table 1 cell (1,1) text equals "line 1\nline 2"
    And DOCX table 1 is stored before the section properties
    And DOCX table 1 cell (0,0) XML preserves boundary spaces and escapes special characters

  @id-docx-table-opaque-preserve
  Scenario: DOCX table cell updates preserve opaque package parts after save and reopen
    Given DOCX table source "fixtures/python-office-mcp-server/tests/_templates/testdata/word/simple_table.docx" is prepared
    And DOCX table opaque part "docProps/core.xml" bytes are remembered
    When DOCX table 1 cell (1,0) text is set to "Voltaic battery"
    And DOCX table document is saved and reopened
    Then DOCX table 1 has 4 rows and 3 columns
    And DOCX table 1 cell (1,0) text equals "Voltaic battery"
    And DOCX table opaque part "docProps/core.xml" bytes are unchanged

  @id-docx-table-stale-cell
  Scenario: DOCX table cell handles go stale after any document edit and refuse atomically
    Given DOCX table source "new-document" is prepared
    And DOCX table 1x1 is appended
    And DOCX table 1 cell (0,0) is remembered
    When DOCX table 1x1 is appended
    And DOCX table current saved bytes are remembered
    And DOCX table stale text set to "stale write" is attempted on the remembered cell
    Then DOCX table refusal code equals "docx-stale-table-cell"
    And DOCX table saved bytes equal the remembered bytes

  @id-docx-table-atomic-refusals
  Scenario Outline: DOCX table refuses <case> atomically
    Given DOCX table source "<source>" is prepared
    And DOCX table current saved bytes are remembered
    When DOCX table refusal "<case>" is attempted
    Then DOCX table refusal code equals "<code>"
    And DOCX table saved bytes equal the remembered bytes

    Examples:
      | source                                                                                         | case        | code                         |
      | fixtures/python-office-mcp-server/tests/_templates/testdata/word/simple_table.docx            | row-oob     | range                        |
      | references/fixtures-ooxml/reference-assets/docx/tests/paper/fixtures/generated/feature-isolated/table-merged-nested.docx  | merged-cell | docx-table-merged-cell       |
      | references/fixtures-ooxml/reference-assets/docx/tests/paper/fixtures/generated/feature-isolated/table-merged-nested.docx  | nested-cell | docx-table-cell-unsupported  |
      | synthetic-grid-before                                                                          | bizarre     | docx-table-unsupported       |
