@planned
Feature: Full source-pinned Office library ports
  The behaviour ledger is the full denominator. These contracts describe the
  remaining families; they receive no credit until bound to executed outcomes.

  @id-parity-inherited
  Scenario Outline: Preserve inherited <format> library behaviour
    Given the pinned <format> source modules and inherited tests
    When every public API and observable test case is mapped to TypeScript
    Then each mapping has executed Gherkin steps and saved reopened outcomes
    And no source or test ledger gaps remain for <format>
    Examples:
      | format |
      | docx   |
      | pptx   |
      | xlsx   |

  @id-parity-docx-review
  Scenario: Native Word reviews survive accept and reject
    Given documents with supported tracked changes and comment threads
    When the native compare and review APIs edit them
    Then accepting yields the revised document and rejecting yields the original
    And thread anchors and unrelated parts survive reopening

  @id-parity-pptx-composition
  Scenario: Slide imports preserve the chosen theme and relationship policy
    Given two presentations with shared media charts workbooks and notes
    When a slide is imported under an explicit reconciliation policy
    Then only the imported slide owns its independently editable chart workbook
    And the import report identifies each touched part and appearance change

  @id-parity-xlsx-structure
  Scenario: Structural spreadsheet edits rewrite all supported references
    Given formulas names charts tables and validation ranges across sheets
    When rows columns sheets or ranges are changed
    Then dependent references and address remaps agree after reopening
    And an unsupported reference refuses the entire edit without mutation

  @id-parity-native-oracle
  Scenario: Native calculation replaces the external Office subprocess
    Given a frozen workbook and independent calculated reference outputs
    When the Bun-native oracle evaluates the supported calculation contract
    Then values errors and exclusions match the pinned comparison corpus
    And no foreign runtime or Office subprocess executes

