@planned
Feature: Cross-language Office workflow regression outcomes
  Derived from isolated Python MCP audit reproductions. These are shared outcome
  contracts, not implementation choices or passing claims for unimplemented APIs.

  @id-pptx-dry-run-preserves-source
  Scenario: Dry runs leave original files untouched
    Given an existing presentation and its original SHA256
    When a text patch is requested with dry_run enabled
    Then the preview describes the requested edit
    And the original file SHA256 is unchanged

  @id-xlsx-strict-batch-atomic
  Scenario: A missing target refuses an entire strict batch
    Given a workbook and a strict batch with one matching and one missing target
    When the batch is applied
    Then the result identifies the missing target
    And no matched edit is saved and the source SHA256 is unchanged

  @id-pptx-batch-accumulates-edits
  Scenario: Safe output contains every successful batch edit
    Given a presentation with separate title and subtitle targets
    When both edits are saved to a new output path in one batch
    Then reopening the output yields both changed strings
    And the source file is unchanged

  @id-docx-match-count-per-target
  Scenario: Match results belong to each target
    Given a document containing one requested placeholder and missing another
    When both placeholders are inspected for replacement
    Then the existing target reports one match
    And the missing target reports zero matches

  @id-xlsx-style-dependency-preservation
  Scenario: A new cell style carries its dependency into the output
    Given a worksheet receiving a new wrapped text style
    When the workbook is saved through preservation mode
    Then every cell style index is valid in the saved style table
    And an independent reader reopens the edited cell without an index error

  @id-xlsx-cross-sheet-cache-freshness
  Scenario: Input edits invalidate dependent caches on untouched sheets
    Given an input sheet and a formula cache on a different sheet
    When the input is edited and the workbook is saved
    Then the dependent cached value is invalidated or recomputed correctly
    And a data-only reader never returns the old cached answer as current
