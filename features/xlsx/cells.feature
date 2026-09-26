@implemented @bun
Feature: XLSX cell reads and style-preserving value edits
  Workbook cells are read from the preserved OOXML package and simple value edits
  patch only the required XML slots while keeping untouched package bytes intact.

  @id-xlsx-read-rel-linked-shared-strings
  Scenario: Read rel-linked shared strings, numbers, booleans and cached formulas across worksheets
    Given a synthetic two-sheet XLSX fixture with rel-linked shared strings, numbers, booleans and cached formulas
    When I open the workbook through the XLSX reader
    Then the workbook sheetnames follow workbook relationships
    And cell reads include shared strings, numbers, booleans and cached formulas across worksheets

  @id-xlsx-preserve-styled-cell-edit
  Scenario: Edit a styled existing cell without disturbing unrelated package parts
    Given the go-ooxml formatting workbook fixture
    When I change the styled cell A2 text to "Elizabeth Lavenza" and save and reopen the workbook
    Then the reopened cell keeps its style attributes and new value
    And unrelated ZIP parts still match the original bytes

  @id-xlsx-clear-cross-sheet-caches
  Scenario: Editing an input cell clears formula caches across worksheets and enables recalculation
    Given a synthetic two-sheet XLSX fixture with cross-sheet cached formulas
    When I change the input cell A1 to 10 and save the workbook
    Then every formula cache is removed across the worksheets
    And workbook calculation flags request full recalculation

  @id-xlsx-refuse-shared-formula-overwrite
  Scenario: Refuse destructive edits to an unsupported shared formula without mutating the workbook
    Given the shared-formula XLSX fixture
    When I try to overwrite the shared formula cell B2
    Then the shared formula edit is refused before mutation
    And saving afterwards keeps the workbook and worksheet parts byte-identical

  @id-xlsx-refuse-array-formula-overwrite
  Scenario: Refuse destructive edits to an unsupported array formula without mutating the workbook
    Given the shared-formula XLSX fixture
    When I try to overwrite the array formula cell D2
    Then the array formula edit is refused before mutation
    And saving afterwards keeps the workbook and worksheet parts byte-identical
