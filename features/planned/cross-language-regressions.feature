@planned
Feature: Office edits report the state of the saved document accurately
  The source file remains intact until a validated mutation commits.
  These scenarios are proposed acceptance contracts, not implemented tests.

  @id-office-pptx-dry-run
  Scenario: Previewing a PowerPoint title change writes nothing
    Given a presentation with the title "Original title"
    And the source file hash and directory contents have been recorded
    When office_patch previews the title "Changed by dry run" in dry_run mode
    Then the reported number of committed changes is 0
    And the source file hash and directory contents are unchanged
    And the reopened presentation title is "Original title"

  @id-office-xlsx-strict-rollback
  Scenario: An unmatched Excel target prevents the entire strict batch
    Given a workbook whose active sheet cell A1 contains "before"
    And the workbook has no worksheet named "Missing"
    And the source file hash has been recorded
    When office_patch runs in strict mode with these changes:
      | target     | value |
      | A1         | after |
      | Missing!B1 | 123   |
    Then the result identifies "Missing!B1" as unmatched
    And the reported number of committed changes is 0
    And the source file hash is unchanged
    And the reopened active sheet cell A1 contains "before"

  @id-office-pptx-batch-output
  Scenario: A distinct output contains every successful PowerPoint edit
    Given a presentation with the title "Original title" and subtitle "Original subtitle"
    And a distinct output path has been selected
    When office_patch runs in safe mode with these changes:
      | target           | value            |
      | slide:1/title    | Changed title    |
      | slide:1/subtitle | Changed subtitle |
    Then the source file hash is unchanged
    And the output presentation title is "Changed title"
    And the output presentation subtitle is "Changed subtitle"
    And the reported number of committed changes is 2

  @id-office-docx-per-target-results
  Scenario: One matched Word placeholder cannot conceal another missing placeholder
    Given a Word document containing "<Present>" but not "<Missing>"
    When office_patch runs in strict mode with these changes:
      | target    | value          |
      | <Present> | changed        |
      | <Missing> | never inserted |
    Then the result identifies "<Missing>" as unmatched
    And the result does not report "<Missing>" as applied
    And the source file hash is unchanged

  @id-office-xlsx-style-closure
  Scenario: A multiline cell edit includes its required style definition
    Given a workbook whose only cell style is the default style
    And cell A1 contains "before"
    When office_patch writes a two-line string to A1 using a distinct output path
    Then the output contains the requested two-line string
    And every worksheet cell style index resolves in the output style table
    And the output reopens successfully
    And unrelated package member payloads are unchanged

  @id-office-xlsx-dependent-cache
  Scenario: An edited input cannot leave a stale cached answer on another sheet
    Given Input!A1 contains 1
    And Calc!A1 has the formula "=Input!A1*2" and cached value 2
    When office_patch changes Input!A1 to 10 without a calculation engine
    Then Calc!A1 retains the formula "=Input!A1*2"
    And Calc!A1 has no cached value
    And the receipt reports that recalculation is required
    And office_read does not report the old cached value 2 as current
