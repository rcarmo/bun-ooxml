@implemented @bun
Feature: Native XLSX workbook creation and missing-cell authoring
  Bun authors minimal native OOXML workbooks, adds worksheets safely, and extends
  supported worksheets with missing rows and cells while preserving unrelated XML.

  @id-xlsx-create-native-default
  Scenario: Create a minimal native workbook and reopen saved values
    Given a native-created XLSX workbook
    When I write "Alpha" to A1, 7 to B2, true to C3 and save and reopen it
    Then the reopened created workbook has Sheet1 values at A1, B2 and C3
    And the saved created workbook contains only the minimal workbook, worksheet and styles parts
    And Sheet1 dimension becomes A1:C3

  @id-xlsx-create-add-worksheet
  Scenario: Add worksheets with independent ids and reopen them
    Given a native-created XLSX workbook
    When I add worksheets "Data" and "Summary", write 5 to Data A1 and "done" to Summary B2, and save and reopen the workbook
    Then the reopened workbook sheetnames are Sheet1, Data and Summary
    And the saved workbook assigns independent worksheet relationship and sheet ids

  @id-xlsx-create-prefixed-missing-cell
  Scenario: Extend a prefixed worksheet with a missing cell without unqualified spreadsheet nodes
    Given a prefixed XLSX fixture with a missing B2 target
    When I write "Prefixed" to its missing B2 and save and reopen it
    Then the saved prefixed worksheet uses qualified dimension row cell and text nodes for B2
    And the reopened prefixed workbook reads B2 as "Prefixed"

  @id-xlsx-create-coordinate-boundary
  Scenario: Write the maximum worksheet coordinate
    Given a native-created XLSX workbook
    When I write 1 to XFD1048576 and save and reopen it
    Then the reopened created workbook reads XFD1048576 as 1
    And Sheet1 dimension becomes XFD1048576:XFD1048576

  @id-xlsx-create-row-ordering
  Scenario: Missing rows and cells are inserted in numeric order
    Given a native-created XLSX workbook
    When I write 3 to C3, 1 to A1, 2 to B1 and 4 to A2 and save and reopen it
    Then the saved created worksheet stores rows and cells in ascending numeric order
    And the reopened created workbook keeps those ordered values

  @id-xlsx-create-invalid-params-atomic
  Scenario: Invalid coordinates and sheet names refuse before mutation
    Given a native-created XLSX workbook
    When I try invalid worksheet coordinates and names
    Then every invalid create-side refusal leaves the workbook bytes unchanged
    And the workbook still has only Sheet1

  @id-xlsx-create-existing-fixture-append
  Scenario: Append a missing cell to an existing worksheet without disturbing unrelated parts
    Given the go-ooxml formatting workbook fixture for append
    When I append the missing cell E7 text to "tail" and save and reopen the workbook
    Then the reopened existing worksheet has the appended value at E7
    And the existing worksheet keeps its surrounding nodes and unrelated parts
