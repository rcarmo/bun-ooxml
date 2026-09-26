@planned
Feature: Bun retains mutation obligations beyond the shared audit seed
  These additive contracts retain checks from the original Bun backlog. Shared
  office_patch scenario IDs refer to operation semantics, not a required MCP transport.

  @id-office-preview-details
  Scenario: A dry-run preview describes the requested edit without committing it
    Given a presentation with the title "Original title"
    When a title change to "Changed by dry run" is previewed
    Then the preview identifies the original target and requested replacement
    And the preview reports zero committed changes

  @id-office-docx-exact-match-counts
  Scenario: Word match counts belong to each requested placeholder
    Given a document containing "<Present>" once and not containing "<Missing>"
    When both placeholders are resolved before mutation
    Then "<Present>" reports exactly one match
    And "<Missing>" reports exactly zero matches

  @id-office-xlsx-independent-style-reader
  Scenario: An independent reader accepts every referenced cell style
    Given a saved workbook containing a newly introduced cell style
    When an independent reader opens the workbook
    Then it reads the edited cell without an invalid style index
    And its identity and version are recorded separately from the writer
