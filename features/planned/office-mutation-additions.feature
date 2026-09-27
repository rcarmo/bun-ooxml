@planned
Feature: Bun retains mutation obligations beyond the shared audit seed
  These additive contracts retain checks from the original Bun backlog. Shared
  office_patch scenario IDs refer to operation semantics, not a required MCP transport.

  @id-office-xlsx-independent-style-reader
  Scenario: An independent reader accepts every referenced cell style
    Given a saved workbook containing a newly introduced cell style
    When an independent reader opens the workbook
    Then it reads the edited cell without an invalid style index
    And its identity and version are recorded separately from the writer
