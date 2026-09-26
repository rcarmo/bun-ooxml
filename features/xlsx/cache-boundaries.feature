@implemented @bun
Feature: Spreadsheet cache invalidation has explicit boundaries
  Clearing worksheet formula caches does not calculate or refresh every derived
  representation in an Office package. Unsupported result ranges must refuse.

  @id-xlsx-array-input-refusal
  Scenario Outline: A <topology> formula prevents unsafe input edits
    Given a two-sheet workbook with a <topology> formula and cached follower cells
    When an ordinary input value edit is attempted
    Then the edit refuses with code "xlsx-cache-topology-unsupported"
    And every workbook member and cached result remains byte-identical
    Examples:
      | topology  |
      | array     |
      | dataTable |

  @id-xlsx-cache-scope-opaque-parts
  Scenario: Worksheet invalidation preserves opaque chart and external-link caches without refreshing them
    Given a workbook with ordinary worksheet formulas and opaque chart and external-link caches
    When its input is changed and the workbook is saved and reopened
    Then worksheet formula cache values are cleared and their formulas retained
    And the opaque chart and external-link cache payloads remain unchanged
