@implemented @bun
Feature: Native Word tracked replacement has accept and reject outcomes
  This slice supports one exact match in a simple run paragraph, not general Word Compare.

  @id-docx-tracked-replace-roundtrip
  Scenario: A fragmented phrase becomes native insertion and deletion revisions
    Given a native Word package with a bold fragmented payment phrase
    When that phrase is replaced as tracked changes with an explicit author and date
    Then reopening exposes the original and current text in their review views
    And accepting yields the revised text while rejecting restores the original text
    And the starting run formatting and unrelated part bytes survive resolution

  @id-docx-tracked-replace-refusal
  Scenario Outline: Unsafe tracked edits leave all bytes unchanged for <case>
    Given a native Word tracked-edit refusal case <case>
    When the tracked replacement is attempted
    Then it refuses with a typed error and preserves the original archive bytes
    Examples:
      | case       |
      | ambiguous  |
      | drawing    |
      | existing   |
      | protected  |
      | bad-date   |
