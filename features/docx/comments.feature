@implemented @bun
Feature: Existing Word comment threads retain their package custody
  This slice reads existing comments and changes only existing commentsExtended done flags.

  @id-docx-comments-inspection
  Scenario: Inspecting the pinned threaded comments does not dirty the package
    Given the pinned threaded Word comments package is opened
    When existing Word comments are inspected
    Then the three comment bodies and reply parent match the pinned fixture
    And the Word comment package bytes remain unchanged

  @id-docx-comments-resolution
  Scenario: Resolve and reopen one comment without rewriting its body or anchors
    Given the pinned threaded Word comments package is opened
    When Word comment "1" is marked resolved
    Then saving and reopening shows that comment resolved and its reply link intact
    And only the existing commentsExtended part differs from the input
    When Word comment "1" is reopened
    Then the original comment package member bytes are restored

  @id-docx-comments-noop
  Scenario: An already-open comment is an exact archive no-op
    Given the pinned threaded Word comments package is opened
    When Word comment "1" is reopened
    Then the Word comment package bytes remain unchanged

  @id-docx-comments-refusal
  Scenario Outline: Unsafe comment resolution refuses before mutation for <case>
    Given an existing Word comment refusal package <case>
    When Word comment resolution is attempted
    Then a typed comment refusal leaves every package byte unchanged
    Examples:
      | case              |
      | missing-extension |
      | duplicate-id      |
      | duplicate-para    |
      | missing-parent    |
      | cycle             |
      | invalid-done      |
      | wrong-mime        |
      | external-link     |
      | protected         |
      | revision-body     |
      | orphan-entry      |
