@planned
Feature: Cross-language limitations require independent full-port contracts
  These obligations are not satisfied by the bounded shared mutation fixtures.

  @id-docx-review-multistory-resolution
  Scenario: Accept and reject revisions across supported document stories
    Given a document with body header footer and footnote revisions
    When supported insertion deletion move and formatting revisions are resolved
    Then accepting yields the revised content in every selected story
    And rejecting yields the original content in every selected story
    And unsupported revision forms are reported without partial mutation

  @id-docx-track-changes-option-outcome
  Scenario: A track-changes option changes the saved review representation
    Given a document edit with explicit track-changes intent
    When the edit is saved through the native revision API
    Then saved Word-native revisions match that intent
    And the option is not ignored by a workflow dispatcher

  @id-docx-commentsextended-content-type
  Scenario: Extended comments use an independently validated content type
    Given conflicting upstream commentsExtended content-type constants
    When the native implementation selects a content type for a fixture
    Then an authoritative schema or format reference supports that choice
    And an independent Office consumer accepts the saved comment thread

  @id-xlsx-derived-cache-completeness
  Scenario: Broader spreadsheet caches require range and relationship awareness
    Given formula ranges chart caches external-link caches and calculation-chain metadata
    When a precedent input is changed under a declared cache policy
    Then every supported dependent representation is invalidated or recalculated correctly
    And an unowned dependency refuses before a successful freshness claim
