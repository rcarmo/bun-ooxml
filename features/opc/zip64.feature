@implemented @bun
Feature: Bounded single-disk ZIP64 custody
  ZIP64 metadata is validated before entry expansion and never relaxes resource limits.

  @id-parity-zip64
  Scenario: A small forced ZIP64 package preserves all member payloads
    Given a ZIP64 archive with two native-written XML members
    When Bun reads it and rewrites a changed member with ZIP64 enabled
    Then the changed member reopens with its new payload
    And the unrelated member retains its exact payload bytes

  @id-zip64-preflight-count
  Scenario: An impossible ZIP64 entry count refuses before allocation
    Given a ZIP64 archive declaring more entries than its directory can contain
    When Bun attempts bounded ZIP64 admission
    Then admission refuses with code "zip-structure-invalid"

  @id-zip64-unsafe-offset
  Scenario: An unrepresentable ZIP64 offset refuses without precision loss
    Given a ZIP64 locator offset above the safe integer range
    When Bun attempts bounded ZIP64 admission
    Then admission refuses with code "zip-zip64-unsupported"

  @id-zip64-resource-limit
  Scenario: Caller entry limits still apply to ZIP64
    Given a valid ZIP64 archive with two entries and a one-entry budget
    When Bun attempts bounded ZIP64 admission
    Then admission refuses with code "zip-too-many-entries"
