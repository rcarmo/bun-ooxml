@implemented @bun
Feature: OPC package custody and transactional part edits
  The shared OPC layer preserves whole-archive bytes on no-op saves, rolls back
  failed transactional edits, and keeps unrelated payload bytes opaque.

  @id-opc-package-corpus-noop
  Scenario: Opening and reopening the fixture corpora without edits preserves whole archives
    Given the go-ooxml and python-office-mcp-server fixture corpora are enumerated
    When each OOXML fixture package is opened and serialized without edits through the OPC layer
    Then every reopened package matches its original whole-archive bytes
    And both fixture corpora contribute their exact known nonzero fixture counts

  @id-opc-package-transaction-rollback
  Scenario: A failed transactional edit rolls back every changed part
    Given a valid OPC package with XML and opaque payload parts
    When a transactional edit changes multiple parts and then fails
    Then the package reverts to the original bytes and parts after the refusal

  @id-opc-package-preserve-unrelated
  Scenario: Changing one part preserves unrelated payload after reopen
    Given a valid OPC package with a main XML part and an unrelated binary payload
    When the main XML part text is changed and the package is reopened
    Then the edited part contains the new text after reopen
    And the unrelated payload bytes remain unchanged
