@implemented @bun
Feature: Office relationship attributes use expanded XML names
  Package relationship elements and officeDocument relationship attributes use
  different namespace URIs. Prefix spellings never establish attribute identity.

  @id-office-relationship-prefix-alias
  Scenario Outline: A valid alternate relationship prefix survives <format> editing
    Given the shared <format> fixture with its relationship prefix changed to link
    When Bun opens it edits its text and saves then reopens it
    Then the requested <format> value survives and all relationship targets resolve
    And the original relationship attribute spelling is preserved
    Examples:
      | format |
      | pptx   |
      | xlsx   |

  @id-office-relationship-wrong-uri
  Scenario Outline: A relationship attribute in the wrong namespace refuses for <format>
    Given the shared <format> fixture with r bound to package relationships
    When Bun attempts to open the namespace-mismatched Office document
    Then the <format> reader refuses with its structural error code
    And the supplied archive bytes remain unchanged
    Examples:
      | format |
      | pptx   |
      | xlsx   |
