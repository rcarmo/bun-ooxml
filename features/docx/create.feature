@implemented @bun
Feature: DOCX native paragraph authoring
  Native DOCX creation authors minimal packages and simple paragraphs without
  borrowed template bytes.

  @id-docx-create-minimal-package
  Scenario: DOCX create authors a minimal package with escaped bold italic text before section properties
    Given DOCX create source "new-document" is prepared
    When DOCX create paragraph "  <Frankenstein & friend>  " with bold and italic formatting is appended
    And DOCX create document is saved and reopened
    Then DOCX create package members equal "[Content_Types].xml,_rels/.rels,word/document.xml"
    And DOCX create paragraph 1 text equals "  <Frankenstein & friend>  "
    And DOCX create paragraph 1 run has bold and italic formatting
    And DOCX create paragraph 1 XML preserves boundary spaces and escapes special characters
    And DOCX create paragraph 1 is stored before the section properties

  @id-docx-create-style-validation
  Scenario: DOCX create appends a paragraph using a known paragraph style id
    Given DOCX create source "fixture-9548a1ce68caae9df12bc85732f1c19a098658c5dce3d79488814e4145299e5e" is prepared
    When DOCX create styled paragraph "Created heading" with style "Heading1" is appended
    And DOCX create document is saved and reopened
    Then DOCX create paragraph 5 text equals "Created heading"
    And DOCX create paragraph 5 uses paragraph style "Heading1"

  @id-docx-create-stale-opaque
  Scenario: DOCX create preserves opaque parts and invalidates earlier spans after append
    Given DOCX create source "fixture-9548a1ce68caae9df12bc85732f1c19a098658c5dce3d79488814e4145299e5e" is prepared
    And DOCX create paragraph 2 exact text "Genevese" span is remembered
    And DOCX create opaque part "docProps/core.xml" bytes are remembered
    When DOCX create plain paragraph "Appendix line" is appended
    And DOCX create current saved bytes are remembered
    And DOCX create stale replacement "Citizen" is attempted on the remembered span
    Then DOCX create refusal code equals "docx-stale-span"
    And DOCX create saved bytes equal the remembered bytes
    And DOCX create opaque part "docProps/core.xml" bytes are unchanged

  @id-docx-create-atomic-refusals
  Scenario Outline: DOCX create refuses <case> atomically
    Given DOCX create source "<source>" is prepared
    And DOCX create current saved bytes are remembered
    When DOCX create append refusal "<case>" is attempted
    Then DOCX create refusal code equals "<code>"
    And DOCX create saved bytes equal the remembered bytes

    Examples:
      | source                                                                       | case                      | code                   |
      | new-document                                                                 | text-number               | docx-invalid-argument  |
      | new-document                                                                 | options-string            | docx-invalid-argument  |
      | new-document                                                                 | bold-string               | docx-invalid-argument  |
      | new-document                                                                 | style-without-styles-part | docx-style-unsupported |
      | fixture-9548a1ce68caae9df12bc85732f1c19a098658c5dce3d79488814e4145299e5e | unknown-style             | docx-style-missing     |
