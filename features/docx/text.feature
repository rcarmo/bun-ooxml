@implemented @bun
Feature: DOCX slice paragraph text search and run-safe replacement
  The first DOCX slice reads body paragraphs, finds exact text across runs,
  preserves unaffected run formatting and refuses broader paragraph topologies
  until later slices land.

  @id-docx-format-preserve
  Scenario: DOCX slice preserves formatting across a cross-run replacement after save and reopen
    Given DOCX slice fixture "fixture-12183fb28e49ea1c1ac63de2252b011580c4cfed0a7cfd35efc1ffb20c94653e" is opened
    When DOCX slice paragraph 1 exact text "text and italic text" is replaced with "Tone and tilted text"
    And DOCX slice document is saved and reopened
    Then DOCX slice paragraph 1 text equals "Bold Tone and tilted text and underlined text and colored text and large text"
    And DOCX slice paragraph 1 run formatting around the replacement is preserved

  @id-docx-xml-space
  Scenario: DOCX slice preserves boundary whitespace across a cross-run replacement after save and reopen
    Given DOCX slice fixture "synthetic-whitespace" is opened
    When DOCX slice paragraph 1 exact text "phaBe" is replaced with "pha Be"
    And DOCX slice document is saved and reopened
    Then DOCX slice paragraph 1 text equals "Alpha Beta"
    And DOCX slice paragraph 1 text nodes preserve boundary whitespace

  @id-docx-table-paragraph
  Scenario: DOCX slice includes table-cell paragraphs in document order for exact replacement
    Given DOCX slice fixture "fixture-8192955ef935f09eb61a9fe6805d4996c811efcf54c0c966f52d983e38e0a79c" is opened
    When DOCX slice paragraph 5 exact text "Galvanic battery" is replaced with "Voltaic battery"
    And DOCX slice document is saved and reopened
    Then DOCX slice paragraph 1 text equals "Research Materials Inventory"
    And DOCX slice paragraph 2 text equals "Item"
    And DOCX slice paragraph 5 text equals "Voltaic battery"
    And DOCX slice paragraph 13 text equals "University library"

  @id-docx-stale-span
  Scenario: DOCX slice refuses a stale span without mutating the package
    Given DOCX slice fixture "fixture-12183fb28e49ea1c1ac63de2252b011580c4cfed0a7cfd35efc1ffb20c94653e" is opened
    And DOCX slice paragraph 1 span "text and italic text" is remembered
    And DOCX slice paragraph 1 exact text "colored text" is replaced with "scarlet text"
    And DOCX slice current saved bytes are remembered
    When DOCX slice stale replacement "Tone and tilted text" is attempted on the remembered span
    Then DOCX slice refusal code equals "docx-stale-span"
    And DOCX slice saved bytes equal the remembered bytes

  @id-docx-refuse-topology
  Scenario Outline: DOCX slice refuses unsupported paragraph topology for <fixture> paragraph <paragraph> exact search "<query>"
    Given DOCX slice fixture "<fixture>" is opened
    When DOCX slice paragraph <paragraph> exact search for "<query>" is attempted
    Then DOCX slice refusal code equals "docx-unsupported-topology"

    Examples:
      | fixture                                                                                | paragraph | query              |
      | fixture-e3c5159fbf254f4d5423354773ae83a5535cb3cef8f603f1adca6d18adab11f2   | 1         | amazing            |
      | fixture-e4f051ec2eb5f48b9b8299e931abb2bca1fa86ca5007865b3b2f9b83ba16676f | 2         | [Enter Title Here] |
      | synthetic-field                                                                        | 1         | 2026-01-01         |
