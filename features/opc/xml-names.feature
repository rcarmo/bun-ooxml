@implemented @bun
Feature: XML namespace names require valid prefix and local components
  XML Name scanning alone cannot establish that both QName components are NCNames.

  @id-xml-invalid-qname-components
  Scenario Outline: Invalid <component> namespace names refuse
    Given XML with a numeric-leading <component> QName component
    When the namespace-name fixture is parsed
    Then parsing refuses with the malformed XML code
    Examples:
      | component       |
      | element-local   |
      | attribute-local |
      | declared-prefix |

  @id-xml-unicode-qname-components
  Scenario: Unicode prefix and local names remain valid
    Given XML with valid Unicode prefix and local name components
    When the namespace-name fixture is parsed
    Then expanded element and attribute names retain their Unicode identity
    And source offsets still address the original Unicode element

  @id-xml-outside-root-nbsp
  Scenario Outline: A non-breaking space <position> the root refuses
    Given XML with a non-breaking space <position> the root element
    When the namespace-name fixture is parsed
    Then parsing refuses with the malformed XML code
    Examples:
      | position |
      | before   |
      | after    |
