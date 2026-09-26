@implemented @bun
Feature: PPTX native title-slide authoring
  This slice authors conservative OOXML title slides that reopen through the
  native reader with validated relationships and content types. It does not
  claim visual or Office rendering validation.

  @id-pptx-create-new-minimal
  Scenario: Create a minimal presentation and reopen two authored text slides in order
    Given a new native PPTX presentation is created from scratch
    When PPTX adds two text slides and saves then reopens the package
    Then PPTX reopens both slides in order with title and subtitle placeholder text and valid minimal relationships

  @id-pptx-create-anchor-preservation
  Scenario: Append to a compatible existing deck without invalidating untouched slide anchors
    Given a compatible existing PPTX title-slide deck is prepared
    When PPTX appends a new text slide and edits the original title through its earlier anchor
    Then PPTX keeps the original slide content ordered and preserved except for the anchored edit and gives the new slide independent relationships

  @id-pptx-create-refusals
  Scenario: Refuse invalid arguments and unsafe existing layouts without mutation
    Given invalid-argument and unsafe-layout PPTX fixtures are prepared
    When PPTX attempts unsupported creation edits on those fixtures
    Then PPTX refuses both requests with rollback and leaves their package bytes unchanged
