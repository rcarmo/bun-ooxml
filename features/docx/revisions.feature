@implemented @bun
Feature: DOCX bounded revisions
  @id-docx-revisions-all-stories
  Scenario: accepting bounded revisions across every reachable supported story
    Given the synthetic DOCX revisions fixture "all-stories" is opened
    Then the DOCX revision story texts for view "current" are
      | part                | text                      |
      | word/document.xml   | Body modernized.          |
      | word/header1.xml    | Header modernized.        |
      | word/footer1.xml    | Footer modernized.        |
      | word/footer2.xml    | Nested footer modernized. |
      | word/footnotes.xml  | Footnote modernized.      |
      | word/endnotes.xml   | Endnote modernized.       |
      | word/comments.xml   | Comment modernized.       |
    And the DOCX revision story texts for view "original" are
      | part                | text                    |
      | word/document.xml   | Body historic.          |
      | word/header1.xml    | Header historic.        |
      | word/footer1.xml    | Footer historic.        |
      | word/footer2.xml    | Nested footer historic. |
      | word/footnotes.xml  | Footnote historic.      |
      | word/endnotes.xml   | Endnote historic.       |
      | word/comments.xml   | Comment historic.       |
    When all DOCX revisions are accepted
    And the DOCX revision package is saved and reopened
    Then the DOCX changed parts are
      | part               |
      | word/document.xml  |
      | word/header1.xml   |
      | word/footer1.xml   |
      | word/footer2.xml   |
      | word/footnotes.xml |
      | word/endnotes.xml  |
      | word/comments.xml  |
    And the DOCX revision count is 0
    And the DOCX revision story texts for view "all" are
      | part                | text                      |
      | word/document.xml   | Body modernized.          |
      | word/header1.xml    | Header modernized.        |
      | word/footer1.xml    | Footer modernized.        |
      | word/footer2.xml    | Nested footer modernized. |
      | word/footnotes.xml  | Footnote modernized.      |
      | word/endnotes.xml   | Endnote modernized.       |
      | word/comments.xml   | Comment modernized.       |

  @id-docx-revisions-rollback
  Scenario: a later unsupported selected nested story rolls the whole resolution back
    Given the synthetic DOCX revisions fixture "rollback-unsupported" is opened
    And the current DOCX revision bytes are remembered
    Then the DOCX revision story texts for view "current" are
      | part              | text       |
      | word/document.xml | Body safe. |
    And the DOCX revision story texts for view "original" are
      | part              | text      |
      | word/document.xml | Body old. |
    When DOCX revisions are accepted for parts
      | part             |
      | word/document.xml |
      | word/footer2.xml  |
    Then the DOCX revision refusal code equals "docx-revisions-unsupported"
    And the DOCX revision refusal mentions "word/footer2.xml"
    And the DOCX revision refusal mentions "format"
    And the DOCX revision saved bytes equal the remembered bytes
    And the DOCX revision count is 2
    And the DOCX revision story texts for view "current" are
      | part              | text       |
      | word/document.xml | Body safe. |
    And the DOCX revision story texts for view "original" are
      | part              | text      |
      | word/document.xml | Body old. |

  @id-docx-revisions-empty-deletion
  Scenario: Rejecting an empty deletion text node keeps well-formed saved XML
    Given the synthetic DOCX revisions fixture "empty-deletion" is opened
    When all DOCX revisions are rejected
    And the DOCX revision package is saved and reopened
    Then the DOCX revision count is 0
    And the DOCX revision story texts for view "all" are:
      | part              | text |
      | word/document.xml |      |

  @id-docx-revisions-protected
  Scenario: protected documents refuse resolution without mutating bytes
    Given the synthetic DOCX revisions fixture "protected" is opened
    And the current DOCX revision bytes are remembered
    When all DOCX revisions are accepted
    Then the DOCX revision refusal code equals "docx-revisions-protected"
    And the DOCX revision refusal mentions "documentProtection"
    And the DOCX revision saved bytes equal the remembered bytes
    And the DOCX revision count is 2
