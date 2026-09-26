@implemented @bun
Feature: DOCX story inspection
  @id-docx-story-parts
  Scenario: story parts follow the reachable relationship graph
    Given DOCX story fixture "synthetic-revisions" is opened
    When DOCX story parts are listed
    Then DOCX story parts equal:
      | part               | kind      |
      | word/document.xml  | body      |
      | word/header1.xml   | header    |
      | word/footer1.xml   | footer    |
      | word/footer2.xml   | footer    |
      | word/footnotes.xml | footnotes |
      | word/endnotes.xml  | endnotes  |
      | word/comments.xml  | comments  |

  @id-docx-story-views
  Scenario: revision views preserve document order and run breaks across stories
    Given DOCX story fixture "synthetic-revisions" is opened
    When DOCX stories are inspected in "current" view
    Then DOCX story "word/document.xml" paragraphs equal:
      | index | text                    |
      | 0     | Body intro              |
      | 1     | Table\tCell\nLine2    |
      | 2     | Inserted body paragraph |
      | 3     | Body kept               |
    And DOCX story "word/header1.xml" paragraphs equal:
      | index | text            |
      | 0     | Header new text |
    And DOCX story "word/footnotes.xml" paragraphs equal:
      | index | text               |
      | 0     | \u0020Footnote new text |
    When DOCX stories are inspected in "original" view
    Then DOCX story "word/document.xml" paragraphs equal:
      | index | text                   |
      | 0     | Body intro             |
      | 1     | Table\tCell\nLine2   |
      | 2     | Deleted body paragraph |
      | 3     | Body kept              |
    And DOCX story "word/header1.xml" paragraphs equal:
      | index | text            |
      | 0     | Header old text |
    And DOCX story "word/footnotes.xml" paragraphs equal:
      | index | text               |
      | 0     | \u0020Footnote old text |
    When DOCX stories are inspected in "all" view
    Then DOCX story "word/document.xml" paragraphs equal:
      | index | text                    |
      | 0     | Body intro              |
      | 1     | Table\tCell\nLine2    |
      | 2     | Inserted body paragraph |
      | 3     | Deleted body paragraph  |
      | 4     | Body kept               |
    And DOCX story "word/header1.xml" paragraphs equal:
      | index | text               |
      | 0     | Header newold text |
    And DOCX story "word/footnotes.xml" paragraphs equal:
      | index | text                  |
      | 0     | \u0020Footnote newold text |

  @id-docx-story-blind
  Scenario: blind regions are reported honestly and inspection does not mutate bytes
    Given DOCX story fixture "synthetic-blind" is opened
    And DOCX story package bytes are remembered
    When DOCX stories are inspected in "current" view
    Then DOCX blind regions equal:
      | part              | kind                        | count |
      | word/document.xml | alternate-content           | 1     |
      | word/document.xml | field                       | 2     |
      | word/document.xml | textbox                     | 1     |
      | word/document.xml | unknown-namespace           | 1     |
      | word/document.xml | external-story-relationship | 1     |
      | word/header9.xml  | orphan-story-part           | 1     |
    And DOCX story package bytes equal the remembered bytes
