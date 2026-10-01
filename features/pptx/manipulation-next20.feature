@implemented @bun
Feature: Bounded local PPTX manipulation alignment
  Staged Python captures retain their original IDs and bytes.
  These concrete local cases earn no shared execution credit.

  # Source staging/python/features/test_pptx_advanced_tools.feature:136 @candidate-python-pptx-advanced-tools-2978ea684e
  @id-bun-pptx-next20-patch-title
  Scenario: title replacement
    Given the concrete PPTX manipulation input is patch-title
    When the production manipulation is performed for patch-title
    And the saved PPTX is reopened for patch-title
    Then the exact manipulated values and XML properties match patch-title
    And unrelated payloads and the supplied archive retain custody for patch-title

  # Source staging/python/features/test_pptx_advanced_tools.feature:147 @candidate-python-pptx-advanced-tools-6b32b859c7
  @id-bun-pptx-next20-patch-body
  Scenario: body replacement
    Given the concrete PPTX manipulation input is patch-body
    When the production manipulation is performed for patch-body
    And the saved PPTX is reopened for patch-body
    Then the exact manipulated values and XML properties match patch-body
    And unrelated payloads and the supplied archive retain custody for patch-body

  # Source staging/python/features/test_pptx_coverage.feature:37 @candidate-python-pptx-coverage-d416dd3f35
  @id-bun-pptx-next20-patch-subtitle
  Scenario: subtitle replacement
    Given the concrete PPTX manipulation input is patch-subtitle
    When the production manipulation is performed for patch-subtitle
    And the saved PPTX is reopened for patch-subtitle
    Then the exact manipulated values and XML properties match patch-subtitle
    And unrelated payloads and the supplied archive retain custody for patch-subtitle

  # Source staging/python/features/test_pptx_coverage.feature:47 @candidate-python-pptx-coverage-eaffd532fb
  @id-bun-pptx-next20-append-title
  Scenario: append title text
    Given the concrete PPTX manipulation input is append-title
    When the production manipulation is performed for append-title
    And the saved PPTX is reopened for append-title
    Then the exact manipulated values and XML properties match append-title
    And unrelated payloads and the supplied archive retain custody for append-title

  # Source staging/python/features/test_pptx_advanced_tools.feature:341 @candidate-python-pptx-advanced-tools-f8284125fc
  @id-bun-pptx-next20-bullet-default
  Scenario: append body bullet
    Given the concrete PPTX manipulation input is bullet-default
    When the production manipulation is performed for bullet-default
    And the saved PPTX is reopened for bullet-default
    Then the exact manipulated values and XML properties match bullet-default
    And unrelated payloads and the supplied archive retain custody for bullet-default

  # Source staging/python/features/test_pptx_coverage.feature:67 @candidate-python-pptx-coverage-574504134e
  @id-bun-pptx-next20-bullet-sequence
  Scenario: append two ordered bullets
    Given the concrete PPTX manipulation input is bullet-sequence
    When the production manipulation is performed for bullet-sequence
    And the saved PPTX is reopened for bullet-sequence
    Then the exact manipulated values and XML properties match bullet-sequence
    And unrelated payloads and the supplied archive retain custody for bullet-sequence

  # Source staging/python/features/test_pptx_advanced_tools_extended.feature:199 @candidate-python-pptx-advanced-tools-extended-190ce8dfa3
  @id-bun-pptx-next20-bullet-level
  Scenario: direct bullet indentation
    Given the concrete PPTX manipulation input is bullet-level
    When the production manipulation is performed for bullet-level
    And the saved PPTX is reopened for bullet-level
    Then the exact manipulated values and XML properties match bullet-level
    And unrelated payloads and the supplied archive retain custody for bullet-level

  # Source staging/python/features/test_pptx_advanced_tools_extended.feature:209 @candidate-python-pptx-advanced-tools-extended-e255895bfd
  @id-bun-pptx-next20-bullet-bold-label
  Scenario: bold label only
    Given the concrete PPTX manipulation input is bullet-bold-label
    When the production manipulation is performed for bullet-bold-label
    And the saved PPTX is reopened for bullet-bold-label
    Then the exact manipulated values and XML properties match bullet-bold-label
    And unrelated payloads and the supplied archive retain custody for bullet-bold-label

  # Source staging/python/features/test_pptx_coverage.feature:57 @candidate-python-pptx-coverage-107f971923
  @id-bun-pptx-next20-clear-bullets
  Scenario: clear body paragraphs
    Given the concrete PPTX manipulation input is clear-bullets
    When the production manipulation is performed for clear-bullets
    And the saved PPTX is reopened for clear-bullets
    Then the exact manipulated values and XML properties match clear-bullets
    And unrelated payloads and the supplied archive retain custody for clear-bullets

  # Source staging/python/features/test_pptx_advanced_tools_extended.feature:112 @candidate-python-pptx-advanced-tools-extended-7e0c39b72e
  @id-bun-pptx-next20-autofit-shrink
  Scenario: shrink autofit
    Given the concrete PPTX manipulation input is autofit-shrink
    When the production manipulation is performed for autofit-shrink
    And the saved PPTX is reopened for autofit-shrink
    Then the exact manipulated values and XML properties match autofit-shrink
    And unrelated payloads and the supplied archive retain custody for autofit-shrink

  # Source staging/python/features/test_pptx_coverage.feature:186 @candidate-python-pptx-coverage-19ee72b28b
  @id-bun-pptx-next20-autofit-none
  Scenario: disable autofit
    Given the concrete PPTX manipulation input is autofit-none
    When the production manipulation is performed for autofit-none
    And the saved PPTX is reopened for autofit-none
    Then the exact manipulated values and XML properties match autofit-none
    And unrelated payloads and the supplied archive retain custody for autofit-none

  # Source staging/python/features/test_pptx_coverage.feature:196 @candidate-python-pptx-coverage-2196c941f8
  @id-bun-pptx-next20-autofit-resize
  Scenario: resize shape autofit
    Given the concrete PPTX manipulation input is autofit-resize
    When the production manipulation is performed for autofit-resize
    And the saved PPTX is reopened for autofit-resize
    Then the exact manipulated values and XML properties match autofit-resize
    And unrelated payloads and the supplied archive retain custody for autofit-resize

  # Source staging/python/features/test_pptx_coverage.feature:7 @candidate-python-pptx-coverage-c8b3a3643d
  @id-bun-pptx-next20-insert-start
  Scenario: insert title slide at start
    Given the concrete PPTX manipulation input is insert-start
    When the production manipulation is performed for insert-start
    And the saved PPTX is reopened for insert-start
    Then the exact manipulated values and XML properties match insert-start
    And unrelated payloads and the supplied archive retain custody for insert-start

  # Source staging/python/features/test_pptx_coverage.feature:17 @candidate-python-pptx-coverage-93b49d5212
  @id-bun-pptx-next20-insert-middle
  Scenario: insert title slide at index one
    Given the concrete PPTX manipulation input is insert-middle
    When the production manipulation is performed for insert-middle
    And the saved PPTX is reopened for insert-middle
    Then the exact manipulated values and XML properties match insert-middle
    And unrelated payloads and the supplied archive retain custody for insert-middle

  # Source staging/python/features/test_pptx_advanced_tools_extended.feature:7 @candidate-python-pptx-advanced-tools-extended-9077938164
  @id-bun-pptx-next20-reorder
  Scenario: relationship-ordered slides
    Given the concrete PPTX manipulation input is reorder
    When the production manipulation is performed for reorder
    And the saved PPTX is reopened for reorder
    Then the exact manipulated values and XML properties match reorder
    And unrelated payloads and the supplied archive retain custody for reorder

  # Source staging/python/features/test_pptx_advanced_tools_extended.feature:17 @candidate-python-pptx-advanced-tools-extended-59251c1cbd
  @id-bun-pptx-next20-reorder-refusal
  Scenario: invalid permutation atomic refusal
    Given the concrete PPTX manipulation input is reorder-refusal
    When the production manipulation is performed for reorder-refusal
    And the saved PPTX is reopened for reorder-refusal
    Then the exact manipulated values and XML properties match reorder-refusal
    And unrelated payloads and the supplied archive retain custody for reorder-refusal

  # Source staging/python/features/test_pptx_advanced_tools.feature:285 @candidate-python-pptx-advanced-tools-dec2108fa9
  @id-bun-pptx-next20-table-values
  Scenario: author exact 3x3 table values
    Given the concrete PPTX manipulation input is table-values
    When the production manipulation is performed for table-values
    And the saved PPTX is reopened for table-values
    Then the exact manipulated values and XML properties match table-values
    And unrelated payloads and the supplied archive retain custody for table-values

  # Source staging/python/features/test_pptx_advanced_tools.feature:297 @candidate-python-pptx-advanced-tools-f082c66869
  @id-bun-pptx-next20-table-geometry
  Scenario: EMU geometry and grid sums
    Given the concrete PPTX manipulation input is table-geometry
    When the production manipulation is performed for table-geometry
    And the saved PPTX is reopened for table-geometry
    Then the exact manipulated values and XML properties match table-geometry
    And unrelated payloads and the supplied archive retain custody for table-geometry

  # Source staging/python/features/test_pptx_advanced_tools.feature:157 @candidate-python-pptx-advanced-tools-fa6b261e45
  @id-bun-pptx-next20-set-notes
  Scenario: edit existing speaker text
    Given the concrete PPTX manipulation input is set-notes
    When the production manipulation is performed for set-notes
    And the saved PPTX is reopened for set-notes
    Then the exact manipulated values and XML properties match set-notes
    And unrelated payloads and the supplied archive retain custody for set-notes

  # Source staging/python/features/test_pptx_coverage.feature:86 @candidate-python-pptx-coverage-6c1a01c0a5
  @id-bun-pptx-next20-notes-readback
  Scenario: multiline notes save/reopen
    Given the concrete PPTX manipulation input is notes-readback
    When the production manipulation is performed for notes-readback
    And the saved PPTX is reopened for notes-readback
    Then the exact manipulated values and XML properties match notes-readback
    And unrelated payloads and the supplied archive retain custody for notes-readback
