@implemented @bun
Feature: PPTX relationship-ordered reads and anchored text replacement
  The first PPTX slice opens real packages, reads existing notes without creating
  missing notes parts, and performs exact anchored text replacement across runs.

  @id-pptx-order-notes-read
  Scenario: Follow presentation relationships and keep notes reads non-mutating
    Given PPTX ordered notes fixtures are prepared
    When PPTX opens the reordered notes fixture and probes notes reads
    Then PPTX keeps slide order, notes blank lines, and notes reads non-mutating without creating missing notes parts

  @id-pptx-readable-unsupported-topology
  Scenario: Read line breaks and field text faithfully but refuse editing that topology
    Given PPTX line-break and field text fixture is prepared from a real template
    When PPTX inspects the paragraph text and attempts an anchored edit on that topology
    Then PPTX exposes line breaks and field text faithfully and refuses the unsupported edit without mutation

  @id-pptx-cross-run-replace
  Scenario: Replace exact anchored text across runs and preserve unrelated members after reopen
    Given PPTX fragmented title fixture is prepared from a real template and an untouched ZIP member
    When PPTX replaces anchored cross-run text and saves then reopens the package
    Then PPTX preserves the replacement text, the starting run formatting, and unrelated ZIP member bytes

  @id-pptx-stale-anchor-refusal
  Scenario: Refuse a stale anchored replacement without mutation
    Given PPTX stale-anchor fixture is prepared from a real template
    When PPTX replaces anchored text once and retries with the stale anchor
    Then PPTX refuses the stale anchor and keeps the post-success bytes unchanged
