Feature: Strict ZIP custody for OPC packages
  The shared ZIP layer reads OOXML package archives only when they have one
  unambiguous ZIP32 interpretation and stay within bounded resources.

  Scenario: Read a valid ZIP32 archive with stored, deflated, and directory entries
    Given a ZIP archive with canonical OPC member names
    And the archive contains stored and deflated file entries
    And the archive may contain zero-byte directory entries and a declared archive comment
    When the archive is read through the shared ZIP module
    Then file entries are returned in central-directory order
    And directory entries do not become package parts
    And each returned payload matches its declared CRC and size

  Scenario: Refuse ambiguous or unsafe ZIP structure
    Given a ZIP archive whose structure has no single safe reading
    When the archive is read through the shared ZIP module
    Then the module refuses duplicate member names
    And the module refuses ASCII case-colliding member names
    And the module refuses noncanonical member paths
    And the module refuses encrypted or unsupported-compression members
    And the module refuses multi-disk and ZIP64 archives until ZIP64 is implemented
    And the module refuses local-header metadata that disagrees with the central directory
    And the module refuses CRC failures, size mismatches, and undeclared trailing structure

  Scenario: Refuse archives that exceed configured bounds before expansion
    Given a ZIP archive whose declared archive size, entry count, entry size,
      total expanded size, or compression ratio exceeds the configured limit
    When the archive is read through the shared ZIP module
    Then the module refuses before allocating unbounded output

  Scenario: Write deterministic UTF-8 ZIP32 output
    Given a map of canonical OPC member names and bytes
    When the map is written through the shared ZIP module twice
    Then both outputs are byte-identical ZIP32 archives
    And file entries use stored or deflated encoding
    And names are emitted with the UTF-8 ZIP flag
    And writer input that would collide by ASCII case is refused
