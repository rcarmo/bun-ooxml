# Third-party sources and fixtures

`references/manifest.json` records repository URLs, immutable commits and file hashes.

* OOXML DOCX includes OOXML by Steve Canny and contributors plus OOXML
  Instruments additions. MIT licence: `references/fixtures-ooxml/reference-assets/docx/LICENSE`.
* OOXML PPTX includes OOXML by Steve Canny and contributors plus OOXML
  Instruments additions. MIT licence: `references/fixtures-ooxml/reference-assets/pptx/LICENSE`.
* OOXML XLSX includes OOXML and OOXML Instruments additions. MIT licence:
  `references/fixtures-ooxml/reference-assets/xlsx/LICENCE.rst`.
* go-ooxml fixtures are MIT under `references/go-ooxml/LICENSE.md`. Their provenance
  is in `fixtures/go-ooxml/testdata/FIXTURES.md`; text uses public-domain Frankenstein.
* python-office-mcp-server has no repository licence at the pinned revision.
  Its files are retained for this private development repository at the owner's
  request. Clarify redistribution permission before making the repository public.
  Its fixture provenance is in `fixtures/python-office-mcp-server/tests/_templates/FIXTURES.md`.

`docs/contracts/office-mutation/audit/` retains the read-only audit export from
2026-09-26, including recipe/test sources and diagnostic observations. Its
`manifest.json` records artifact hashes and the reviewed revision. It does not
supply additional licensed runtime code or validated Office output goldens. Keep
the same private-development redistribution restriction as the MCP reference.

`docs/contracts/shared-v2/pack/` contains four deterministic fixtures derived from
committed MCP templates, eight planned workflow contracts, validation evidence and
preparation scripts. Origin and per-member hashes are recorded in the fixture
manifest; `pack/notices/` retains the original template/library notices. These
files share the private-development restriction until redistribution is reviewed.
Preparation scripts are not executed by Bun tests or runtime operations.

Frozen Python source/test files are reference inputs. Runtime exports include only
TypeScript under `src/`. Gherkin, messages and TypeScript packages are development
dependencies with their own bundled licence notices.
