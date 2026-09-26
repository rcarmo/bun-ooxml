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

Frozen Python source/test files are reference inputs. Runtime exports include only
TypeScript under `src/`. Gherkin, messages and TypeScript packages are development
dependencies with their own bundled licence notices.
