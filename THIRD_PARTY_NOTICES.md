# Licences and shared reference inputs

The Bun implementation is MIT licensed. Development dependencies retain their
package licences. No external implementation or external test source is vendored.

The optional development oracle downloads DocumentFormat.OpenXml 3.5.1 and its
framework (MIT) and System.IO.Packaging 10.0.2 (MIT) through NuGet. Their package
licence metadata stays with the installed packages; no SDK source is vendored.
LibreOffice and Poppler are separately installed test tools with their own
licences. Production APIs never invoke them.

Document fixtures and required original notices are held by the tagged
references/fixtures-ooxml submodule. Its manifest, NOTICES.md and notices directories
retain origin revisions, hashes and licence obligations. The root MIT licence does
not relicense those inputs. Do not remove required notices when repackaging.
