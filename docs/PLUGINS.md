# Plugin interfaces

This repository is the community edition. It reads spreadsheets. It does not
implement provenance, instrument templates, batch folder processing, or
output signing.

A separate package can register implementations under these entry-point
groups. Nothing in this package registers a plugin.

| Group | Interface | Community default |
| --- | --- | --- |
| `mcp_advanced_excel.features` | `FeatureGate.enabled(feature)` | returns false, no I/O |
| `mcp_advanced_excel.provenance` | `ProvenanceProvider.annotate(view)` | returns `{}` |
| `mcp_advanced_excel.templates` | `TemplateProvider.match(view)` | returns `None` |
| `mcp_advanced_excel.batch` | `BatchProvider.run(folder, options)` | not active |
| `mcp_advanced_excel.signing` | `SigningProvider.sign(payload)` | returns the payload unchanged |

`enabled` is an offline hook. The default does not open a socket, read a
licence file, or contact a server. A replacement may read a local file. It
must not phone home. A feature runs only when the gate returns true and a
non-default provider is installed. With the community defaults, read payloads
are unchanged.

The read path asks the template provider for a `header_row` hint before
detection, and it passes the finished payload through provenance and signing.
Both calls are no-ops here. Batch is not a community tool. The interface
exists so another package can supply `run` without forking the reader.

Public modules that stay in this repo: `grid`, `detect`, `numbers`, `units`,
`normalize`, `safety`, `paths`, `config`, `server`, `plugins`.

Pro implementations do not belong in this tree.
