# Plugin interfaces

This repository is the free public reader. It reads spreadsheets. Provenance,
instrument templates, batch folder processing, and output signing ship as the
GLP Lab Data Pack, not in this tree.

Those features ship as the GLP Lab Data Pack inside mcp-glp-audit-proxy Pro.
The pack is unlocked by the same offline GLP Pro licence key and registers
implementations under the entry-point groups below. Nothing in this package
registers a plugin.

For regulated labs, a GLP Lab Data Pack (provenance, signed output, instrument
templates) is available as part of mcp-glp-audit-proxy Pro; contact via
https://antonio-castellon.github.io

| Group | Interface | Default in this package |
| --- | --- | --- |
| `mcp_advanced_excel.features` | `FeatureGate.enabled(feature)` | returns false, no I/O |
| `mcp_advanced_excel.provenance` | `ProvenanceProvider.annotate(view)` | returns `{}` |
| `mcp_advanced_excel.templates` | `TemplateProvider.match(view)` | returns `None` |
| `mcp_advanced_excel.batch` | `BatchProvider.run(folder, options)` | not active |
| `mcp_advanced_excel.signing` | `SigningProvider.sign(payload)` | returns the payload unchanged |

`enabled` is an offline hook. The default does not open a socket, read a
licence file, or contact a server. A replacement may read a local licence
file. It must not phone home. A feature runs only when the gate returns true
and a non-default provider is installed. With the defaults in this package,
read payloads are unchanged.

The read path asks the template provider for a `header_row` hint before
detection, and it passes the finished payload through provenance and signing.
Both calls are no-ops here. Batch is not a tool in this package. The interface
exists so the GLP Lab Data Pack can supply `run` without forking the reader.

Public modules that stay in this repo: `grid`, `detect`, `numbers`, `units`,
`normalize`, `safety`, `paths`, `config`, `server`, `plugins`.

Extension-pack implementations do not belong in this tree.
