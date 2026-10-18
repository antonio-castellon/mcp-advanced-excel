# Plugin interfaces

This repository is the free public reader. The GLP Lab Data Pack is a paid
extension for regulated (GLP / GxP) labs. It is part of mcp-glp-audit-proxy
Pro and is unlocked by the same offline GLP Pro licence key. The licence
check stays on the machine and does not phone home. The pack plugs into the
interfaces below. Nothing in this package registers a plugin.

The pack is designed to support GLP / 21 CFR Part 11 / EU GMP Annex 11
workflows.

The open Community audit proxy is
https://github.com/antonio-castellon/mcp-glp-audit-proxy

- Provenance hashes. A SHA-256 of the source file and of the extraction
  parameters is stored with every table, so a result can be traced back to
  the raw data (ALCOA+).
- Instrument templates. Ready parsers recognise common instrument export
  layouts, so a lab can read those files without writing a new header rule
  for each instrument.
- Advanced LOD/LOQ and molar units. Results below the limit of detection or
  quantification stay flagged, and concentrations can be converted to molar
  units, so a report keeps the original qualifier and can compare amounts.
- Batch processing. Many files are processed in one run and returned with one
  summary, so a study folder has a single manifest instead of one extraction
  per export.
- Signed output. Ed25519 signatures are attached to the output, so a later
  change to that output is detectable.

For regulated labs, a GLP Lab Data Pack (provenance, signed output, instrument
templates) is available as part of mcp-glp-audit-proxy Pro; contact
[antonio@castellon.ch](mailto:antonio@castellon.ch).

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
