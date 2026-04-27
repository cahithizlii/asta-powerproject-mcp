# Construction Planning MCP Suite

Three Model Context Protocol servers for construction project planning across
Asta Powerproject and Microsoft Project.

## Servers

| Server | Module | Purpose |
|---|---|---|
| `asta_powerproject_mcp` | `asta_mcp_core.py` | Live COM control of Asta Powerproject (8 tools): tasks, links, progress, resources, schedule, codes, views, exports. |
| `asta_powerproject_file` | `asta_mcp_file.py` | File-based read access via MPXJ/MSPDI for Asta `.pp/.xer/.xml` (4 tools): query, resources, calendar, edit. |
| `msproject_mcp` | `msproject_mcp_core.py` | Microsoft Project COM with hybrid bulk routing (3 tools, 17 actions): `msproject_task`, `msproject_link`, `msproject_schedule`. |

## Phase 1 Status — MS Project MCP

Phase 1 delivers a COM-based Microsoft Project MCP with hybrid bulk routing:

- **com_direct**: 1-9 ops via direct COM calls (lowest latency).
- **com_batch**: 10-19 ops with batch mode (manual calc, screen-update off, events disabled) for steady throughput.
- **mspdi_bulk**: 20+ ops via in-memory MSPDI XML build + single COM `OpenProject` import — 200 tasks in **2.30 s** measured (target was <5 s).

All 17 actions across the three tools are operational with full pytest coverage
(43/43 passing) including a 200-task end-to-end villa acceptance test.

## Quick Start

1. Install deps: `pip install -r requirements.txt`
2. Open Asta Powerproject and/or MS Project so a live COM endpoint exists.
3. Register the servers with Claude Desktop / your MCP client — see
   [`samples/claude_mcp_config_snippet.json`](samples/claude_mcp_config_snippet.json)
   for the exact `mcpServers` block to merge into your config.
4. Restart the client; the three servers (`asta_powerproject_mcp`,
   `asta_powerproject_file`, `msproject_mcp`) appear as tool sources.

## Testing

```
python -m pytest tests/ -v
```

Tests automatically skip if MS Project / Asta is not running on the host.

## Repository Layout

- `msproject_mcp_core.py` — MS Project MCP (Phase 1).
- `msproject_bulk.py` — MSPDI bulk-import writer.
- `mspdi_parser.py` — Native MSPDI XML reader/writer (shared by both ecosystems).
- `asta_mcp_core.py` — Asta Powerproject COM MCP.
- `asta_mcp_file.py` — Asta file-based MCP.
- `tools/` — One-shot helper scripts (typelib dump, fixture generator).
- `tests/` — Pytest suite. Fixtures in `tests/fixtures/` (see fixtures README).
- `docs/plans/` — Phase 1 design + implementation plan.
- `samples/` — End-to-end build scripts and MCP config snippet.
