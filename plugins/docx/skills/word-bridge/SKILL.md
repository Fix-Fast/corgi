---
name: word-bridge
description: "Use this skill when the user wants Claude to drive a live Microsoft Word document — reading, editing, formatting, or running office.js against the document that is currently open. Triggers include: 'edit the open Word doc', 'change the heading in Word', 'redline this paragraph', 'run office.js', 'track changes on', or any request that implies acting on a Word session rather than a standalone .docx file. For offline .docx file creation/editing (no open Word session), use Anthropic's `docx` skill (from the `document-skills` plugin in `anthropics/skills`) instead."
---

# Word bridge — live Word control via MCP

This skill covers the live-document workflow: Claude drives an open Word document via the `carson-word-bridge` MCP server, which talks over `127.0.0.1` to the Carson Word add-in task pane running inside Word.

Use this skill when the user wants Claude to **act on a Word document that is open right now**. For offline `.docx` work (create, read, or edit a file without Word running), use Anthropic's `docx` skill (from the `document-skills` plugin in `anthropics/skills`) instead — the two skills compose.

## Prerequisites

1. Carson desktop is running locally (hosts the bridge on `127.0.0.1:3137`).
2. The Carson Word add-in is installed and the task pane is open in at least one document.
3. The `carson-word-bridge` MCP server is connected (shipped with this plugin).

If any of these is missing, call `bridge_status` first — it returns health + queue counters and will surface the failure mode.

## MCP tools

The `carson-word-bridge` MCP server exposes these tools. Prefer the higher-level tools over `execute_office_js` unless you specifically need a primitive not covered below.

**Inspection**
- `bridge_status` — health + queue counters; use as a first call if anything looks off.
- `list_documents` — enumerate open Word documents the bridge can see.
- `refresh_doc_state` — recompute compact state for the active document (call after external edits).
- `inspect_selection` — what the user currently has selected.
- `read_doc_section` — read a section by heading or paragraph range.
- `search_doc_text` — literal-text search across the active document.
- `snapshot_document` — capture a labeled snapshot (e.g. `before` / `after`) for diff/verification.

**Editing**
- `apply_formatting` — structured formatting ops (headings, bold, styles, etc.).
- `execute_office_js` — escape hatch: run arbitrary office.js against the document. Use sparingly.
- `set_track_changes` — toggle Word Track Changes.
- `add_note` — attach a comment/note.

**Verification & session**
- `verify_doc` — run validators on document state.
- `export_session` — export the current Corgi session object.

## Operating guidance

- **Always read before you write.** Call `read_doc_section` or `inspect_selection` before applying edits, so you're acting on current state.
- **Use snapshots for non-trivial edits.** `snapshot_document` with `label="before"`, do the edit, `snapshot_document` with `label="after"`. This lets you recover and lets the user audit.
- **Prefer `apply_formatting` over `execute_office_js`.** The structured tool handles the common cases safely; only drop to office.js for things it can't express.
- **Turn on Track Changes for redlines.** Call `set_track_changes` with `enabled=true` before redline-style edits so the user can review them in Word.
- **Verify after substantial changes.** Run `verify_doc` to catch broken styles or malformed tables.

## When to use this skill vs Anthropic's `docx` skill

**Default: prefer the live Word bridge (this skill) whenever the bridge is reachable.** Editing in Word gives the user immediate visual feedback, Track Changes, and the ability to accept/reject. Only fall back to the Anthropic `docx` skill (from the `document-skills` plugin in the `anthropics/skills` marketplace) when the bridge can't do the job.

Routing:

| Situation | Skill |
| --- | --- |
| User is clearly working live ("edit the open doc", redlines, formatting) | `word-bridge` |
| User hands you a path to a `.docx` file to create or transform | **`word-bridge` first** — open it in Word via the bridge and edit live. Fall back to Anthropic `docx` only if the bridge is unreachable or the task is purely batch (e.g. processing many files). |
| User wants to generate a polished report/memo/letter from scratch, no Word running | Anthropic `docx` (then optionally open the result in Word afterwards) |
| Batch / headless processing of many `.docx` files | Anthropic `docx` |
| Bridge returns "unreachable" and user can't start Carson desktop right now | Anthropic `docx` |

If uncertain, call `bridge_status` first. If it's healthy, default to `word-bridge`. If it's not, use Anthropic `docx` and tell the user why.

## Troubleshooting

- **Bridge unreachable** — confirm Carson desktop is running; the bridge listens on `127.0.0.1:3137`.
- **No active document** — `list_documents` returns empty. Ask the user to open a document with the add-in task pane visible.
- **Edits not showing** — call `refresh_doc_state`; the document may have been edited externally.
