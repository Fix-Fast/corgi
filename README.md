# corgi — drive Microsoft Word live from Claude Code

A Claude Code plugin that lets Claude act on an **open** Word document — read, edit, format, run office.js, toggle Track Changes — via the Carson local Word bridge.

Ships with:

- **`word-bridge` skill** — guidance for driving a live Word session via the MCP below. Tells Claude to prefer live-Word edits when the bridge is reachable, and fall back to Anthropic's `docx` skill for offline file work.
- **`carson-word-bridge` MCP server** — stdio MCP that talks to the Carson Word add-in on `127.0.0.1`, so Claude can manipulate the document that's open right now.

The installer also adds Anthropic's [`document-skills`](https://github.com/anthropics/skills) plugin (docx/xlsx/pptx/pdf) as a fallback for offline `.docx` work — this plugin is the live-Word layer on top of that.

## Install

One-liner (installs this plugin and the Word add-in):

```sh
curl -fsSL https://word.usecarson.com/install.sh | bash
```

Manual:

```sh
# Anthropic's document-skills (fallback for offline .docx work)
claude plugin marketplace add anthropics/skills
claude plugin install document-skills@anthropic-agent-skills

# This plugin (live Word via the Carson bridge)
claude plugin marketplace add Fix-Fast/corgi
claude plugin install docx@corgi
```

Then sideload the Word add-in from <https://word.usecarson.com>.

## Prerequisites

- **Carson desktop** running locally — hosts the bridge server on `127.0.0.1:3137` and multiplexes multiple Claude sessions across multiple Word documents.
- **Word add-in** installed and open in at least one document — this is what the MCP actually controls.
- **Node 20+** (for `npx carson-word-bridge-mcp`).
- **Claude Code** with Anthropic's `document-skills` plugin (auto-installed by the one-liner) for offline `.docx` work.

## How it fits together

```
Claude Code
   │
   ├── Anthropic document-skills (docx/xlsx/pptx/pdf)  ──► python-docx etc. (offline files)
   │
   ├── word-bridge skill (this plugin)  ──► guides use of the MCP below, MCP-first
   │
   └── MCP: carson-word-bridge
              │  stdio
              ▼
         carson-word-bridge-mcp  (npm)
              │  https://127.0.0.1:3137
              ▼
         Carson desktop bridge  (multiplexes N Claude sessions ↔ M Word docs)
              │
              ▼
         Word add-in task pane in an open document
```

The desktop bridge is required because multiple Claude sessions and multiple open Word documents need to be multiplexed onto one local control plane; the MCP server is a thin stdio adapter, not a standalone bridge.

## Insurance Policy Formatter

This repo also contains an offline formatter for Corgi-Tech insurance
policy `.docx` files.

- spec: [format.md](/Volumes/workspace/src/corgi/format.md)
- packaged human rules: [plugins/docx/skills/insure-policy-format/format.md](/Volumes/workspace/src/corgi/plugins/docx/skills/insure-policy-format/format.md)
- skill: [plugins/docx/skills/insure-policy-format/SKILL.md](/Volumes/workspace/src/corgi/plugins/docx/skills/insure-policy-format/SKILL.md)
- scripts: [plugins/docx/skills/insure-policy-format/scripts](/Volumes/workspace/src/corgi/plugins/docx/skills/insure-policy-format/scripts)

Public scripts:

- `rule_1.py` — converge text hierarchy: title, section headings,
  subheadings, and body text
- `rule_2.py` — converge list structure and list formatting
- `rule_3.py` — converge page layout and running header

Shared modules:

- `formatter.py`
- `_ooxml.py`
- `format.py`

### Quick Start

Run the full formatter:

```sh
uv run plugins/docx/skills/insure-policy-format/scripts/format.py \
  /abs/path/to/input.docx \
  -o /abs/path/to/output.docx \
  --parts-in /tmp/policy.parts.json
```

Run the same process rule by rule:

```sh
uv run plugins/docx/skills/insure-policy-format/scripts/rule_1.py \
  /abs/path/to/input.docx \
  -o /tmp/policy.docx \
  --parts-in /tmp/policy.parts.json

uv run plugins/docx/skills/insure-policy-format/scripts/rule_3.py \
  /abs/path/to/input.docx \
  /tmp/policy.docx \
  --parts-in /tmp/policy.parts.json

uv run plugins/docx/skills/insure-policy-format/scripts/rule_2.py \
  /tmp/policy.docx
```

### Scope And Guarantees

- The formatter is currently intended for Corgi-Tech insurance policy
  documents, not arbitrary `.docx` files.
- Claude Code is expected to inspect the source document first and pass
  an explicit document-parts manifest into the formatter.
- The rule pipeline reproduces the known fixture formatting for
  `../formatter/assets/policy_documents/cgl_original.docx`.
- Accepted match:
  `word/document.xml`, `word/numbering.xml`, and `word/styles.xml`
  must be identical to the reference output.
- `docProps/core.xml` timestamps may differ between runs.

### Reference Fixture

Input:

- [cgl_original.docx](</Volumes/workspace/src/formatter/assets/policy_documents/cgl_original.docx>)

Expected formatted output:

- [cgl_original.formatted.docx](</Volumes/workspace/src/formatter/assets/policy_documents/cgl_original.formatted.docx>)

## Troubleshooting

- **"bridge unreachable"** — start Carson desktop and make sure a Word document with the add-in task pane is open.
- **MCP not showing up** — run `claude plugin list` to confirm `docx@corgi` is enabled; restart Claude Code.
- **Re-install the add-in** — `curl -fsSL https://word.usecarson.com/install.sh | bash` is idempotent.

## License

Plugin contents are © Fix-Fast. The MCP server is published separately as [`carson-word-bridge-mcp`](https://www.npmjs.com/package/carson-word-bridge-mcp). Claude Code's built-in `docx` skill is owned by Anthropic and is **not** redistributed here — Claude Code ships it natively.
