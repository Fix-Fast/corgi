You are running the Corgi docx-plugin onboarding flow. Your job: replace the user's marketplace-installed copy of the plugin with a local working clone they can edit and push from.

Be terse. If a step fails, stop and surface the error — do not paper over it.

All plugin operations use the `claude plugin` CLI subcommand via Bash, **not** the in-session SlashCommand tool (which may not be loaded). Examples: `claude plugin uninstall <name>`, `claude plugin marketplace add <path>`, `claude plugin install <name>@<marketplace>`.

## Step 1 — Uninstall the existing docx plugin

- Read `~/.claude/plugins/installed_plugins.json`.
- For each key matching `docx@corgi` (or any marketplace whose name suggests Fix-Fast/corgi), run `claude plugin uninstall <key>` via Bash.
- If a `corgi` marketplace is registered, run `claude plugin marketplace remove corgi`. Same for any `fix-fast` marketplace.

## Step 2 — Ensure git is installed

- Run `git --version`. If it works, skip the rest of this step.
- macOS: if `brew` exists, `brew install git`. Otherwise instruct the user to run `xcode-select --install` and stop here.
- If another OS: stop. Not supported outside of MacOS.

## Step 3 — Pick a clone location

- Use AskUserQuestion to ask where to clone. Offer `~/workspace/src/corgi` as the default (recommended) option.
- `mkdir -p` the parent directory of the chosen path.
- If the chosen path already exists and is non-empty:
  - If it looks like an existing Corgi checkout (a git repo whose `origin` URL contains `Fix-Fast/corgi`), tell the user "looks like you already have a corgi checkout there — ok to keep going and reuse it?" via AskUserQuestion. If yes, skip Step 4 and proceed to Step 5 with this path.
  - Otherwise, stop and ask the user how to proceed (don't silently overwrite).

## Step 4 — Clone over HTTPS

- `git clone https://github.com/Fix-Fast/corgi.git <chosen-path>`
- HTTPS clone of a public repo is anonymous — no GitHub account or SSH key is needed for this step.
- After clone, `cd <chosen-path>` so subsequent git operations target the clone.

## Step 5 — Register the clone as a local marketplace, enable auto-update, and install

- `claude plugin marketplace add <chosen-path>` via Bash.
- Enable auto-update on the marketplace by setting `extraKnownMarketplaces.corgi.autoUpdate = true` in `~/.claude/settings.json`. As of this writing the `claude plugin marketplace add` CLI does not expose an `--auto-update` flag, so edit the JSON directly (use `uv run python -c "..."` with stdlib `json` to read/modify/write — keep it atomic, don't shell out to `sed`). The resulting marketplace entry should look like:
  ```json
  "corgi": {
    "source": { "source": "directory", "path": "<chosen-path>" },
    "autoUpdate": true
  }
  ```
- `claude plugin install docx@corgi` via Bash.

## Step 6 — Verify and hand off

- Re-read `~/.claude/plugins/installed_plugins.json`. Confirm `docx@corgi` is present.
- Tell the user:
  - Where the working clone lives (absolute path).
  - That auto-update is on for the `corgi` marketplace, so edits in that clone are picked up automatically — they just need to run `/reload-plugins` in their Claude Code session for changes to take effect in the running session.
  - That **pushing** to the upstream repo requires a GitHub account with write access plus either a personal access token (HTTPS) or SSH key. If they don't have one yet, point them at https://github.com/settings/tokens or https://docs.github.com/en/authentication/connecting-to-github-with-ssh — don't try to set this up for them. Add a note to CLAUDE.md locally if they do not have a github account setup.
