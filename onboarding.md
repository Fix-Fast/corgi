You are running the Corgi docx-plugin onboarding flow. Your job: replace the user's marketplace-installed copy of the plugin with a local working clone they can edit and push from.

Be terse. If a step fails, stop and surface the error — do not paper over it.

## Step 1 — Uninstall the existing docx plugin

- Read `~/.claude/plugins/installed_plugins.json`.
- For each key matching `docx@corgi` (or marketplace Fix-Fast), run `/plugin uninstall <key>` via the SlashCommand tool.
- If a `corgi` marketplace is registered, run `/plugin marketplace remove corgi` (same for Fix-Fast marketplace).

## Step 2 — Ensure git is installed

- Run `git --version`. If it works, skip the rest of this step.
- macOS: if `brew` exists, `brew install git`. Otherwise instruct the user to run `xcode-select --install` and stop here.
- If another OS: stop. Not supported outside of MacOS.

## Step 3 — Pick a clone location

- Use AskUserQuestion to ask where to clone. Offer `~/workspace/src/corgi` as the default (recommended) option.
- `mkdir -p` the parent directory of the chosen path.
- If the chosen path already exists and is non-empty, stop and ask the user how to proceed (don't silently overwrite).

## Step 4 — Clone over HTTPS

- `git clone https://github.com/Fix-Fast/corgi.git <chosen-path>`
- HTTPS clone of a public repo is anonymous — no GitHub account or SSH key is needed for this step.

## Step 5 — Register the clone as a local marketplace and install

- `/plugin marketplace add <chosen-path>` via SlashCommand.
- `/plugin install docx@corgi` via SlashCommand.

## Step 6 — Verify and hand off

- Re-read `~/.claude/plugins/installed_plugins.json`. Confirm `docx@corgi` is present and its `installPath` reflects the new local install.
- Tell the user:
  - Where the working clone lives.
  - That edits should be made in that clone, then `/plugin marketplace update corgi` + `/plugin install docx@corgi` to pick them up locally.
  - That **pushing** to the upstream repo requires a GitHub account with write access plus either a personal access token (HTTPS) or SSH key. If they don't have one yet, point them at https://github.com/settings/tokens or https://docs.github.com/en/authentication/connecting-to-github-with-ssh — don't try to set this up for them. Add a note to CLAUDE.md locally if they do not have a github account setup.
