# Mac Claude — First-Session Setup Prompt

> Paste the block below into a fresh Claude Code session on the reformatted Mac.
> It tells Claude everything it needs to bootstrap from the Google Drive vault.
> No file copying required — the vault is the source of truth.

---

## Copy-paste this into Claude on the Mac

You are starting a fresh session on a freshly reinstalled Mac. The user is Neil
Daley. Everything you need to work with him already lives in his Google Drive
vault — your job is to bootstrap from it, confirm everything is reachable, and
then stop and wait for instructions.

### What this machine is for

This Mac is Neil's writing machine for the eToro & Investing project. He will
mainly use it to draft Substack articles, Substack Notes, X posts, and eToro
trade posts, then save them to the vault Drafts folder. He also has a Windows
PC that does the same work; both machines read the same Google Drive vault, so
the source of truth is shared.

### Machine paths on this Mac

The macOS username is `neilmax`. Google Drive on this Mac mounts under
`CloudStorage`. The paths you need:

- **Vault root:**
  `/Users/neilmax/Library/CloudStorage/GoogleDrive-ndaley1313@gmail.com/My Drive/Daley's Brain`
- **eToro Sync folder (portfolio + valuation data):**
  `/Users/neilmax/Library/CloudStorage/GoogleDrive-ndaley1313@gmail.com/My Drive/eToro Sync`

If those exact paths do not exist, look under
`/Users/neilmax/Library/CloudStorage/` for the actual Google Drive mount name
(it sometimes varies by Drive client version) and use what you find.

The Windows equivalents (which the vault's CLAUDE.md files refer to in places)
are `C:\Users\Neil\My Drive\Daley's Brain` and `C:\Users\Neil\ClaudeCode\eToro\`.
Translate Windows paths in the vault files to the Mac paths above when reading
or writing files. The vault content is identical; only the path prefix changes.

### Step 1 — Mount the vault and confirm access

Use `request_cowork_directory` (or the equivalent in this Claude client) to
gain access to the vault at:

`/Users/neilmax/Library/CloudStorage/GoogleDrive-ndaley1313@gmail.com/My Drive/Daley's Brain`

Then list the top-level contents and confirm you can see at least:

- `CLAUDE.md`
- `About Me/`
- `Projects/`
- `Personal/`
- `Goals.md`

If any of those are missing, stop and tell Neil — Google Drive may not be
fully synced yet on the fresh install.

### Step 2 — Read the canonical instructions in order

Read these four files, in this order, before doing anything else:

1. `Daley's Brain/CLAUDE.md` — global system instructions (rules + project
   context that applies everywhere)
2. `Daley's Brain/About Me/about-me.md` — who Neil is
3. `Daley's Brain/About Me/writing-rules.md` — voice, style, banned words
4. `Daley's Brain/About Me/memory.md` — project status, decisions, open loops,
   session log

These four files are the canonical source of truth. They are kept in the vault
specifically so they persist across sessions and across machines. Do not read
or edit any local `~/.claude/About Me/` files — that location is retired.

After reading them, also read the project-scoped instructions for the active
project:

5. `Daley's Brain/Projects/eToro & Investing/CLAUDE.md`

This one tells you the publishing cadence, pre-publish checklist, voice rules,
and how the data files relate to each other.

### Step 3 — Confirm the eToro data files are reachable

The project's valuation work depends on three files in the eToro Sync folder
(the Windows CLAUDE.md calls this folder `C:\Users\Neil\ClaudeCode\eToro\data\`
— on this Mac it is the Google Drive folder noted above). Confirm you can read:

- `eToro_Master.xlsx` — master output workbook (portfolio tracker)
- `etoro_master.json` — JSON lookup by ticker (generated from eToro_Master)
- `combined_portfolio.json` — current holdings snapshot
- `StockValuerEtoro.xlsx` — upstream valuation model (DCF/DDM/EPV logic)

The first three are the ones touched most often during article writing. The
fourth is the upstream model; you read it when methodology questions come up.

If any are missing on the Mac side, tell Neil — they should be in Google Drive
and syncing automatically.

### Step 4 — Report back, then stop

Once steps 1–3 are done, send Neil a short summary (under 200 words):

- Confirm the vault path you actually used
- Confirm the eToro Sync path you actually used
- Confirm you've read the four About Me / CLAUDE.md files plus the eToro
  project CLAUDE.md
- Confirm you can read each of the four data files
- Flag anything missing or different from what's described above

Then stop and wait for Neil's first real instruction. Do **not** start writing,
drafting, or "exploring" the vault beyond what's needed to confirm access. He
will tell you what to work on.

### Background facts you should know going in

These are settled and do not need re-asking:

- Neil writes in British English. No em dashes anywhere. No banned words from
  `writing-rules.md`. Numbers lead. Specific over adjectives.
- The publishing rhythm is Monday outlook → Tuesday FTSE Tracker article →
  Wednesday/Thursday Notes and tweets → Friday deep-dive article.
- Drafts go in `Projects/eToro & Investing/Drafts/`. Published articles in
  `Published/`. Never overwrite a file in `Published/`.
- Never publish, send, or post anything. Drafts only. This includes Gmail,
  Substack, X, eToro, Reddit, Threads, YouTube.
- The two spreadsheets do different jobs: `StockValuerEtoro.xlsx` is the
  upstream model; `eToro_Master.xlsx` is the downstream output. The project
  CLAUDE.md explains this in detail.
- Auto-memory on this Mac will start fresh. That is fine. The persistent
  context is in vault `memory.md`, which you just read.

### One last thing

If anything in this prompt conflicts with what you find in the vault files,
the vault files win. They are the canonical source; this prompt is just the
bootstrap.

---

## End of prompt

When the Mac Claude finishes step 4 and reports back, the machine is ready
for normal work.
