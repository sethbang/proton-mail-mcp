# Demo recording guide

A single 30–45s GIF at the top of the README is the highest-conversion asset for
this project. People skim; one clip of an agent doing something only this server
enables is worth more than any paragraph. This file is the storyboard + the
mechanics to record it.

## What to show (and why this clip)

Goal: in one take, communicate the **unique value** (privacy-first Proton +
Bridge-native IMAP) and the **trust differentiator** (safety: dry-run preview
before any destructive action). Inbox triage hits both.

**Scenario: "Clean up newsletters, safely."**

The narrative arc the viewer should read in ~40 seconds:

1. A natural-language ask to the agent.
2. The agent *looks* before it leaps (read-only analysis).
3. The agent *previews* a bulk action with `dryRun` — the safety beat.
4. The agent commits, and the result is verified.

## Beat-by-beat script

| Time | On screen | Point it makes |
|------|-----------|----------------|
| 0:00–0:04 | User prompt: **"Which newsletters clutter my inbox most, and archive the noisiest one."** | Real, relatable task |
| 0:04–0:12 | Agent calls `top_senders(folder: "INBOX", since: "2026-01-01")` → a tidy frequency table renders | Inbox analytics; Proton-native |
| 0:12–0:20 | Agent calls `bulk_move(match: { listId: "..." }, destination: "Archive", dryRun: true)` → preview lists the exact UIDs it *would* move | **Safety beat** — previews before acting |
| 0:20–0:30 | Agent calls `bulk_move(...)` for real → response confirms N moved | The payoff |
| 0:30–0:38 | Agent calls `count_messages` / `list_messages` on Archive to verify | Honest accounting; it checks its work |
| 0:38–0:42 | Brief end card / cursor rest | Breathing room for the loop |

Keep it to **one clear task**. Resist showing 6 tools — the dry-run → commit →
verify loop is the story.

## Recording mechanics

- **Use a throwaway / QA Proton account**, not your personal one. Seed it with a
  handful of real newsletters so `top_senders` has something to show.
- Set `RESTRICT_OUTBOUND_TO_SELF=true` while recording so a stray send can't
  fan out to a real address.
- Recommended client: **Claude Desktop** (clean tool-call UI) at a slightly
  reduced window size so text is legible when scaled down.
- Hide anything sensitive: real email addresses, the account name, tokens. Blur
  or use obviously-fake seed data.

### Capture → GIF

macOS, terminal/desktop capture to an optimized GIF:

```bash
# 1. Record a screen region (macOS screen recorder, Cmd+Shift+5) → save demo.mov

# 2. Convert to a web-friendly, looping GIF (needs ffmpeg + gifski)
ffmpeg -i demo.mov -vf "fps=12,scale=900:-1:flags=lanczos" -f yuv4mpegpipe - \
  | gifski --quality 80 --fps 12 -o docs/demo.gif -
```

Targets: **≤ ~6 MB**, width ~900px, 12 fps. GitHub inlines it; keep it small so
the README loads fast. If it's over budget, drop to 800px / 10 fps or trim dead
air.

> Prefer a terminal-only demo? [`asciinema`](https://asciinema.org) + `agg`
> produces a crisp, tiny GIF and avoids capturing any desktop chrome.

## Wiring it into the README

Once `docs/demo.gif` exists, in `README.md` replace the placeholder block under
`## Demo` with:

```markdown
![Proton Mail MCP demo](docs/demo.gif)
```

(The commented-out `<!-- ![...](docs/demo.gif) -->` line is already there — just
uncomment it and delete the "coming soon" line.)

## Reuse

This same GIF is the hero of every launch post (r/ProtonMail, r/mcp, Show HN) and
makes a good GitHub social-preview image (crop a representative frame to
1280×640). Record once, use everywhere.
