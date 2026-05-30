# Launch posts

Ready-to-post drafts for announcing v1.0.0. Post **after** the demo GIF exists —
every one of these is stronger with the clip embedded.

General rules that apply to all of them:

- **Disclose it's yours and that it's unofficial.** Reddit and HN punish stealth
  self-promotion hard; leading with "I built this, not affiliated with Proton"
  earns goodwill instead.
- **Lead with the use case, not the feature count.** "Triage my inbox from
  Claude" beats "30+ tools."
- **Pre-empt the trust question.** The first comment on anything that touches
  email + an LLM + your password is "is this safe?" Answer it in the post.
- **Be present for the first 2–3 hours** to reply. Engagement in the first hour
  is most of what determines reach.
- Don't cross-post all three the same hour; space them out over a few days so you
  can fold feedback from the first into the next.

---

## r/mcp  (warmest audience — post here first)

**Title:** `Proton Mail MCP server — triage, search, and send Proton email from any MCP client (unofficial, MIT)`

**Body:**

> I built an MCP server for Proton Mail and just tagged v1.0.0. It gives an MCP
> client (Claude Desktop, Claude Code, Cursor, …) full read/write access to a
> Proton mailbox over SMTP + IMAP-through-Bridge.
>
> **What it can do:** send/reply/reply-all/forward (with Markdown bodies and
> attachments), read and search (by sender, subject, body, date, attachment
> name/type, List-ID…), bulk move/delete/flag/label with dry-run previews,
> folder + label management, thread walks, and inbox analytics (`top_senders`,
> `folder_stats`). ~30 tools total.
>
> **Why it's a bit different from the usual email MCP:** it's the
> privacy-focused, Proton one, and I leaned hard into *safety* because it's an
> agent touching real email — delete moves to Trash by default, `READONLY=true`
> mode, `dryRun` on every bulk op, HTML sanitization on outbound by default, a
> self-only outbound lock for throwaway accounts, and message bodies are fenced
> as untrusted (prompt-injection surface) when handed to the model.
>
> Unofficial — not affiliated with Proton AG. MIT licensed.
>
> Repo: https://github.com/sethbang/proton-mail-mcp
> npm: `npx -y proton-mail-mcp`
>
> Happy to answer questions about the Bridge/IMAP quirks I had to work around —
> there were a few fun ones (SEARCH lagging FETCH, All Mail eventual
> consistency, the `\Noselect` label thing).

---

## r/ProtonMail  (check the subreddit rules first; some require flair / mod OK for projects)

**Title:** `I built an (unofficial) MCP server so you can use Proton Mail from AI assistants like Claude`

**Body:**

> Heads up: this is a community project, **not** affiliated with or endorsed by
> Proton — I just use Proton and wanted my AI assistant to be able to help with
> email without handing it to a third party.
>
> It runs **locally** and talks to Proton the same way any mail client does:
> SMTP for sending, and IMAP through **Proton Mail Bridge** for reading. Nothing
> goes to any server I run — there's no backend, it's a local process your MCP
> client launches.
>
> With it, you can ask an assistant to do things like *"summarize my unread from
> this week,"* *"archive every newsletter from this sender,"* or *"draft a reply
> to this thread,"* and it carries them out through your own Proton account.
>
> On the safety side (because it's email + an LLM): it defaults to moving deletes
> to Trash, has a full **read-only mode**, previews every bulk action before
> doing it, sanitizes outbound HTML, and treats message contents as untrusted
> input. You use a Proton **SMTP password** (the app-specific one), not your main
> login password.
>
> Open source (MIT): https://github.com/sethbang/proton-mail-mcp
>
> Feedback welcome — especially from anyone who wants to try it against their own
> Bridge setup.

---

## Show HN  (highest ceiling, one shot — save for after the demo + a day of r/mcp feedback)

**Title:** `Show HN: Proton Mail MCP server – let an AI assistant work your Proton inbox`

(HN titles: no emoji, no "v1.0.0", keep it plain. The "Show HN:" prefix is
required for Show HN.)

**Body (first comment, posted by you immediately after submitting):**

> Author here. This is an MCP (Model Context Protocol) server that lets an MCP
> client — Claude Desktop, Cursor, etc. — read, search, organize, and send email
> through a Proton Mail account. It's unofficial and not affiliated with Proton.
>
> The interesting/hard parts, since this crowd will ask:
>
> - **It's local-only.** Sending goes over Proton's standard SMTP submission
>   endpoint; reading goes over IMAP via the locally-run Proton Mail Bridge. No
>   private API, no backend service — the MCP client spawns it as a subprocess.
> - **Safety was the main design constraint,** because it's an autonomous-ish
>   agent with write access to real email. Deletes go to Trash by default;
>   there's a `READONLY` mode; every bulk op (move/delete/flag/label) has a
>   `dryRun` that resolves and shows exactly what it would touch first; outbound
>   HTML is sanitized by default; there's an opt-in lock that refuses to send to
>   anyone but yourself (for test accounts); and message bodies are fenced as
>   untrusted content when passed to the model, since an inbox is a prompt-
>   injection vector.
> - **Bridge has quirks** I had to defend against: IMAP SEARCH lagging FETCH on
>   freshly-sent mail, All Mail being eventually consistent after moves, Proton's
>   label model cascading folder moves, and populated labels reporting
>   `\Noselect`. The tools detect and explain these rather than silently
>   returning wrong results.
>
> Node/TypeScript, MIT. `npx -y proton-mail-mcp`, or clone and build.
> Repo: https://github.com/sethbang/proton-mail-mcp
>
> Would love feedback on the tool surface and the safety model.

---

## Other low-effort spots (drop the same one-liner + repo link)

- **X/Twitter / Mastodon (Fosstodon):** one-liner + the demo GIF + repo link.
  Tag `#MCP` and `#ProtonMail`. The GIF carries this one.
- **MCP Discord servers** (`#showcase` channels).
- **lobste.rs** if you have an account — tag `ai`, `show`. Same plain title as HN.
- A short **dev.to / blog post** titled something like "Giving Claude safe access
  to my Proton inbox" — narrative version of the HN comment, good for SEO and
  linkable from the others.
