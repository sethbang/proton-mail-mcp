# Contributing

Thanks for your interest in improving Proton Mail MCP! This is an independent,
community-built project (unofficial — not affiliated with Proton AG), and
contributions are welcome.

## Ways to help

- **Try it and report what breaks.** Bug reports against real Proton Mail Bridge
  setups are especially valuable — Bridge has subtle, version-dependent
  behaviors and more eyes catch more of them.
- **Improve docs.** If the README or a tool description tripped you up, a PR that
  clarifies it helps the next person.
- **Pick up an issue.** Look for issues labeled
  [`good first issue`](https://github.com/sethbang/proton-mail-mcp/labels/good%20first%20issue)
  or [`help wanted`](https://github.com/sethbang/proton-mail-mcp/labels/help%20wanted).

## Development setup

Requires **Node.js 24+**.

```bash
git clone https://github.com/sethbang/proton-mail-mcp.git
cd proton-mail-mcp
npm install
npm run build
```

Useful scripts:

```bash
npm run build        # compile TypeScript
npm run watch        # compile in watch mode
npm test             # run the vitest suite
npm run test:watch   # tests in watch mode
npm run lint         # ESLint
npm run format       # Prettier (write)
npm run inspector    # launch the MCP inspector against the built server
```

### Trying changes against a live mailbox

Use a **throwaway / test Proton account**, never your personal one, when
exercising send/delete/move behavior. Two env flags make this safe:

- `READONLY=true` disables every mutating tool.
- `RESTRICT_OUTBOUND_TO_SELF=true` refuses any live send to a non-self
  recipient (so a stray send can't fan out to a real address in seed data).

See the README's [Configuration](README.md#configuration) section for the full
env var list.

## Before you open a PR

CI runs `build`, `lint`, and `test` (see `.github/workflows/`). Run them locally
first:

```bash
npm run build && npm run lint && npm test
```

- Keep PRs focused — one logical change per PR is easiest to review.
- Add or update tests for behavior changes.
- If you change a tool's surface (params, defaults, response shape), update the
  matching section of the README so the docs stay honest.
- Commit messages follow [Conventional Commits](https://www.conventionalcommits.org/)
  (`feat:`, `fix:`, `docs:`, `chore:`, …).

## Reporting security issues

Please **do not** open a public issue for security-sensitive reports. See
[SECURITY.md](SECURITY.md) for how to report privately.

## License

By contributing, you agree that your contributions will be licensed under the
project's [MIT License](LICENSE).
