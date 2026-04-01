![Stream Transcript Extractor banner](assets/banner.svg)

`Stream Transcript Extractor` pulls Microsoft Teams recording transcripts from
Microsoft Stream through a real signed-in Chrome or Edge profile. It prefers
the transcript payload behind the Stream UI, falls back to DOM capture when
needed, and keeps the public source surface small enough to audit quickly.

This repository is designed around three outcomes:

- Keep extraction reliable across Stream transport changes.
- Keep the operator workflow small when the Stream UI gets inconsistent.
- Keep the source workflow easy to share: one core extractor runtime, one
  workflow-first CLI entrypoint, a shared CLI library, one build file, and
  generated README assets with a documented source of truth.

---

![Overview](assets/header-overview.svg)

The runtime surface stays intentionally small. [`extract.js`](./extract.js)
holds the Stream-specific extractor runtime and workflow routing,
[`lib/cli.js`](./lib/cli.js) holds the shared CLI parsing, help, and
terminal-selection logic, [`lib/cdp.js`](./lib/cdp.js)
holds the reusable browser session and navigation helpers, and
[`lib/network-capture.js`](./lib/network-capture.js) holds the reusable
response-body capture helpers. [`build.js`](./build.js) builds release
binaries, [`assets/`](./assets) holds the architecture diagram and generated
README graphics, and [`scripts/`](./scripts) is the source of truth for those
graphics and README validation.

The project supports two operator flows:

- `extract.js` is the single-meeting workflow. It supports three capture
  modes.
- `extract.js crawl` opens Stream home, switches to **Meetings**, scrolls the
  full list, merges newly discovered items into a persistent crawl state file,
  lets you select items in the terminal with progress-aware labels, and
  reuses the same browser session to run the `automatic` extractor for each
  selected meeting.

The help surface now follows the workflow too. `bun ./extract.js --help`
shows the top-level map, `bun ./extract.js --mode automatic --help` shows the
single-meeting network workflow, and `bun ./extract.js crawl --help` shows
the batch discovery, queue state, and selection flow.

When you launch `bun ./extract.js` with no arguments in a real terminal, it
opens the same interactive launcher that the compiled binary uses. That
launcher lets you choose `extract` or `crawl`, then set the recommended
workflow options before the run starts.

The extractor capture modes stay the same:

- `automatic` is the recommended and default path. It builds on the same
  network extractor, reloads the page with capture armed, tries to open the
  **Transcript** panel for you, nudges the UI to trigger lazy loading, and
  falls back to manual help only after it explains what it could and could
  not confirm.
- `network` is the advanced low-level path. It captures the Stream transcript
  payload directly, decrypts it when needed, and writes richer meeting
  metadata when you want to drive the panel-open timing yourself.
- `dom` is the fallback path. It scrolls the rendered transcript panel and
  extracts entries from the visible UI when the transport layer is unsuitable
  or when you need UI-level debugging.

Recommended CLI defaults:

- single-meeting extraction: `automatic`
- transcript output format: `md`
- crawl discovery settle window: `10000` ms (`10` seconds)
- diagnostics: leave `--debug` off unless you are diagnosing a failure

The crawl workflow always wraps the `automatic` extractor in one browser
session.

The tool does not use Microsoft Graph transcript APIs. It works through the
browser session you already have.

> [!IMPORTANT]
> The selected Chrome or Edge profile must already be able to open the target
> Microsoft Stream recordings and transcripts.

Use the source workflow when you want the simplest shareable entrypoint. Use a
release binary when you want a standalone executable without requiring Bun.

Requirements:

- macOS or Windows
- Google Chrome or Microsoft Edge
- A signed-in browser profile that can already open the recording and
  transcript
- Bun `1.2.4` or later for the source workflow
- Python 3 if you want to regenerate README assets or validate the README

You do not need Bun or Python when you run a built binary.

---

![Install and run](assets/header-install.svg)

You can run the extractor directly from source or use a standalone binary from
the release page. Both paths still require Chrome or Edge on the target
machine because the extractor attaches to a real browser profile.

Run from source:

```bash
mise install
bun ./extract.js
bun ./extract.js --help
bun ./extract.js --mode automatic --help
bun ./extract.js crawl --help
```

Run from a release binary:

```bash
# macOS arm64 example
./stream-transcript-extractor-macos-arm64
./stream-transcript-extractor-macos-arm64 --help
./stream-transcript-extractor-macos-arm64 crawl --help

# Windows x64 example
.\stream-transcript-extractor-windows-x64.exe
.\stream-transcript-extractor-windows-x64.exe --help
.\stream-transcript-extractor-windows-x64.exe crawl --help
```

The latest binaries are published on the
[latest release page](https://github.com/benjaminwestern/stream-transcript-extractor/releases/latest).

These commands cover the common operator paths:

```bash
# Recommended interactive launcher path
bun ./extract.js

# Recommended direct single-meeting path
bun ./extract.js --mode automatic

# Force a mode
bun ./extract.js --mode network
bun ./extract.js --mode dom

# Pick a browser or profile
bun ./extract.js --browser chrome --profile Work
bun ./extract.js --browser edge --profile "user@example.com"

# Choose output
bun ./extract.js --format md --output weekly-standup
bun ./extract.js --format both --output weekly-standup

# Diagnostics only when needed
bun ./extract.js --output-dir ./exports --debug

# Crawl Meetings and batch-extract selected items
bun ./extract.js crawl
bun ./extract.js crawl --browser chrome --profile Work
bun ./extract.js crawl --output-dir ./exports --format md
bun ./extract.js crawl --output-dir ./exports --format both
bun ./extract.js crawl --debug
bun ./extract.js crawl --state-file ./exports/team.state.json
bun ./extract.js crawl --wait-before-discovery-ms 10000
bun ./extract.js crawl --select pending --browser edge

# Show help or version
bun ./extract.js --help
bun ./extract.js --version
bun ./extract.js crawl --help
```

The extract workflow CLI options are:

| Option | Purpose |
| --- | --- |
| `--mode <network\|automatic\|dom>` | Choose the extractor mode. Recommended and default: `automatic`. |
| `--browser <chrome\|edge>` | Use a specific browser instead of choosing interactively. |
| `--profile <query>` | Match a profile by name, email, or directory name. |
| `--output <name>` | Override the output filename prefix. |
| `--output-dir <path>` | Write output files to a custom directory. |
| `--format <json\|md\|both>` | Write JSON, Markdown, or both. Recommended and default: `md`. |
| `--debug-port <port>` | Force a specific remote debugging port. |
| `--debug` | Write extra diagnostics for the selected mode. Recommended default: off. |
| `--keep-browser-open` | Leave the launched browser open after extraction. |
| `--help` | Print the CLI help text. |
| `--version` | Print the embedded build version. |

The no-arguments interactive launcher covers the common settings from the same
CLI contract:

- workflow selection: `extract` or `crawl`
- extract mode: `automatic`, `network`, or `dom`
- output format: `md`, `json`, or `both`
- diagnostics: `--debug` on or off
- browser preference: auto-detect, Chrome, or Edge
- crawl settle wait: recommended `10000` ms, `30000` ms, or a custom value
- advanced overrides: `--profile`, `--output-dir`, `--output`,
  `--debug-port`, and for crawl `--start-url`, `--state-file`, and `--select`

The crawl workflow CLI options are:

| Option | Purpose |
| --- | --- |
| `--browser <chrome\|edge>` | Use a specific browser instead of choosing interactively. |
| `--profile <query>` | Match a profile by name, email, or directory name. |
| `--start-url <url>` | Override the Stream home URL used before the crawler switches to **Meetings**. |
| `--state-file <path>` | Override the persistent crawl queue state file. |
| `--select <spec>` | Select queue items non-interactively with `pending`, `new`, `failed`, `done`, `all`, or numeric ranges. |
| `--wait-before-discovery-ms <ms>` | Delay discovery after the Stream URL opens so auth and page load can settle. Recommended and default: `10000` ms (`10` seconds). |
| `--output <name>` | Override the batch status filename prefix. |
| `--output-dir <path>` | Write transcript and batch status files to a custom directory. |
| `--format <json\|md\|both>` | Write JSON, Markdown, or both transcript outputs. Recommended and default: `md`. |
| `--debug-port <port>` | Force a specific remote debugging port. |
| `--debug` | Save extended network diagnostics for each extracted item. Recommended default: off. |
| `--keep-browser-open` | Leave the launched browser open after the batch run. |
| `--help` | Print the CLI help text. |
| `--version` | Print the embedded build version. |

---

![Modes](assets/header-modes.svg)

All three modes solve the same user problem, but they optimize for different
failure patterns. The entrypoint stays stable. The operator flow changes only
after you choose the mode.

| Mode | What it does | Use it when |
| --- | --- | --- |
| `automatic` | Uses the network path, reloads with capture armed, tries to open the **Transcript** panel, nudges the UI, retries once, and only then falls back to manual help. | Recommended default. You want the lowest operator effort while still using the network extractor. |
| `network` | Captures transcript-related network responses, loads response bodies, decrypts payloads when needed, and parses transcript entries from the transport layer. | You want the lowest-level path, richer metadata, and direct control over when the panel is opened. |
| `dom` | Scrolls the rendered transcript UI and extracts entries from the visible DOM. | You need a fallback when the transport changes, or you want to inspect the rendered transcript itself. |

Mode-specific operator flow:

- `network`: Open the recording page with the **Transcript** panel closed.
  Return to the terminal, arm capture, then open the panel after the extractor
  tells you capture is live.
- `automatic`: Open the recording page and return to the terminal. The
  extractor reloads the page, attempts the transcript-panel actions for you,
  and explains why it fell back if it still needs manual help.
- `dom`: Open the recording page and the **Transcript** panel before
  extraction begins. The extractor then scrolls the visible transcript UI.

Leave `--debug` off for normal runs. Use `--debug` when you want to see the
deeper diagnostics for the selected mode. In `automatic` mode, that includes
the full UI action trace, retry flow, and fallback reason.

<br>

---

![Outputs and diagnostics](assets/header-outputs.svg)

Each run can write JSON, Markdown, or both. The transcript outputs stay
intentionally simple so they work for review, archive, LLM input, or diffing
across browsers and operating systems.

The crawl workflow also writes two queue artifacts:

- `*.state.json` is the persistent crawl queue. It records discovered meeting
  URLs, progress state, and the latest extraction outcome per item so the next
  run can merge in new discoveries without losing prior progress.
- `*.batch.json` is the current run summary. It records the current discovery
  result, the terminal selection, and the success or failure state for each
  extracted item.

Mode-specific output behavior:

| Mode | Transcript files | Diagnostic behavior |
| --- | --- | --- |
| `network` | Writes `.json`, `.md`, or both based on `--format`. | `*.network.json` is written only when you use `--debug`, whether the run succeeds or fails. The terminal still prints a live confirmation when likely transcript traffic is seen. |
| `automatic` | Writes the same transcript files as `network`. | `*.network.json` is written only when you use `--debug`. `--debug` also prints the UI action trace, retry attempts, and fallback diagnosis. |
| `dom` | Writes the same transcript files as the other modes. | `--debug` writes `*.debug.json` with DOM-level diagnostics. |

Run summaries also distinguish between potentially relevant network responses
and actual parsed transcript payload matches, which makes failed captures less
misleading during triage.

<details>
<summary>Sample JSON transcript output</summary>

```json
{
  "meeting": {
    "title": "Weekly Standup - Project Alpha",
    "date": "March 30, 2026",
    "createdBy": "Jane Smith",
    "createdByEmail": "jane.smith@example.com",
    "sourceUrl": "https://contoso.sharepoint.com/.../stream.aspx?id=...",
    "recordingStartDateTime": "2026-03-30T00:04:18.0183508Z",
    "recordingEndDateTime": "2026-03-30T00:40:06.0592754Z"
  },
  "extractedAt": "2026-03-30T08:50:18.000Z",
  "entryCount": 142,
  "speakers": [
    "Alice Johnson",
    "Bob Chen",
    "Charlie Davis"
  ],
  "entries": [
    {
      "speaker": "Alice Johnson",
      "timestamp": "0:03",
      "text": "Whether it is ChatGPT or Gemini, that is more of an SEO problem."
    }
  ]
}
```

</details>

<details>
<summary>Sample Markdown transcript output</summary>

```md
Transcript information

Title: Weekly Standup - Project Alpha
Date: March 30, 2026
Start date/time: March 30, 2026 at 12:04:18 AM UTC
End date/time: March 30, 2026 at 12:40:06 AM UTC
Created by: Jane Smith <jane.smith@example.com>
Speakers: Alice Johnson, Bob Chen, Charlie Davis
Source URL: https://contoso.sharepoint.com/.../stream.aspx?id=...
Extracted at: 2026-03-30T08:50:18.000Z
Entry count: 142

---

Alice Johnson - 0:03:
Whether it is ChatGPT or Gemini, that is more of an SEO problem.
```

</details>

---

![How it works](assets/header-architecture.svg)

The extractor keeps the user flow small, but there is still a fair amount
happening under the hood so the browser session stays usable and the output
remains stable across Chrome, Edge, macOS, and Windows.

![Extractor architecture diagram](./assets/extractor-architecture.svg)

The diagram source lives in
[`assets/extractor-architecture.d2`](./assets/extractor-architecture.d2), and
the rendered asset lives in
[`assets/extractor-architecture.svg`](./assets/extractor-architecture.svg).

Three design decisions matter most:

- The shared extraction runtime and workflow routing live in
  [`extract.js`](./extract.js), so the batch flow reuses the same code paths
  instead of forking them behind a second entrypoint.
- Generic browser-session and network-capture plumbing live in
  [`lib/cdp.js`](./lib/cdp.js) and
  [`lib/network-capture.js`](./lib/network-capture.js), while Stream-specific
  scoring, parsing, and extraction policy stay in [`extract.js`](./extract.js).
- The extractor uses your existing signed-in browser profile instead of asking
  for credentials or app registrations.
- The `network` path is preferred because Stream's transcript UI is virtualized
  and more likely to drift than the underlying payload transport.

## Roadmap

The near-term roadmap stays scoped to the same meeting-page workflow. The goal
is to extend what the current browser session can do without growing a second
separate capture stack.

- TODO: Add optional MP4 download support to the crawl workflow so the
  selected meeting can save its recording while the browser is already on the
  Stream item page.

---

![Scripts and assets](assets/header-scripts.svg)

The README now has an explicit source of truth instead of hand-edited visual
assets. If you change the banner, section headers, or visual system, update
[`scripts/generate_assets.py`](./scripts/generate_assets.py) first, then
regenerate the SVGs in [`assets/`](./assets).

Repository script surface:

| Command | Purpose |
| --- | --- |
| `bun ./extract.js --help` | Print the extractor CLI contract. |
| `bun ./extract.js --mode automatic --help` | Print the mode-specific network capture workflow. |
| `bun ./extract.js crawl --help` | Print the crawl workflow CLI contract. |
| `bun ./build.js --help` | Print the standalone build contract. |
| `python3 scripts/generate_assets.py` | Regenerate the README banner and section header SVG files. |
| `python3 scripts/validate_readme.py` | Validate README asset references and local documentation links. |
| `mise run generate-assets` | Convenience wrapper for `python3 scripts/generate_assets.py`. |
| `mise run validate-docs` | Convenience wrapper for `python3 scripts/validate_readme.py`. |

Generated README asset files:

- [`assets/banner.svg`](./assets/banner.svg)
- [`assets/header-overview.svg`](./assets/header-overview.svg)
- [`assets/header-install.svg`](./assets/header-install.svg)
- [`assets/header-modes.svg`](./assets/header-modes.svg)
- [`assets/header-outputs.svg`](./assets/header-outputs.svg)
- [`assets/header-architecture.svg`](./assets/header-architecture.svg)
- [`assets/header-scripts.svg`](./assets/header-scripts.svg)
- [`assets/header-build.svg`](./assets/header-build.svg)
- [`assets/header-troubleshooting.svg`](./assets/header-troubleshooting.svg)
- [`assets/header-repository.svg`](./assets/header-repository.svg)

Treat those SVGs as generated files. Edit the Python generator, rerun it, and
commit the resulting assets together with the script change that produced them.

---

![Build](assets/header-build.svg)

You can run the extractor directly from source, or you can compile a
standalone binary with Bun. The built binary still requires Chrome or Edge on
the target machine, but it does not require Bun or Node.js.

Build from source:

```bash
# Show build usage
bun ./build.js --help

# Build all supported targets
bun ./build.js

# Build a single target
bun ./build.js macos-arm64
bun ./build.js macos-x64
bun ./build.js windows-x64
```

Artifacts land in `dist/`:

| Artifact | Target |
| --- | --- |
| `dist/stream-transcript-extractor-macos-arm64` | Apple Silicon Macs |
| `dist/stream-transcript-extractor-macos-x64` | Intel Macs |
| `dist/stream-transcript-extractor-windows-x64.exe` | Windows x64 |

Set `BUILD_VERSION` if you want a specific embedded version string:

```bash
BUILD_VERSION=1.0.0 bun ./build.js windows-x64
```

<details>
<summary>macOS signing example</summary>

If you distribute the macOS binary to other machines, sign it first so
Gatekeeper does not block it:

```bash
security find-identity -v -p codesigning

cat > signing-entitlements.plist <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>com.apple.security.cs.allow-jit</key>
  <true/>
  <key>com.apple.security.cs.allow-unsigned-executable-memory</key>
  <true/>
  <key>com.apple.security.cs.disable-executable-page-protection</key>
  <true/>
  <key>com.apple.security.cs.allow-dyld-environment-variables</key>
  <true/>
  <key>com.apple.security.cs.disable-library-validation</key>
  <true/>
</dict>
</plist>
PLIST

codesign \
  --entitlements signing-entitlements.plist \
  --deep \
  --force \
  --sign "Developer ID Application: Your Name (TEAMID)" \
  ./dist/stream-transcript-extractor-macos-arm64

codesign --verify --verbose=4 ./dist/stream-transcript-extractor-macos-arm64
```

</details>

---

![Troubleshooting](assets/header-troubleshooting.svg)

Most failures are operational rather than logic bugs. Start with the browser
state, the selected profile, and the mode-specific flow before assuming the
extractor itself is wrong.

- If the browser fails to open correctly, close all Chrome or Edge processes,
  including background or tray processes, then run the extractor again.
- If the extractor says no pages were found, open Microsoft Stream before you
  continue.
- If `network` mode finds no transcript payload, confirm that you left the
  **Transcript** panel closed until the extractor armed capture.
- If `automatic` mode falls back to manual help, rerun with `--debug`. The
  action trace will show whether it found a transcript control, whether the
  panel appeared to open, and whether transcript-like traffic was ever seen.
- If `dom` mode finds zero entries, confirm that the **Transcript** panel is
  already open and populated before extraction begins.
- If the crawl workflow finds zero meetings, confirm that the selected profile
  can open the Stream **Meetings** view and that the page has finished
  loading before discovery starts. If a `*.state.json` file already exists,
  the workflow can still operate on the saved queue.
- If one crawl item fails but the rest succeed, inspect the saved
  `*.state.json` file first, then the matching `*.batch.json` file. They
  record the item URL, the latest queue status, any saved `*.network.json`
  path when `--debug` was enabled, and the final error message.
- If a macOS binary runs locally but fails on another Mac, sign it before
  distribution and notarize it if your distribution model requires that.

---

![Repository layout](assets/header-repository.svg)

The repository stays deliberately small, but the build helper, docs assets,
and maintenance scripts sit alongside the extractor so the project remains
self-explanatory when you share it.

```text
stream-transcript-extractor/
├── README.md
├── assets/
│   ├── banner.svg
│   ├── extractor-architecture.d2
│   ├── extractor-architecture.svg
│   ├── header-overview.svg
│   └── ...
├── build.js
├── extract.js
├── lib/
│   ├── cdp.js
│   ├── cli.js
│   └── network-capture.js
├── mise.toml
└── scripts/
    ├── generate_assets.py
    └── validate_readme.py
```

[`extract.js`](./extract.js) remains the core runtime and CLI entrypoint for
the Bun-based workflow. [`lib/cli.js`](./lib/cli.js) keeps the shared CLI
surface in one place, [`lib/cdp.js`](./lib/cdp.js) keeps the reusable browser
transport and navigation flow in one place, and
[`lib/network-capture.js`](./lib/network-capture.js) keeps the reusable
response-body helpers in one place. [`build.js`](./build.js),
[`scripts/generate_assets.py`](./scripts/generate_assets.py), and
[`scripts/validate_readme.py`](./scripts/validate_readme.py) keep the batch
workflow, binary-release contract, and documentation contract explicit.
