# Stream Transcript Extractor

![Runtime: Bun](https://img.shields.io/badge/runtime-Bun-black?logo=bun)
![Platforms: macOS and Windows](https://img.shields.io/badge/platforms-macOS%20%7C%20Windows-0f766e)
![Source: single file](https://img.shields.io/badge/source-single--file%20extractor-1d4ed8)

Extract Microsoft Teams recording transcripts from Microsoft Stream using your
signed-in Chrome or Edge profile.

## Overview

This repository keeps the working surface deliberately small. You run
[`extract.js`](./extract.js), you optionally compile it with
[`build.js`](./build.js), and this README documents the workflows and outputs.
That gives you one shareable source file for day-to-day extraction without
turning the project into a multi-file application tree.

The extractor supports two modes behind the same CLI:

- `network` is the default and recommended mode. It captures the Stream
  transcript payload directly, decrypts it when needed, and writes richer
  meeting metadata.
- `dom` is the fallback mode. It extracts entries from the rendered transcript
  panel when you need UI-level debugging or when the network path is not
  usable.

The tool does not use Microsoft Graph transcript APIs. It works through the
browser session you already have.

<!-- prettier-ignore -->
> [!IMPORTANT]
> The extractor only works if the selected Chrome or Edge profile can already
> open the Microsoft Stream recording and transcript.

## Requirements

You need a machine that can already open the recording in Microsoft Stream.
The extractor relies on your real browser profile, so browser access matters
more than any API credentials.

- macOS or Windows
- Bun `1.2.4` or later for the source workflow
- Google Chrome or Microsoft Edge
- A signed-in browser profile that can open the recording and transcript

## Quick start

The fastest path to a working extraction is the default `network` mode. It is
the more reliable path for current Stream recordings because it reads the
payload behind the transcript UI instead of depending on a virtualised DOM.

1. Close Chrome or Edge completely.
2. Run the extractor from the repository root.

```bash
bun ./extract.js
```

3. When the browser opens, navigate to the Stream recording page.
4. Leave the **Transcript** panel closed.
5. Return to the terminal and press Enter so the extractor can attach to the
   page and arm capture.
6. Open the **Transcript** panel, let it load, then return to the terminal and
   press Enter again.

The default output is a JSON file in [`output/`](./output). If you also want
Markdown, add `--format md` or `--format both`.

## Choose a mode

Both modes solve the same user problem, but they optimise for different kinds
of stability and debugging. The single entrypoint is stable. The workflow
changes only after you choose a mode.

### Network mode

`network` mode is the default. It watches Microsoft Stream's in-browser
requests, extracts the transcript payload, and writes a companion
`.network.json` capture file that helps with debugging and cross-browser
comparison.

Use it when you want:

- the most reliable extraction path
- richer meeting metadata, such as creator and recording times
- a saved network capture for debugging
- better parity across macOS, Windows, Chrome, and Edge

Use it like this:

```bash
bun ./extract.js --mode network
bun ./extract.js --mode network --debug
```

In `network` mode, open the recording page first with the **Transcript** panel
closed. Only open the panel after the extractor says capture is armed.

### DOM mode

`dom` mode is the older fallback path. It scrolls the rendered transcript panel
and extracts entries from the visible page DOM.

Use it when you want:

- to debug the actual rendered transcript UI
- a fallback if the network transport changes
- to compare DOM extraction against payload extraction

Use it like this:

```bash
bun ./extract.js --mode dom
bun ./extract.js --mode dom --debug
```

In `dom` mode, open the recording page and the **Transcript** panel before the
extractor begins extraction.

## CLI usage

The CLI is interactive by default, but you can preselect the browser and
profile once you know the stable path for your environment. The public contract
for the tool is the help output from `bun ./extract.js --help`, so the README
mirrors that flow rather than inventing a second interface.

```bash
# Default path
bun ./extract.js

# Force a mode
bun ./extract.js --mode network
bun ./extract.js --mode dom

# Pick a browser
bun ./extract.js --browser chrome
bun ./extract.js --browser edge

# Pick a profile by name, email, or directory name
bun ./extract.js --browser chrome --profile Work
bun ./extract.js --browser edge --profile "user@example.com"

# Choose the output format
bun ./extract.js --format json
bun ./extract.js --format md
bun ./extract.js --format both

# Customize output naming
bun ./extract.js --output weekly-standup
bun ./extract.js --output-dir ./exports

# Enable diagnostics
bun ./extract.js --debug

# Show help or version
bun ./extract.js --help
bun ./extract.js --version
```

The supported options are:

| Option | Purpose |
| --- | --- |
| `--mode <network\|dom>` | Choose the extractor mode. `network` is the default. |
| `--browser <chrome\|edge>` | Use a specific browser instead of choosing interactively. |
| `--profile <query>` | Match a profile by name, email, or directory name. |
| `--output <name>` | Override the output filename prefix. |
| `--output-dir <path>` | Write output files to a custom directory. |
| `--format <json\|md\|both>` | Write JSON, Markdown, or both. |
| `--debug-port <port>` | Force a specific remote debugging port. |
| `--debug` | Write extra diagnostics for the selected mode. |
| `--keep-browser-open` | Leave the launched browser open after extraction. |
| `--help` | Print the CLI help text. |
| `--version` | Print the embedded build version. |

## What gets written

Each run can write JSON, Markdown, or both. In `network` mode, the extractor
also writes a `.network.json` capture file. In `dom` mode, `--debug` writes a
`.debug.json` file. The transcript outputs stay intentionally simple so they
work for review, archive, LLM input, or diffing across browsers and operating
systems.

### JSON transcript output

JSON is the default because it preserves the structured payload. A typical file
looks like this:

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

### Markdown transcript output

Markdown is the compact, LLM-friendly output. It starts with meeting metadata,
then writes each transcript entry as a speaker and timestamp block. The header
includes the page-level and recording-level metadata that Stream exposes,
including speakers, recording start and end time, creator details, and the
source URL.

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

### Mode-specific diagnostics

The diagnostic files are mode-specific. They are intended to make course
correction practical when Stream changes its transport or rendering behaviour.

- `network` mode always writes `*.network.json`
- `network` mode with `--debug` writes larger captures with more request and
  response detail
- `dom` mode with `--debug` writes `*.debug.json`

## How it works

The extractor keeps the user flow small, but there is a fair amount happening
under the hood so the browser session stays usable and the output remains
stable across Chrome, Edge, macOS, and Windows.

![Extractor architecture diagram](./assets/extractor-architecture.svg)

The diagram source lives in
[`assets/extractor-architecture.d2`](./assets/extractor-architecture.d2), and
the rendered asset lives in
[`assets/extractor-architecture.svg`](./assets/extractor-architecture.svg).

Three design decisions matter most:

- The entire public source workflow lives in [`extract.js`](./extract.js), so
  you can share one file around without bundling extra source entrypoints.
- The extractor uses your existing signed-in browser profile instead of asking
  for credentials or app registrations.
- The `network` path is preferred because Stream's transcript UI is
  virtualized, brittle, and more likely to drift across browsers or operating
  systems.

## Build standalone binaries

You can run the extractor directly from source, or you can compile a standalone
binary with Bun. The built binary still requires Chrome or Edge on the target
machine, but it does not require Bun or Node.js. The build path stays in one
small file so the repository remains easy to audit and share.

Use the direct Bun commands:

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

The produced artifacts land in [`dist/`](./dist):

| Artifact | Target |
| --- | --- |
| `dist/stream-transcript-extractor-macos-arm64` | Apple Silicon Macs |
| `dist/stream-transcript-extractor-macos-x64` | Intel Macs |
| `dist/stream-transcript-extractor-windows-x64.exe` | Windows x64 |

Run the artifact directly on the matching operating system:

```bash
# macOS
./dist/stream-transcript-extractor-macos-arm64

# Windows
.\dist\stream-transcript-extractor-windows-x64.exe
```

If you want a specific embedded version string in the compiled binary, set
`BUILD_VERSION` when you build:

```bash
BUILD_VERSION=1.0.0 bun ./build.js windows-x64
```

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

## FAQ

This repository is intentionally small, so a few design choices look unusual at
first glance.

- Why is the build script `build.js`?
  There is no real reason to keep the build script in TypeScript here. The
  build path is plain JavaScript so the repo stays easier to share, audit, and
  run directly with Bun.
- Why is the extractor one large file?
  The goal is easy source sharing. A single `extract.js` file is easier to pass
  around than a small source tree when you want the unpackaged workflow.
- Why not use Playwright CLI?
  We tried that path first, but Stream's transcript UI was brittle at exactly
  the layer Playwright had to drive. The hard problem was nested, virtualised
  sub-DOM scrolling inside the transcript pane, plus reliable browser-profile
  reuse and direct CDP access for the network extractor. After fighting that
  for too long, a small custom wrapper gave us tighter control, fewer
  dependencies, and a better fit for the single-file design.
- Why keep `dom` mode at all?
  It is still useful as a fallback and as a way to debug what Stream renders in
  the page when the transport layer changes.

## Troubleshooting

Most failures are operational rather than logic bugs. Start with the browser
state, the selected profile, and the mode-specific flow before assuming the
extractor itself is wrong.

- If the browser fails to open correctly, close all Chrome or Edge processes,
  including background or tray processes, then run the extractor again.
- If the extractor says no pages were found, open Microsoft Stream before you
  continue.
- If `network` mode finds no transcript payload, confirm that you left the
  **Transcript** panel closed until the extractor armed capture.
- If `dom` mode finds zero entries, confirm that the **Transcript** panel is
  already open and populated before extraction begins.
- If the browser stays open after extraction, rerun with the latest code. The
  current shutdown path first tries `Browser.close` over CDP, then falls back
  to process termination.
- If a macOS binary runs locally but fails on another Mac, sign it before
  distribution and consider notarization if your distribution model requires
  it.

## Repository layout

The repository is deliberately small. Only one file is required for the source
workflow, but the build helper, documentation, and architecture assets stay
alongside it so the project is still self-explanatory when shared.

```text
stream-transcript-extractor/
|- .gitignore
|- README.md
|- assets/
|  |- extractor-architecture.d2
|  `- extractor-architecture.svg
|- build.js
|- extract.js
`- mise.toml
```

The only source file you need to share to run the extractor from Bun is
[`extract.js`](./extract.js). The [`assets/`](./assets) folder stores the D2
source and rendered SVG for the system diagram, and [`mise.toml`](./mise.toml)
declares Bun for repo-local tool installation.
