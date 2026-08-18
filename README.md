# OneMore MCP (An MCP server for the OneMore CLI / OneNote)

A Windows system-tray application that bridges an LLM and [OneMore's command-line interface][cli]
for **desktop OneNote**, exposing it to MCP clients (Claude Desktop, Claude Code, Codex, etc.) over a
local **Streamable HTTP** endpoint. Each tool call runs a `OneMoreCli.exe` command and returns its
output — read the notebook hierarchy and pages, search, append notes, add hashtags, export, and more.

The tray app — rather than a stdio MCP server — is a persistent host you start once and leave
running, with a visible status icon and quick access to its config and logs. It also owns the single
OneNote session, so concurrent tool calls are serialised safely.

[cli]: https://onemoreaddin.com/the-basics/OneMore%20CLI.htm

## Requirements

- Windows 10/11
- **Desktop OneNote with the [OneMore add-in][onemore] installed.** OneMore ships the CLI at
  `%ProgramFiles%\River\OneMoreAddIn\OneMoreCli.exe`, which this app drives. (The OneMore CLI docs
  quote an older `River Software\OneMore` path; the app auto-detects both.)
- **To run**: both the [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)
  **and** the ASP.NET Core 10 Runtime (the app hosts its MCP server over Kestrel). Release downloads
  are framework-dependent, so both must be installed first — they are the "Desktop Runtime" and
  "ASP.NET Core Runtime" installers on that download page.
- **To build from source**: the .NET SDK 10 (it includes both runtimes).

[onemore]: https://onemoreaddin.com/

## Install and run (from a release)

No build tools or development experience needed — install two free Microsoft runtimes, then download
and unzip the app.

1. **Install the .NET 10 runtimes.** From the [.NET 10 download page](https://dotnet.microsoft.com/download/dotnet/10.0),
   download and run **both** installers (they're free, and on the same page): the **.NET Desktop
   Runtime** and the **ASP.NET Core Runtime**. Pick **x64** for a normal PC (or **Arm64** for an ARM
   device). You only do this once.

2. **Download the app.** Open the latest release on the
   [Releases page](https://github.com/tridian-tn/OneMoreMcp/releases) and download the zip for your
   CPU (`OneMoreMcp-<version>-win-x64.zip` for most PCs).

3. **Extract and run.** Unzip it to a folder you'll keep (e.g. `C:\Tools\OneMoreMcp`), then
   double-click **`OneMoreMcp.exe`**. A note icon appears in the system tray (near the clock).
   - Windows SmartScreen may warn that the publisher is unknown (the release isn't code-signed).
     Click **More info → Run anyway**.

4. **Check it found OneMore.** Right-click the tray icon; the menu shows the detected
   `OneMoreCli.exe` path. If it says *not found*, set `CliPath` in the config (see [Configure](#configure)).

5. **Connect an LLM.** Follow [Connect an LLM](#connect-an-llm) below.

Optional: right-click the tray icon → **Start with Windows** so it launches automatically at logon.

> Prefer to build it yourself? See [Building from source](#building-from-source) at the end.

## Configure

On first launch the app writes a starter config to:

```
%APPDATA%\OneMoreMcp\config.json
```

Edit it (tray menu → **Open configuration…**). `CliPath`, `AllowWrites`, `DefaultFormat`,
`CommandTimeoutSeconds`, and `ExportRoot` are read live (reloaded on save). `Port`, `UseHttps`, and
`TrustCertificate` are read once at startup, so changing those needs an app restart.

```json
{
  "OneMoreMcp": {
    "Port": 3002,
    "UseHttps": false,
    "TrustCertificate": true,
    "CliPath": "C:\\Program Files\\River\\OneMoreAddIn\\OneMoreCli.exe",
    "AllowWrites": false,
    "DefaultFormat": "markdown",
    "CommandTimeoutSeconds": 150,
    "ExportRoot": ""
  }
}
```

> [!IMPORTANT]
> This is a JSON file, so each backslash in a Windows path must be **doubled**: write
> `C:\\Notes\\Exports`, not `C:\Notes\Exports`. (A single forward slash also works.)

- **Port**: loopback TCP port; the server binds `127.0.0.1`/`::1` only.
- **UseHttps** / **TrustCertificate**: serve over TLS with a self-signed localhost certificate; see
  [HTTPS](#https). Off by default — loopback-only HTTP never leaves your machine.
- **CliPath**: full path to `OneMoreCli.exe`. Defaults to the standard install location shown above
  (`River\OneMoreAddIn`). If that path doesn't exist (or is blank), the app auto-detects it under
  Program Files, Program Files x86, and LocalAppData — trying both the current `River\OneMoreAddIn`
  layout and the older documented `River Software\OneMore` one. Set it explicitly for a non-standard install.
- **AllowWrites**: master gate for content-changing tools (overwrite pages, hashtags, TOC,
  export, cleanup). **Off by default**, so a fresh install is read-only. The append-only
  `append_to_page` tool is exempt — see [Appending vs. writing](#appending-vs-writing).
- **DefaultFormat**: how read tools return page/hierarchy content — `markdown` (default, compact) or
  `xml` (raw OneNote XML). Each read tool can override this per call.
- **SyncAfterWrite**: after a content write (`append_to_page` / `update_page`), sync the
  affected notebook to storage so the change reliably lands and is visible on the next read. On by
  default; best-effort (a failed sync is logged, not surfaced as a write failure).
- **CommandTimeoutSeconds**: how long a single CLI call may run before it's cancelled. OneMore
  operations spanning many pages can take a minute or two, so the default is generous (150s).
- **ExportRoot**: optional folder that `export` output paths must stay within. Blank allows any path.

## HTTPS

HTTPS is off by default: the endpoint is loopback-only, so plain HTTP never leaves your machine. If
you enable it (`"UseHttps": true`), on first run the app:

1. generates a self-signed certificate for `localhost` (SAN: `localhost`, `127.0.0.1`, `::1`), valid
   5 years, persisted at `%APPDATA%\OneMoreMcp\onemoremcp-localhost.pfx`;
2. with `TrustCertificate` on, installs it into your **current-user Trusted Root** store (one-time
   consent prompt) so programs that read the Windows certificate store trust it.

If you skipped the prompt, re-run it any time from the tray: **Trust HTTPS certificate (for Claude)…**.
Node-based clients (Claude Code, and the `mcp-remote` bridge) need one extra setting to honour that
certificate; see [Connect an LLM](#connect-an-llm).

## Connect an LLM

The server speaks MCP over **Streamable HTTP** at the root path: `http(s)://localhost:<Port>/`,
port `3002` by default. It listens on loopback only, so it's reachable from programs on this machine
but not from the network.

### HTTP or HTTPS?

Because the endpoint is loopback-only, **plain HTTP is the simplest option and nothing leaves your
computer**, so it's the default. HTTPS also works, but its self-signed certificate needs an extra
step for Node-based clients, so reach for it only if you specifically want TLS locally.

- **HTTP** (default): use `http://localhost:3002/`.
- **HTTPS**: set `"UseHttps": true`, restart the app, follow the certificate step below, and use
  `https://localhost:3002/`.

### Claude Code

Register the server with the `http` transport, at *user* scope so it's available from every
directory:

```bash
claude mcp add --transport http --scope user onemore http://localhost:3002/
```

Claude Code loads MCP servers at startup, so open a fresh session and confirm:

```bash
claude mcp get onemore        # Status: ✔ Connected
```

The tools (`list_hierarchy`, `get_page`, `append_to_page`, …) are then available in any session.

#### HTTPS with Claude Code: the Node certificate step

Claude Code runs on Node.js, which validates TLS against its own bundled CA list and **ignores the
Windows certificate store by default**. So even with `TrustCertificate: true`, Node rejects the
self-signed certificate (`DEPTH_ZERO_SELF_SIGNED_CERT`). Tell Node to use the Windows store with
`--use-system-ca` (Node 23.8.0+, backported to current v22/v24), then fully restart Claude Code:

```powershell
setx NODE_OPTIONS "--use-system-ca"
```

To avoid certificates altogether, use the `http://` URL instead.

### Claude Desktop

Claude Desktop's **custom connectors** are for *remote* servers, so a `localhost` URL entered there
won't connect. Bridge this local server with [`mcp-remote`][mcp-remote], a small Node proxy, in
`claude_desktop_config.json` (**Settings → Developer → Edit Config**):

```json
{
  "mcpServers": {
    "onemore": {
      "command": "npx",
      "args": ["mcp-remote", "http://localhost:3002/"]
    }
  }
}
```

Because the bridge runs on Node, the **same certificate step as Claude Code** applies if you point it
at an `https://` URL. Restart Claude Desktop after editing the file.

> [!NOTE]
> The Claude Desktop / `mcp-remote` route is the expected setup based on how `mcp-remote` bridges a
> local server; the Claude Code path is the primary tested one.

[mcp-remote]: https://www.npmjs.com/package/mcp-remote

### Codex

In the Codex desktop app: **Settings → Integrations → MCP Servers → Add Server**.
- Transport: **Streamable HTTP** (not STDIO)
- URL: `http://localhost:3002/` (use `https://` if `UseHttps` is enabled)

## Using the tray app

Launch it by double-clicking **`OneMoreMcp.exe`**. A tray icon appears. Right-click for: the server
URL, whether **writes** are enabled, the detected **OneMore CLI** path, copy URL, open configuration,
open log folder, **Start with Windows**, About, and Exit — plus trust HTTPS certificate when HTTPS is
enabled. Logs roll daily under `%APPDATA%\OneMoreMcp\logs`.

### Single instance

Only one copy runs per logged-in user (enforced with a session-local named mutex). Launching it again
just points you at the existing tray icon and exits.

### Start with Windows (run at logon)

Toggle **Start with Windows** in the tray menu. It adds/removes a per-user entry under
`HKCU\Software\Microsoft\Windows\CurrentVersion\Run` (no administrator rights needed). The same can
be scripted:

```bash
OneMoreMcp.exe --enable-autostart
OneMoreMcp.exe --disable-autostart
```

## Tools

| Tool | Purpose | Writes? |
| --- | --- | --- |
| `list_hierarchy` | Notebook/section/page tree (markdown or raw XML). | — |
| `get_page` | A page's content by notebook+section+page, or `current`; markdown (default) or XML. | — |
| `search` | Full-text search a **notebook** (required) for pages matching a query; optional section/page scope. | — |
| `search_titles` | Search page **titles** across a notebook — quicker/more precise for finding a page by name. | — |
| `search_hashtags` | Search OneNote hashtags; optional notebook/section/page scope. | — |
| `sync` | Sync a notebook's pending changes to storage (flush recent edits). | — |
| `append_to_page` | **Append** text (markdown/html/plain) to a page. Fetches, adds, writes back — never sends the page's existing content to the model, never overwrites. | **ungated** |
| `update_page` | Overwrite an **existing** page from raw OneNote page XML (as `get_page` format=xml). Cannot create pages — see below. | gated |
| `create_page` | Create a new page with a title. The page is created **empty** and can't be given body text — see below. | gated |
| `add_hashtag` / `remove_hashtag` | Add/remove hashtags on the current page(s). | gated |
| `export` | Export pages to a folder (HTML/PDF/Word/XML/Markdown/OneNote); confined to `ExportRoot` if set. | gated |
| `archive` | Archive a notebook (or section) to a `.zip`; confined to `ExportRoot` if set. | gated |
| `goto` | Navigate OneNote to a page/object. | — |
| `diagnostics` | Dump OneNote/OneMore diagnostic info (connectivity, versions, paths). | — |
| `run_cleanup` | Page maintenance on a notebook (`applyStyles`, `removeEmpty`, `trim`, `recalculate`, `enableSpellCheck`/`disableSpellCheck`, `embed`, …) via an `operation` argument. | gated |

## Reading & output formats

`list_hierarchy`, `get_page`, and `search` return **compact Markdown by default** — OneNote's native
XML is verbose and namespaced, and sending it verbatim is token-heavy. Pass `format: "xml"` (or set
`DefaultFormat` to `xml`) to get the raw OneNote page/hierarchy XML instead — useful when you intend
to edit it and write it back with `update_page`. The Markdown projection is lossy on
styling by design: it preserves structure and text, not formatting.

## New pages can be created, but not filled

`PutPage` documents a create path — naming a page **without** `--force` *"will attempt to create a
page of that name"* — but a single call doesn't do what it says. Live testing against OneMore 7.3.0
shows the page shell is created while **the title and outline are silently discarded**, leaving an
empty **"Untitled"** page. `create_page` works around this in two steps: it creates the shell, finds
the new page by diffing the section's page IDs (its name is "Untitled", so the name can't identify
it), then applies the title with a second write targeted by the page's ID. That part works.

**Body content can't be added at all.** OneNote ignores a new top-level `one:Outline`, so a page that
has no outline can't be given one — verified on both freshly created and long-standing pages, with
explicit `Position`/`Size`, with a body `quickStyleIndex`, and with `omHash` stripped so nothing was
pruned. `InsertToc` — an unrelated command that builds content — is equally silent. Modifying content
that already exists works fine, which is why `append_to_page` and `update_page` are reliable.

So a page created here stays empty until someone types in it in OneNote; after that, the normal write
tools work on it. Because the CLI has no delete command, failed attempts leave junk "Untitled" pages
behind, so `update_page` checks its target exists before writing rather than producing one.

Two faults in OneMore's own source explain the silence
([`PutPageCommand.cs`](https://github.com/stevencohn/OneMore/blob/main/OneMore/Cli/Commands/PutPageCommand.cs)):

- `PutPageCommand` discards the `bool` returned by `OneNote.Update`, so a failed content update is
  still reported as success.
- `Page.OptimizeForSave` drops any child element whose `omHash` still matches its content — a
  deliberate "unchanged, skip" optimisation which also means **re-sending a stale copy of a page
  silently does nothing**. Always write back XML you've just read.

## Appending vs. writing

There are two ways to change a page, with deliberately different safety models:

- **`append_to_page` is always available**, even when `AllowWrites` is `false`. The model supplies
  only the *new* text; the server fetches the page's XML **locally**, inserts the text as new
  paragraph(s), and writes it back. The existing content never becomes LLM tokens and can never be
  overwritten — only added to. This is the recommended, cheap, low-risk way to have an LLM add notes.
- **Everything else that changes content is gated** behind `AllowWrites` (off by default):
  `update_page` (which *overwrites* a page), `create_page`, `add_hashtag`/`remove_hashtag`,
  `export`, and `run_cleanup`. Enable them by setting `"AllowWrites": true` in the config. Gated
  tools return a clear "writes are disabled" error until you do.

> [!IMPORTANT]
> **Requires OneMore 7.3.0 or newer** (the release that fixed [#2322][2322] via [#2323][2323] +
> [#2324][2324]). This server reads page/hierarchy content via the CLI's `--output <file>` option to
> avoid stdout/console-encoding corruption of non-ASCII content, writes page XML with its `omHash`
> attributes intact, and uses bare-switch booleans — all as of 7.3.0. On 7.2.0 or earlier the read
> commands fail. Content-writes (`append_to_page`, `update_page`, hashtags,
> `run_cleanup`) are **verified working on 7.3.0** (append round-trips and persists).
>
> [2322]: https://github.com/stevencohn/OneMore/issues/2322
> [2323]: https://github.com/stevencohn/OneMore/pull/2323
> [2324]: https://github.com/stevencohn/OneMore/pull/2324

## Concurrency note

OneNote automation is effectively single-threaded, so the app **serialises every CLI invocation**
through one internal lock — two tool calls can't collide on the one OneNote session. A single call
can still take a minute or two for multi-page operations; `CommandTimeoutSeconds` bounds it, and a
cancelled or timed-out call kills the CLI process cleanly.

---

The rest of this document covers the internals and building from source — you don't need any of it to
install and use the app.

## Building from source

```bash
dotnet build
dotnet test
```

Run it straight from the checkout (no need to publish a release first):

```bash
dotnet run --project src/OneMoreMcp.App
```

The app icon (white note lines on a purple rounded square) is drawn in code by `TrayIconFactory`,
the single source of truth. The tray icon is rendered from that code at runtime; the committed
`src/OneMoreMcp.App/Resources/App.ico` is the **executable** icon (Explorer, taskbar, Alt-Tab),
generated from the same drawing so the two always match. Regenerate the asset with:

```bash
OneMoreMcp.exe --write-icon src/OneMoreMcp.App/Resources/App.ico
```

### Versioning

The version is derived from Git on every build by [MinVer](https://github.com/adamralph/minver)
(configured in `Directory.Build.props`). A commit tagged `vX.Y.Z` builds as exactly `X.Y.Z`; any
other commit builds as the next patch pre-release with the commit height (e.g. `0.0.3-alpha.0.5`),
and the commit hash is appended to the informational version. The running build's version shows in
the tray menu under **About OneMore MCP…**. A source tree with no `.git` (e.g. an extracted zip)
falls back to `0.0.0-alpha.0`. To cut a release, tag and push:

```bash
git tag v1.2.3 && git push origin v1.2.3
```

## Solution layout

| Project | Target | Role |
| --- | --- | --- |
| `src/OneMoreMcp.Core` | `net10.0` | Pure, testable engine: CLI argument builder (`OneMoreCommand`), OneNote XML → Markdown projections (`OneNoteContent`), and the append engine (`PageAppender`). No Windows/UI/hosting dependencies. |
| `src/OneMoreMcp.App` | `net10.0-windows` | WinForms tray icon + ASP.NET Core host running the MCP server; the `OneMoreCliRunner` (process orchestration) and the MCP tool surface (`OneMoreTools`). |
| `tests/OneMoreMcp.Core.Tests` | `net10.0` | xUnit suite: argument building, XML transforms, and append round-trips. |
| `tests/OneMoreMcp.App.Tests` | `net10.0-windows` | xUnit suite: write-gating policy, the ungated append path, format selection, and export confinement, via a fake runner. |

### Why a separate Core library

All CLI-argument and XML correctness lives in `OneMoreMcp.Core` with **no** UI or hosting
dependencies, so it's unit-testable in isolation and the tray/host layer is a thin adapter.

### OneMore CLI mapping

The server is a **process orchestrator**: each tool builds a `OneMoreCli.exe` invocation, runs it
(exit `0` = success, `1` = error → surfaced as a tool error with stderr), and returns stdout.

| Tool | `OneMoreCli.exe` command |
| --- | --- |
| `list_hierarchy` | `GetHierarchy [--notebook] [--section] [--books]` (XML; `--books` yields JSON) |
| `get_page` | `GetPage --notebook --section --page` (all three required), or `--current` |
| `search` | `Search --notebook --query [--section] [--page]` (results rendered as page paths) |
| `search_titles` | `SearchTitles --query [--notebook]` (page paths) |
| `search_hashtags` | `SearchHashtags --query [--allTags] [--notebook] [--section] [--page]` |
| `sync` | `Sync --notebook` |
| `append_to_page` | `GetPage` (raw) → local append → `PutPage --infile <temp> --force` → `Sync` (if `SyncAfterWrite`) |
| `create_page` | `GetHierarchy` → `PutPage --notebook --section --page --infile` (no `--force`) → `GetHierarchy` → `Goto --pageId` → `GetPage --current` → `PutPage --infile` (no `--page`, targets the embedded ID) |
| `update_page` | `GetPage` (existence check) → `PutPage --notebook --section --page --infile <temp> --force` → `Sync` (if `SyncAfterWrite`) → `GetPage` (verify) |
| `add_hashtag` / `remove_hashtag` | `AddHashtag --tags` / `RemoveHashtag --tags` |
| `export` | `Export --outpath --format [--pageId] [--backup]` |
| `archive` | `Archive --notebook --outfile [--section]` |
| `goto` | `Goto --pageId [--objectId]` |
| `diagnostics` | `Diagnostics [--windows]` |
| `run_cleanup` | `<op> --notebook [--section] [--page]` — ApplyStyles / RemoveEmpty / Trim / Embed --refresh / … (all require `--notebook`) |

> `PutPage --force` overwrites the page identified by the `ID` carried in the supplied XML. The exact
> targeting of `--section`/`--page` may need tuning against your OneMore version — see the tool
> descriptions and adjust if a write lands somewhere unexpected.
