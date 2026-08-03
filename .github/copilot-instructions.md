# Copilot instructions — OneMore MCP

A Windows system-tray application that bridges an LLM and the **OneMore CLI** (`OneMoreCli.exe`) for
desktop OneNote, hosting an MCP server over a local (loopback) HTTP endpoint. Each MCP tool call builds
a `OneMoreCli.exe` invocation, runs it, and returns the output. Apply the rules below when reviewing or
generating code.

## Architecture

- **`src/OneMoreMcp.Core`** (`net10.0`) — the pure engine. **No UI, hosting, or Windows dependencies**,
  so it stays unit-testable. Holds `OneMoreCommand` (the CLI argument builder), `OneNoteContent`
  (OneNote XML → Markdown projections), `PageAppender` (append text into page XML), and
  `CliResult` / `CliException`.
- **`src/OneMoreMcp.App`** (`net10.0-windows`) — WinForms tray icon + ASP.NET Core (Kestrel) host
  running the MCP server, plus `OneMoreCliRunner` (process orchestration), the MCP tools
  (`Mcp/OneMoreTools`), options, HTTPS/certificate handling, single-instance, and run-at-logon.
- **`tests/OneMoreMcp.Core.Tests`** (`net10.0`, xUnit) — argument building, XML transforms, and append
  round-trips. **`tests/OneMoreMcp.App.Tests`** (`net10.0-windows`) — tool policy (write
  gating, the ungated append path, export confinement) via a fake runner.

Keep `Core` free of UI/host/Windows references. New CLI-argument or XML behaviour belongs in `Core`,
with tests. The tray/host layer is a thin adapter.

## Top priority: OneMore CLI contract fidelity

The real CLI diverges from its documentation in ways that silently break writes. These were verified
against `OneMoreCli.exe` directly; **guard them, and flag any change that reintroduces them**:

- **Boolean flags are bare switches** (OneMore 7.3.0+): presence enables them, absence uses the default.
  Emit `--force` (via `OneMoreCommand.Switch`, which adds `--flag` only when true), never `--force yes`.
  (7.2.0 briefly required `--flag yes`; that's gone.)
- **Supply every REQUIRED parameter.** `GetPage`/`PutPage` need `--notebook` + `--section` + `--page`
  (unless `--current`); `AddHashtag`/`RemoveHashtag`/`InsertToc` need `--notebook`. A missing required
  parameter makes the CLI drop into an **interactive prompt that re-prompts forever** with no console,
  emitting unbounded output. `OneMoreCliRunner` guards this (closes child stdin, caps captured output,
  kills on overflow) — do not remove that guard, and keep command factories complete.
- **Read content via `--output <file>`, not stdout.** Content commands (`GetPage`, `GetHierarchy`,
  `Search`, `SearchHashtags`) are run through `ReadContent`, which appends `--output <temp>` and reads
  the file — the CLI's stdout capture corrupts non-ASCII content. Writes (`PutPage`, etc.) keep stdout
  capture for success/error detection. Page XML is sent to `PutPage` **with `omHash` intact** (the CLI
  accepts it and uses it for change detection); do not strip it.
- **The CLI exits 0 even on errors** (the message goes to stdout). Do not gate success on the exit code
  alone: reads return stdout; `PutPage` prints nothing on success, so any non-empty `PutPage` output is
  treated as a failure. Preserve that check.
- **Append is ungated; everything else that writes is gated.** `append_to_page` is append-only — it
  fetches the page locally, adds text, and writes back, never exposing existing content and never
  overwriting — so it is exempt from `AllowWrites`. All other content-changing tools
  (`create_or_update_page`, hashtags, `insert_toc`, `run_cleanup`, `export`) must call
  `EnsureWritesAllowed()`. Keep that distinction.
- **Known upstream limitation:** OneMore content-writes (`UpdatePageContent`) currently don't persist —
  tracked as stevencohn/OneMore#2322. Don't attempt to "fix" this in the wrapper; it's external. The
  write tools stay in place to work once it's resolved.

## Host and safety conventions

- **Loopback only.** Kestrel binds `127.0.0.1` / `::1`. Never introduce a non-loopback binding.
- **Serialise CLI calls.** OneNote automation is effectively single-threaded, so all invocations go
  through the runner's single `SemaphoreSlim(1,1)`. Do not add parallelism around the CLI.
- **CLI discovery.** `OneMoreCliRunner` uses the configured `CliPath` when it exists, else auto-detects
  under Program Files / Program Files x86 / LocalAppData, trying both the `River\OneMoreAddIn` and the
  older documented `River Software\OneMore` layouts. Keep the fallback.
- **One-shot commands** (`--enable-autostart`, `--disable-autostart`, `--write-icon`) run before the
  single-instance mutex and host start, then exit. Preserve that ordering.
- **XML output is Markdown by default**, raw XML on request (protects the LLM context window). Keep the
  read tools defaulting to summarised Markdown.

## Code style

- File-scoped namespaces, target-typed `new`, expression-bodied members.
- Nullable reference types enabled; avoid the null-forgiving `!` — prefer an explicit guard.
- Single-line `if`/`else` bodies are unbraced (`if (foo) return null;`).
- Method-header comments use XML doc-comment format (`/// <summary>`), not block comments.
- No spec/ticket references in code comments; describe the *what* and *why* of the code itself.

## Build and test

```
dotnet build
dotnet test
```

Every change must keep both suites green. Any change to a CLI command's arguments must add or update a
`OneMoreCommandTests` assertion (they pin the exact argv), and tool-policy changes belong in
`OneMoreToolsTests`.
