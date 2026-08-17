using System.ComponentModel;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using ModelContextProtocol.Server;
using OneMoreMcp.Core;

namespace OneMoreMcp.App.Mcp;

/// <summary>
/// The curated MCP tool surface over the OneMore CLI. Read tools return notebook/page content
/// (Markdown by default, raw OneNote XML on request). Content-changing tools are gated behind the
/// <c>AllowWrites</c> setting — except <see cref="AppendToPage"/>, which is append-only and never
/// exposes existing page content, so it is always available.
/// </summary>
[McpServerToolType]
public sealed class OneMoreTools
{
    private readonly IOneMoreRunner _runner;
    private readonly IOptionsMonitor<OneMoreMcpOptions> _options;
    private readonly ILogger<OneMoreTools> _log;

    public OneMoreTools(IOneMoreRunner runner, IOptionsMonitor<OneMoreMcpOptions> options, ILogger<OneMoreTools> log)
    {
        _runner = runner;
        _options = options;
        _log = log;
    }

    // ---------------- Read ----------------

    [McpServerTool(Name = "list_hierarchy")]
    [Description("List the OneNote notebook/section/page hierarchy. Optionally scope to one notebook or section. Returns an indented tree (markdown) or the raw OneNote XML.")]
    public async Task<string> ListHierarchy(
        [Description("Only this notebook (by name). Omit for all notebooks.")] string? notebook = null,
        [Description("Only this section (by name, '/'-delimited for nested sections).")] string? section = null,
        [Description("Include notebooks level only (sections/pages omitted).")] bool booksOnly = false,
        [Description("Output format: 'markdown' (default) or 'xml'.")] string? format = null,
        CancellationToken cancellationToken = default)
    {
        var xml = await ReadContent(OneMoreCommand.GetHierarchy(notebook, section, booksOnly), cancellationToken);
        return AsXml(format) ? xml : OneNoteContent.HierarchyToMarkdown(xml);
    }

    [McpServerTool(Name = "get_page")]
    [Description("Get a page's content. Target it by notebook + section + page name, or set current=true for the page open in OneNote. Returns Markdown (default) or raw OneNote page XML.")]
    public async Task<string> GetPage(
        [Description("Notebook name. Required unless current=true.")] string? notebook = null,
        [Description("Section name ('/'-delimited for nested sections). Required unless current=true.")] string? section = null,
        [Description("Page name (quote-free; spaces are fine). Required unless current=true.")] string? page = null,
        [Description("Use the page currently open in OneNote instead of notebook+section+page.")] bool current = false,
        [Description("Output format: 'markdown' (default) or 'xml'.")] string? format = null,
        CancellationToken cancellationToken = default)
    {
        var xml = await ReadPageXml(notebook, section, page, current, cancellationToken);
        return AsXml(format) ? xml : OneNoteContent.PageToMarkdown(xml);
    }

    [McpServerTool(Name = "search")]
    [Description("Full-text search a notebook for pages matching a query, optionally scoped to a section/page. Returns matching page paths (markdown) or raw XML.")]
    public async Task<string> Search(
        [Description("The search query.")] string query,
        [Description("Notebook to search.")] string notebook,
        [Description("Section to scope to (optional).")] string? section = null,
        [Description("Page to scope to (optional).")] string? page = null,
        [Description("Output format: 'markdown' (default) or 'xml'.")] string? format = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(notebook))
            throw new ArgumentException("A notebook is required for search.", nameof(notebook));
        var xml = await ReadContent(OneMoreCommand.Search(query, notebook, section, page), cancellationToken);
        return AsXml(format) ? xml : OneNoteContent.SearchResultsToMarkdown(xml);
    }

    [McpServerTool(Name = "search_titles")]
    [Description("Search page TITLES across a notebook — faster and more precise than full-text search for finding a page by name. Returns matching page paths (markdown) or raw XML.")]
    public async Task<string> SearchTitles(
        [Description("The title query.")] string query,
        [Description("Notebook to search. Omit to search all loaded notebooks.")] string? notebook = null,
        [Description("Output format: 'markdown' (default) or 'xml'.")] string? format = null,
        CancellationToken cancellationToken = default)
    {
        var xml = await ReadContent(OneMoreCommand.SearchTitles(query, notebook), cancellationToken);
        return AsXml(format) ? xml : OneNoteContent.SearchResultsToMarkdown(xml);
    }

    [McpServerTool(Name = "search_hashtags")]
    [Description("Search OneNote hashtags, optionally scoped to a notebook/section/page. Returns matching results as raw XML.")]
    public async Task<string> SearchHashtags(
        [Description("The hashtag query.")] string query,
        [Description("Require all tags to match (AND) rather than any (OR).")] bool allTags = false,
        [Description("Notebook to scope to (optional).")] string? notebook = null,
        [Description("Section to scope to (optional).")] string? section = null,
        [Description("Page to scope to (optional).")] string? page = null,
        CancellationToken cancellationToken = default) =>
        await ReadContent(OneMoreCommand.SearchHashtags(query, allTags, notebook, section, page), cancellationToken);

    // ---------------- Append (ungated, token-free) ----------------

    [McpServerTool(Name = "append_to_page")]
    [Description("Append text to the end of a OneNote page WITHOUT sending the page's existing content to the model. "
        + "The server fetches the page locally, adds your text as new paragraph(s), and writes it back — it can only add, "
        + "never overwrite. Always available (independent of the write gate). Requires notebook + section + page.")]
    public async Task<string> AppendToPage(
        [Description("The text to append. One paragraph per line.")] string text,
        [Description("Notebook name.")] string notebook,
        [Description("Section name ('/'-delimited for nested sections).")] string section,
        [Description("Page name.")] string page,
        [Description("Format of the text: 'markdown' (default), 'html', or 'plain'.")] string? format = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("There is no text to append.", nameof(text));

        // 1. Fetch the page's raw XML locally — this is NOT returned to the caller.
        var pageXml = await ReadPageXml(notebook, section, page, current: false, cancellationToken);

        // 2. Insert the new text, preserving everything else.
        var updated = PageAppender.Append(pageXml, text, ParseAppendFormat(format));

        // 3. Write it back via PutPage --page --force (overwrites the same page with the added content).
        await PutPageXml(updated, notebook, section, page, cancellationToken);
        _log.LogInformation("Appended text to a OneNote page (notebook='{Notebook}', section='{Section}', page='{Page}').",
            notebook, section, page);
        return "Appended.";
    }

    // ---------------- Write (gated) ----------------

    [McpServerTool(Name = "create_or_update_page")]
    [Description("Create or overwrite a page from raw OneNote page XML (as returned by get_page with format='xml'). "
        + "OVERWRITES the target page. Requires writes to be enabled. For simply adding text, use append_to_page instead.")]
    public async Task<string> CreateOrUpdatePage(
        [Description("The full OneNote page XML (one:Page document).")] string pageXml,
        [Description("Notebook name to place the page in.")] string? notebook = null,
        [Description("Section name to place the page in ('/'-delimited).")] string? section = null,
        [Description("Page name. Omit to use the title/ID in the XML.")] string? page = null,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        if (string.IsNullOrWhiteSpace(pageXml))
            throw new ArgumentException("The page XML is empty.", nameof(pageXml));
        // Fail early on malformed XML with a clear message (rather than a vaguer downstream CLI error).
        try { XDocument.Parse(pageXml); }
        catch (XmlException ex) { throw new ArgumentException("The page XML could not be parsed.", nameof(pageXml), ex); }
        await PutPageXml(pageXml, notebook, section, page, cancellationToken);
        return "Page written.";
    }

    [McpServerTool(Name = "add_hashtag")]
    [Description("Add one or more hashtags to pages in a notebook (optionally scoped to a section/page). Requires writes to be enabled.")]
    public async Task<string> AddHashtag(
        [Description("Space-separated tags, e.g. '#todo #review'.")] string tags,
        [Description("Notebook name to act on.")] string notebook,
        [Description("Section name to scope to (optional).")] string? section = null,
        [Description("Page name to scope to (optional).")] string? page = null,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        await RunChecked(OneMoreCommand.AddHashtag(tags, notebook, section, page), cancellationToken);
        return "Hashtag(s) added.";
    }

    [McpServerTool(Name = "remove_hashtag")]
    [Description("Remove one or more hashtags from pages in a notebook (optionally scoped to a section/page). Requires writes to be enabled.")]
    public async Task<string> RemoveHashtag(
        [Description("Space-separated tags to remove.")] string tags,
        [Description("Notebook name to act on.")] string notebook,
        [Description("Section name to scope to (optional).")] string? section = null,
        [Description("Page name to scope to (optional).")] string? page = null,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        await RunChecked(OneMoreCommand.RemoveHashtag(tags, notebook, section, page), cancellationToken);
        return "Hashtag(s) removed.";
    }

    [McpServerTool(Name = "insert_toc")]
    [Description("Insert (or refresh) a table of contents on a page. Requires notebook + section + page, and writes to be enabled.")]
    public async Task<string> InsertToc(
        [Description("Notebook name.")] string notebook,
        [Description("Section name ('/'-delimited).")] string section,
        [Description("Page name to insert the TOC on.")] string page,
        [Description("Refresh an existing TOC instead of inserting a new one.")] bool refresh = false,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        await RunChecked(OneMoreCommand.InsertToc(notebook, section, page, refresh), cancellationToken);
        return refresh ? "Table of contents refreshed." : "Table of contents inserted.";
    }

    [McpServerTool(Name = "export")]
    [Description("Export pages to a folder in a chosen format (HTML, PDF, Word, XML, Markdown, or OneNote). Writes to disk. Requires writes to be enabled.")]
    public async Task<string> Export(
        [Description("Destination folder path.")] string outpath,
        [Description("Format: HTML, PDF, Word, XML, Markdown, or OneNote.")] string format,
        [Description("Specific page ID to export. Omit to export the current selection.")] string? pageId = null,
        [Description("Produce a backup-style export.")] bool backup = false,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        EnsureWithinExportRoot(outpath);
        await RunChecked(OneMoreCommand.Export(outpath, format, pageId, backup), cancellationToken);
        return $"Exported to {outpath} ({format}).";
    }

    [McpServerTool(Name = "sync")]
    [Description("Sync a notebook's pending changes to storage, ensuring recent edits are flushed and visible to subsequent reads. Always available.")]
    public async Task<string> Sync(
        [Description("Notebook name to sync.")] string notebook,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(notebook))
            throw new ArgumentException("A notebook is required to sync.", nameof(notebook));
        var output = await RunChecked(OneMoreCommand.Sync(notebook), cancellationToken);
        return string.IsNullOrWhiteSpace(output) ? $"Synced '{notebook}'." : output.Trim();
    }

    [McpServerTool(Name = "goto")]
    [Description("Navigate OneNote to a page (and optionally an object within it). Read-only navigation; always available.")]
    public async Task<string> Goto(
        [Description("The page ID to navigate to.")] string pageId,
        [Description("An object ID within the page to focus.")] string? objectId = null,
        CancellationToken cancellationToken = default)
    {
        await RunChecked(OneMoreCommand.Goto(pageId, objectId), cancellationToken);
        return "Navigated.";
    }

    [McpServerTool(Name = "run_cleanup")]
    [Description("Run a page-maintenance operation on a notebook (optionally scoped to a section/page). Requires writes to be enabled. "
        + "operation: applyStyles, clearBackground, removeEmpty, trim, recalculate, removeAuthors, removeInk, removeCitations, "
        + "removeTags, restoreAutosize, enableSpellCheck, disableSpellCheck, embed.")]
    public async Task<string> RunCleanup(
        [Description("The maintenance operation to run (see the list in this tool's description).")] string operation,
        [Description("Notebook to act on.")] string notebook,
        [Description("Section to scope to (optional).")] string? section = null,
        [Description("Page to scope to (optional).")] string? page = null,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        if (string.IsNullOrWhiteSpace(notebook))
            throw new ArgumentException("A notebook is required for cleanup operations.", nameof(notebook));
        var command = ResolveCleanup(operation, notebook, section, page);
        await RunChecked(command, cancellationToken);
        return $"Ran {command.Name}.";
    }

    [McpServerTool(Name = "archive")]
    [Description("Archive a notebook (or a section) to a .zip backup file. Requires writes to be enabled.")]
    public async Task<string> Archive(
        [Description("Notebook to archive.")] string notebook,
        [Description("Destination .zip file path.")] string outfile,
        [Description("Section to archive (optional; omit to archive the whole notebook).")] string? section = null,
        CancellationToken cancellationToken = default)
    {
        EnsureWritesAllowed();
        if (string.IsNullOrWhiteSpace(notebook))
            throw new ArgumentException("A notebook is required to archive.", nameof(notebook));
        if (string.IsNullOrWhiteSpace(outfile))
            throw new ArgumentException("An output file path is required to archive.", nameof(outfile));
        EnsureWithinExportRoot(outfile);
        await RunChecked(OneMoreCommand.Archive(notebook, section, outfile), cancellationToken);
        return $"Archived to {outfile}.";
    }

    [McpServerTool(Name = "diagnostics")]
    [Description("Dump diagnostic information about OneNote and OneMore (connectivity, versions, paths, etc.). Read-only.")]
    public async Task<string> Diagnostics(
        [Description("Include window/layout details.")] bool includeWindows = false,
        CancellationToken cancellationToken = default) =>
        await ReadContent(OneMoreCommand.Diagnostics(includeWindows), cancellationToken);

    // ---------------- Helpers ----------------

    /// <summary>Throws an actionable error when the CLI reported a non-zero exit.</summary>
    private static void EnsureOk(OneMoreCommand command, CliResult result)
    {
        if (!result.Ok)
            throw new InvalidOperationException(
                $"OneMore CLI '{command.Name}' failed (exit {result.ExitCode}): {Trim(result.StdErr, result.StdOut)}");
    }

    /// <summary>Runs a command, throws on a non-zero exit, and returns stdout. Used by write/action tools.</summary>
    private async Task<string> RunChecked(OneMoreCommand command, CancellationToken cancellationToken)
    {
        var result = await _runner.RunAsync(command, cancellationToken);
        EnsureOk(command, result);
        return result.StdOut;
    }

    /// <summary>
    /// Runs a content-producing command with <c>--output &lt;temp&gt;</c> and returns the file's contents.
    /// The CLI writes output straight to the file, avoiding the stdout/console-encoding corruption of
    /// non-ASCII page content that can occur when capturing piped output. Requires a OneMore build with
    /// <c>--output</c> support.
    /// </summary>
    private async Task<string> ReadContent(OneMoreCommand command, CancellationToken cancellationToken)
    {
        var temp = NewTempFile();
        try
        {
            var result = await _runner.RunAsync(command.Output(temp), cancellationToken);
            EnsureOk(command, result);

            if (!File.Exists(temp))
                return result.StdOut;

            var length = new FileInfo(temp).Length;
            if (length > MaxContentBytes)
                throw new InvalidOperationException(
                    $"OneMore CLI '{command.Name}' produced {length / (1024 * 1024)} MB of output, " +
                    $"exceeding the {MaxContentBytes / (1024 * 1024)} MB limit.");

            // The CLI exits 0 even on errors, printing the diagnostic to stdout and leaving the file
            // empty. A real content read is never empty, so an empty file means fall back to stdout.
            var content = await File.ReadAllTextAsync(temp, cancellationToken);
            return string.IsNullOrEmpty(content) ? result.StdOut : content;
        }
        finally
        {
            try { File.Delete(temp); } catch { /* best effort */ }
        }
    }

    private async Task<string> ReadPageXml(string? notebook, string? section, string? page, bool current, CancellationToken cancellationToken)
    {
        if (!current && (string.IsNullOrWhiteSpace(notebook) || string.IsNullOrWhiteSpace(section) || string.IsNullOrWhiteSpace(page)))
            throw new ArgumentException("Specify notebook, section, and page together — or set current=true.");
        return await ReadContent(OneMoreCommand.GetPage(notebook, section, page, current), cancellationToken);
    }

    /// <summary>Writes page XML to a temp file and hands it to PutPage --force, cleaning up afterwards.</summary>
    private async Task PutPageXml(string pageXml, string? notebook, string? section, string? page, CancellationToken cancellationToken)
    {
        // The XML is sent as-is (omHash intact): PutPage accepts it and uses omHash for change detection.
        var temp = NewTempFile();
        await File.WriteAllTextAsync(temp, pageXml, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), cancellationToken);
        try
        {
            // PutPage prints nothing on success but reports schema/other errors on stdout while still
            // exiting 0, so treat any output as a failure rather than falsely reporting success.
            var output = await RunChecked(OneMoreCommand.PutPage(notebook, section, page, temp, force: true), cancellationToken);
            if (!string.IsNullOrWhiteSpace(output))
                throw new InvalidOperationException($"OneMore PutPage did not apply the change: {output.Trim()}");

            // Best-effort flush so the write reliably lands / is visible on the next read.
            if (_options.CurrentValue.SyncAfterWrite && !string.IsNullOrWhiteSpace(notebook))
            {
                try { await RunChecked(OneMoreCommand.Sync(notebook), cancellationToken); }
                // Cancellation isn't a sync failure — let it propagate; only warn on real errors.
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    _log.LogWarning(ex, "Auto-sync after write failed for notebook '{Notebook}'.", notebook);
                }
            }
        }
        finally
        {
            try { File.Delete(temp); } catch { /* best effort */ }
        }
    }

    // Upper bound on a single read's content, mirroring the runner's stdout cap so the --output file
    // read can't load an unbounded payload into memory.
    private const long MaxContentBytes = 64L * 1024 * 1024;

    private static string NewTempFile() =>
        Path.Combine(Path.GetTempPath(), $"onemoremcp_{Guid.NewGuid():N}.xml");

    private void EnsureWritesAllowed()
    {
        if (!_options.CurrentValue.AllowWrites)
            throw new InvalidOperationException(
                "Writes are disabled. Enable them by setting \"AllowWrites\": true in the configuration " +
                "(tray menu → Open configuration…), then restart. (append_to_page works without this.)");
    }

    private void EnsureWithinExportRoot(string outpath)
    {
        var root = _options.CurrentValue.ExportRoot;
        if (string.IsNullOrWhiteSpace(root)) return;

        var fullRoot = Path.GetFullPath(root);
        var fullOut = Path.GetFullPath(outpath);
        var rootWithSep = fullRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!fullOut.Equals(fullRoot, StringComparison.OrdinalIgnoreCase) &&
            !fullOut.StartsWith(rootWithSep, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Export path must be within the configured ExportRoot ('{fullRoot}').");
    }

    private bool AsXml(string? format)
    {
        var chosen = string.IsNullOrWhiteSpace(format) ? _options.CurrentValue.DefaultFormat : format;
        return chosen.Trim().ToLowerInvariant() switch
        {
            "xml" => true,
            "markdown" or "md" or "" => false,
            _ => throw new ArgumentException($"Unknown format '{format}'. Use 'markdown' or 'xml'."),
        };
    }

    private static AppendFormat ParseAppendFormat(string? format) =>
        (format ?? "markdown").Trim().ToLowerInvariant() switch
        {
            "markdown" or "md" or "" => AppendFormat.Markdown,
            "html" => AppendFormat.Html,
            "plain" or "text" => AppendFormat.Plain,
            _ => throw new ArgumentException($"Unknown format '{format}'. Use 'markdown', 'html', or 'plain'."),
        };

    private static OneMoreCommand ResolveCleanup(string operation, string? notebook, string? section, string? page)
    {
        var name = operation.Trim().ToLowerInvariant() switch
        {
            "applystyles" => "ApplyStyles",
            "clearbackground" => "ClearBackground",
            "removeempty" => "RemoveEmpty",
            "trim" => "Trim",
            "recalculate" => "Recalculate",
            "removeauthors" => "RemoveAuthors",
            "removeink" => "RemoveInk",
            "removecitations" => "RemoveCitations",
            "removetags" => "RemoveTags",
            "restoreautosize" => "RestoreAutosize",
            "enablespellcheck" => "EnableSpellCheck",
            "disablespellcheck" => "DisableSpellCheck",
            "embed" => "Embed",
            _ => throw new ArgumentException(
                $"Unknown cleanup operation '{operation}'. Valid: applyStyles, clearBackground, removeEmpty, trim, " +
                "recalculate, removeAuthors, removeInk, removeCitations, removeTags, restoreAutosize, " +
                "enableSpellCheck, disableSpellCheck, embed."),
        };

        var command = OneMoreCommand.Cleanup(name, notebook, section, page);
        // Embed's --refresh switch is required — it's the operation trigger.
        if (name == "Embed") command.Switch("refresh", true);
        return command;
    }

    private static string Trim(string stderr, string stdout)
    {
        var s = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        s = s.Trim();
        return s.Length == 0 ? "(no error detail)" : s;
    }
}
