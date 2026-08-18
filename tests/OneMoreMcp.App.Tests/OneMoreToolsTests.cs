using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using ModelContextProtocol;
using OneMoreMcp.App;
using OneMoreMcp.App.Mcp;
using OneMoreMcp.Core;

namespace OneMoreMcp.App.Tests;

/// <summary>
/// Covers the tool layer's policy — write gating, the ungated append path, format selection, and
/// export confinement — with a fake runner so no real OneMore CLI or OneNote is needed.
/// </summary>
public class OneMoreToolsTests
{
    private const string SamplePageXml =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\">" +
        "<one:Title><one:OE><one:T><![CDATA[Title]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[Existing]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline></one:Page>";

    // Distinct from SamplePageXml (the fake's default "current page" content) so write tests that need a
    // real before/after difference — for the silent-no-op read-back check — don't false-positive on writing
    // content that's byte-identical to what's already "on the page".
    private const string SampleUpdatedPageXml =
        "<?xml version=\"1.0\"?>" +
        "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P1}\">" +
        "<one:Title><one:OE><one:T><![CDATA[Title]]></one:T></one:OE></one:Title>" +
        "<one:Outline><one:OEChildren>" +
        "<one:OE><one:T><![CDATA[Updated]]></one:T></one:OE>" +
        "</one:OEChildren></one:Outline></one:Page>";

    private static (OneMoreTools tools, FakeRunner runner) Build(
        bool allowWrites = false, string? exportRoot = null, bool syncAfterWrite = true)
    {
        var options = new OneMoreMcpOptions
        {
            AllowWrites = allowWrites,
            ExportRoot = exportRoot ?? "",
            SyncAfterWrite = syncAfterWrite,
        };
        var runner = new FakeRunner();
        var tools = new OneMoreTools(runner, new StubMonitor(options), NullLogger<OneMoreTools>.Instance);
        return (tools, runner);
    }

    [Fact]
    public async Task Create_or_update_page_is_refused_when_writes_disabled()
    {
        var (tools, runner) = Build(allowWrites: false);
        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage(SamplePageXml, notebook: "N", section: "S", page: "P"));
        Assert.Contains("Writes are disabled", ex.Message);
        Assert.Empty(runner.Commands); // nothing was executed
    }

    [Fact]
    public async Task Create_or_update_page_runs_putpage_when_writes_enabled()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.CreateOrUpdatePage(SampleUpdatedPageXml, notebook: "N", section: "S", page: "P");
        Assert.Contains(runner.Commands, c => c.Name == "PutPage");
    }

    [Fact]
    public async Task Create_or_update_page_requires_a_notebook()
    {
        // PutPage marks --notebook and --section required; omitting either drops the CLI into an
        // interactive prompt the runner can only catch via an output-overflow kill. Reject upfront.
        var (tools, runner) = Build(allowWrites: true);
        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage(SamplePageXml, notebook: "  ", section: "S", page: "P"));
        Assert.Contains("notebook", ex.Message);
        Assert.Empty(runner.Commands); // the CLI was never invoked
    }

    [Fact]
    public async Task Create_or_update_page_requires_a_section()
    {
        var (tools, runner) = Build(allowWrites: true);
        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage(SamplePageXml, notebook: "N", section: "  ", page: "P"));
        Assert.Contains("section", ex.Message);
        Assert.Empty(runner.Commands);
    }

    [Fact]
    public async Task Append_to_page_requires_a_section()
    {
        // append_to_page shares PutPageXml, so the same required-argument guard must apply there.
        var (tools, _) = Build();
        await Assert.ThrowsAsync<McpException>(
            () => tools.AppendToPage("note", notebook: "N", section: "  ", page: "P", format: "plain"));
    }

    [Fact]
    public async Task Append_to_page_works_when_writes_disabled_and_round_trips_locally()
    {
        var (tools, runner) = Build(allowWrites: false, syncAfterWrite: false); // focus on the read/write round-trip
        var result = await tools.AppendToPage("New note", notebook: "N", section: "S", page: "P", format: "plain");

        Assert.Equal("Appended.", result);
        // It reads the page, writes it back, then reads it again to verify the write took effect.
        Assert.Equal(new[] { "GetPage", "PutPage", "GetPage" }, runner.Commands.Select(c => c.Name).ToArray());
        // The written XML preserved the existing content and added the new line.
        Assert.Contains("Existing", runner.LastInfileXml);
        Assert.Contains("New note", runner.LastInfileXml);
    }

    [Fact]
    public async Task Append_to_page_requires_a_target()
    {
        var (tools, _) = Build();
        await Assert.ThrowsAsync<McpException>(
            () => tools.AppendToPage("text", notebook: "", section: "", page: "", format: "plain"));
    }

    [Fact]
    public async Task Get_page_returns_markdown_by_default_and_xml_on_request()
    {
        var (tools, _) = Build();
        var md = await tools.GetPage(notebook: "N", section: "S", page: "P", current: false, format: null);
        Assert.StartsWith("# Title", md);

        var xml = await tools.GetPage(notebook: "N", section: "S", page: "P", current: false, format: "xml");
        Assert.Contains("one:Page", xml);
    }

    [Fact]
    public async Task Export_is_refused_outside_the_configured_root()
    {
        var (tools, _) = Build(allowWrites: true, exportRoot: @"C:\Exports");
        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.Export(outpath: @"C:\Elsewhere\out", format: "PDF"));
        Assert.Contains("ExportRoot", ex.Message);
    }

    [Fact]
    public async Task Add_hashtag_is_gated_by_writes()
    {
        var (tools, _) = Build(allowWrites: false);
        await Assert.ThrowsAsync<McpException>(() => tools.AddHashtag("#todo", notebook: "N"));
    }

    [Fact]
    public async Task Create_or_update_page_sends_omHash_bearing_xml_unchanged()
    {
        var (tools, runner) = Build(allowWrites: true);
        const string xml =
            "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\" ID=\"{P}\">" +
            "<one:PageSettings omHash=\"ABC123\" />" +
            "<one:Outline omHash=\"DEF456\"><one:OEChildren>" +
            "<one:OE><one:T><![CDATA[x]]></one:T></one:OE></one:OEChildren></one:Outline></one:Page>";

        await tools.CreateOrUpdatePage(xml, notebook: "N", section: "S", page: "P");

        // The whole point of the 7.3.0 change: omHash flows through to PutPage unstripped.
        Assert.Contains("omHash=\"ABC123\"", runner.LastInfileXml);
        Assert.Contains("omHash=\"DEF456\"", runner.LastInfileXml);
    }

    [Fact]
    public async Task Create_or_update_page_rejects_malformed_xml()
    {
        var (tools, _) = Build(allowWrites: true);
        await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage("<one:Page", notebook: "N", section: "S", page: "P"));
    }

    [Fact]
    public async Task Get_page_reads_from_the_output_file_not_stdout()
    {
        var (tools, runner) = Build();
        runner.Content["GetPage"] =
            "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">FROM_FILE</one:Page>";
        runner.Stdout = "FROM_STDOUT_SHOULD_BE_IGNORED";

        var xml = await tools.GetPage(notebook: "N", section: "S", page: "P", current: false, format: "xml");

        Assert.Contains("FROM_FILE", xml);
        Assert.DoesNotContain("FROM_STDOUT_SHOULD_BE_IGNORED", xml);
        Assert.Contains("--output", runner.Commands.Single(c => c.Name == "GetPage").Build());
    }

    [Fact]
    public async Task Read_content_falls_back_to_stdout_when_the_output_file_is_empty()
    {
        var (tools, runner) = Build();
        runner.Content["GetPage"] = "";                       // CLI wrote an empty --output file
        runner.Stdout = "No page is currently active.";       // …and put the diagnostic on stdout

        var result = await tools.GetPage(notebook: "N", section: "S", page: "P", current: false, format: "xml");

        Assert.Equal("No page is currently active.", result);
    }

    [Fact]
    public async Task Read_content_throws_when_the_cli_reports_failure()
    {
        var (tools, runner) = Build();
        runner.ExitCode = 1;

        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.GetPage(notebook: "N", section: "S", page: "P", current: false, format: "xml"));
        Assert.Contains("GetPage", ex.Message);
    }

    [Fact]
    public async Task List_hierarchy_routes_through_output_and_renders_markdown()
    {
        var (tools, runner) = Build();
        runner.Content["GetHierarchy"] =
            "<one:Notebooks xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\">" +
            "<one:Notebook name=\"Work\"><one:Section name=\"Ideas\" /></one:Notebook></one:Notebooks>";

        var md = await tools.ListHierarchy(format: null);

        Assert.Contains("- Work", md);
        Assert.Contains("  - Ideas", md);
        Assert.Contains("--output", runner.Commands.Single(c => c.Name == "GetHierarchy").Build());
    }

    [Fact]
    public async Task Search_hashtags_returns_raw_output_file_content()
    {
        var (tools, runner) = Build();
        runner.Content["SearchHashtags"] = "<one:Hits>raw</one:Hits>";

        var result = await tools.SearchHashtags("#todo", allTags: true);

        Assert.Equal("<one:Hits>raw</one:Hits>", result);
        var argv = runner.Commands.Single(c => c.Name == "SearchHashtags").Build();
        Assert.Contains("--allTags", argv);   // bare switch, end to end
        Assert.Contains("--output", argv);
    }

    [Fact]
    public async Task Write_and_action_commands_do_not_use_output()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.CreateOrUpdatePage(SampleUpdatedPageXml, notebook: "N", section: "S", page: "P");
        await tools.AddHashtag("#x", notebook: "N");
        await tools.RunCleanup("trim", notebook: "N");

        // Writes/actions must keep stdout capture, so their success/error detection still works.
        foreach (var name in new[] { "PutPage", "AddHashtag", "Trim" })
            Assert.DoesNotContain("--output", runner.Commands.Single(c => c.Name == name).Build());
    }

    [Fact]
    public async Task Put_page_treats_non_empty_output_as_a_failure()
    {
        var (tools, runner) = Build(allowWrites: true);
        runner.Content["PutPage"] = "schema error";   // PutPage prints to stdout while exiting 0

        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage(SamplePageXml, notebook: "N", section: "S", page: "P"));
        Assert.Contains("did not apply", ex.Message);
    }

    [Fact]
    public async Task Append_to_page_throws_when_the_write_silently_does_not_take_effect()
    {
        // PutPage exits clean (no stdout) but the page is unchanged afterwards — e.g. a OneDrive sync
        // reverting the change post-write. The read-back check must catch this, not report "Appended.".
        var (tools, runner) = Build(syncAfterWrite: false);
        runner.SimulateNoOpWrite = true;

        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.AppendToPage("note", notebook: "N", section: "S", page: "P", format: "plain"));
        Assert.Contains("unchanged", ex.Message);
    }

    [Fact]
    public async Task Create_or_update_page_throws_when_the_write_silently_does_not_take_effect()
    {
        var (tools, runner) = Build(allowWrites: true, syncAfterWrite: false);
        runner.SimulateNoOpWrite = true;

        var ex = await Assert.ThrowsAsync<McpException>(
            () => tools.CreateOrUpdatePage(SampleUpdatedPageXml, notebook: "N", section: "S", page: "P"));
        Assert.Contains("unchanged", ex.Message);
    }

    [Fact]
    public async Task Create_or_update_page_skips_verification_when_no_page_is_named()
    {
        // The create path: with no page name the CLI derives the page from the XML's title, so there's no
        // name to read back by and the no-op check is skipped rather than false-flagging the creation.
        var (tools, runner) = Build(allowWrites: true, syncAfterWrite: false);
        runner.SimulateNoOpWrite = true;

        await tools.CreateOrUpdatePage(SampleUpdatedPageXml, notebook: "N", section: "S");

        Assert.Contains(runner.Commands, c => c.Name == "PutPage");
    }

    [Fact]
    public async Task Search_titles_routes_through_output_and_lists_page_paths()
    {
        var (tools, runner) = Build();
        runner.Content["SearchTitles"] =
            "<Results query=\"q\" count=\"1\"><Page id=\"{1}\" path=\"Work/Ideas/Roadmap\" /></Results>";

        var md = await tools.SearchTitles("Roadmap", notebook: "Work");

        Assert.Contains("- Work/Ideas/Roadmap", md);
        Assert.Contains("--output", runner.Commands.Single(c => c.Name == "SearchTitles").Build());
    }

    [Fact]
    public async Task Search_renders_result_paths_and_passes_the_notebook()
    {
        var (tools, runner) = Build();
        runner.Content["Search"] =
            "<Results query=\"q\" count=\"1\"><Page id=\"{1}\" path=\"Work/Ideas/Roadmap\" /></Results>";

        var md = await tools.Search("Roadmap", notebook: "Work");

        Assert.Contains("- Work/Ideas/Roadmap", md);
        Assert.Contains("--notebook", runner.Commands.Single(c => c.Name == "Search").Build());
    }

    [Fact]
    public async Task Sync_runs_the_sync_command_and_returns_its_output()
    {
        var (tools, runner) = Build();
        runner.Content["Sync"] = "Synced 1 notebook(s): Work";

        var result = await tools.Sync("Work");

        Assert.Equal("Synced 1 notebook(s): Work", result);
        Assert.Equal(new[] { "Sync", "--notebook", "Work" }, runner.Commands.Single(c => c.Name == "Sync").Build());
    }

    [Fact]
    public async Task Search_requires_a_notebook()
    {
        var (tools, _) = Build();
        await Assert.ThrowsAsync<McpException>(() => tools.Search("q", notebook: "  "));
    }

    [Fact]
    public async Task Sync_requires_a_notebook()
    {
        var (tools, _) = Build();
        await Assert.ThrowsAsync<McpException>(() => tools.Sync(""));
    }

    [Fact]
    public async Task Append_auto_syncs_the_notebook_when_enabled()
    {
        var (tools, runner) = Build(syncAfterWrite: true);
        await tools.AppendToPage("note", notebook: "N", section: "S", page: "P", format: "plain");

        // GetPage -> PutPage -> Sync -> GetPage (the last read-back verifies the write took effect).
        Assert.Equal(new[] { "GetPage", "PutPage", "Sync", "GetPage" }, runner.Commands.Select(c => c.Name).ToArray());
        Assert.Contains("N", runner.Commands.Single(c => c.Name == "Sync").Build());
    }

    [Fact]
    public async Task Append_does_not_sync_when_disabled()
    {
        var (tools, runner) = Build(syncAfterWrite: false);
        await tools.AppendToPage("note", notebook: "N", section: "S", page: "P", format: "plain");

        Assert.DoesNotContain(runner.Commands, c => c.Name == "Sync");
    }

    [Fact]
    public async Task Create_or_update_page_auto_syncs_when_enabled()
    {
        // Auto-sync lives in the shared PutPageXml, so it must apply to create_or_update_page too.
        var (tools, runner) = Build(allowWrites: true, syncAfterWrite: true);
        await tools.CreateOrUpdatePage(SampleUpdatedPageXml, notebook: "N", section: "S", page: "P");

        // GetPage (before-snapshot) -> PutPage -> Sync -> GetPage (verifies the write took effect).
        Assert.Equal(new[] { "GetPage", "PutPage", "Sync", "GetPage" }, runner.Commands.Select(c => c.Name).ToArray());
    }

    [Fact]
    public async Task Run_cleanup_scopes_to_notebook_and_maps_the_operation()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.RunCleanup("removeEmpty", notebook: "N", section: "S");

        var argv = runner.Commands.Single(c => c.Name == "RemoveEmpty").Build();
        Assert.Equal(new[] { "RemoveEmpty", "--notebook", "N", "--section", "S" }, argv);
    }

    [Fact]
    public async Task Run_cleanup_embed_adds_the_refresh_switch()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.RunCleanup("embed", notebook: "N");

        var argv = runner.Commands.Single(c => c.Name == "Embed").Build();
        Assert.Equal(new[] { "Embed", "--notebook", "N", "--refresh" }, argv);
    }

    [Fact]
    public async Task Run_cleanup_requires_a_notebook()
    {
        var (tools, _) = Build(allowWrites: true);
        await Assert.ThrowsAsync<McpException>(() => tools.RunCleanup("trim", notebook: ""));
    }

    [Fact]
    public async Task Run_cleanup_rejects_an_unknown_operation()
    {
        var (tools, _) = Build(allowWrites: true);
        await Assert.ThrowsAsync<McpException>(() => tools.RunCleanup("bogus", notebook: "N"));
    }

    [Fact]
    public async Task Archive_is_refused_when_writes_disabled()
    {
        var (tools, _) = Build(allowWrites: false);
        await Assert.ThrowsAsync<McpException>(
            () => tools.Archive("N", outfile: @"C:\Backups\n.zip"));
    }

    [Fact]
    public async Task Archive_is_refused_outside_the_export_root()
    {
        var (tools, _) = Build(allowWrites: true, exportRoot: @"C:\Backups");
        await Assert.ThrowsAsync<McpException>(
            () => tools.Archive("N", outfile: @"C:\Elsewhere\n.zip"));
    }

    [Fact]
    public async Task Archive_requires_notebook_and_outfile()
    {
        var (tools, _) = Build(allowWrites: true);
        await Assert.ThrowsAsync<McpException>(() => tools.Archive("", outfile: @"C:\Backups\n.zip"));
        await Assert.ThrowsAsync<McpException>(() => tools.Archive("N", outfile: "  "));
    }

    [Fact]
    public async Task Archive_runs_when_allowed()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.Archive("N", outfile: @"C:\Backups\n.zip", section: "S");
        Assert.Equal(
            new[] { "Archive", "--notebook", "N", "--section", "S", "--outfile", @"C:\Backups\n.zip" },
            runner.Commands.Single(c => c.Name == "Archive").Build());
    }

    [Fact]
    public async Task Diagnostics_returns_the_output_file_content()
    {
        var (tools, runner) = Build();
        runner.Content["Diagnostics"] = "{\"onenote\":\"ok\"}";

        var result = await tools.Diagnostics();

        Assert.Equal("{\"onenote\":\"ok\"}", result);
        Assert.Contains("--output", runner.Commands.Single(c => c.Name == "Diagnostics").Build());
    }

    // --- Test doubles ---

    private sealed class FakeRunner : IOneMoreRunner
    {
        public List<OneMoreCommand> Commands { get; } = new();
        public string LastInfileXml { get; private set; } = "";

        /// <summary>Per-command payload — written to the --output file, or returned as stdout when no --output.</summary>
        public Dictionary<string, string> Content { get; } = new();

        /// <summary>Stdout returned when a command uses --output (the CLI's diagnostic channel).</summary>
        public string Stdout { get; set; } = "";

        /// <summary>Exit code the fake reports.</summary>
        public int ExitCode { get; set; } = 0;

        /// <summary>What a subsequent GetPage returns once a PutPage has "landed" — simulates real persistence
        /// so tests exercise the same before/after read-back the app does. Null until the first PutPage.</summary>
        public string? PersistedPageXml { get; private set; }

        /// <summary>When true, PutPage exits clean but never updates <see cref="PersistedPageXml"/> — simulates
        /// the silent no-op write from issue #10 (e.g. a OneDrive sync reverting the change).</summary>
        public bool SimulateNoOpWrite { get; set; }

        public string? TryResolveCliPath() => "fake-onemore.exe";

        public Task<CliResult> RunAsync(OneMoreCommand command, CancellationToken cancellationToken = default)
        {
            Commands.Add(command);
            var argv = command.Build();

            // Capture the XML handed to PutPage via --infile so tests can assert on it, and simulate the
            // write landing so a follow-up GetPage reflects it (unless a no-op write is being simulated).
            if (command.Name == "PutPage")
            {
                var infile = OptionValue(argv, "--infile");
                if (infile != null) LastInfileXml = File.ReadAllText(infile);
                if (!SimulateNoOpWrite) PersistedPageXml = LastInfileXml;
            }

            var payload = Content.TryGetValue(command.Name, out var c) ? c
                : command.Name == "GetPage" ? (PersistedPageXml ?? SamplePageXml) : "";

            // With --output the CLI writes the payload to the file and stdout carries only diagnostics;
            // without it, the payload comes back on stdout (write/action commands).
            var outputFile = OptionValue(argv, "--output");
            if (outputFile != null)
            {
                File.WriteAllText(outputFile, payload);
                return Task.FromResult(new CliResult(ExitCode, Stdout, ""));
            }

            return Task.FromResult(new CliResult(ExitCode, payload, ""));
        }

        private static string? OptionValue(IReadOnlyList<string> argv, string flag)
        {
            var i = argv.ToList().IndexOf(flag);
            return i >= 0 && i + 1 < argv.Count ? argv[i + 1] : null;
        }
    }

    private sealed class StubMonitor : IOptionsMonitor<OneMoreMcpOptions>
    {
        public StubMonitor(OneMoreMcpOptions value) => CurrentValue = value;
        public OneMoreMcpOptions CurrentValue { get; }
        public OneMoreMcpOptions Get(string? name) => CurrentValue;
        public IDisposable? OnChange(Action<OneMoreMcpOptions, string?> listener) => null;
    }
}
