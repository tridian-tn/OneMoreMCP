using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
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

    private static (OneMoreTools tools, FakeRunner runner) Build(bool allowWrites = false, string? exportRoot = null)
    {
        var options = new OneMoreMcpOptions { AllowWrites = allowWrites, ExportRoot = exportRoot ?? "" };
        var runner = new FakeRunner();
        var tools = new OneMoreTools(runner, new StubMonitor(options), NullLogger<OneMoreTools>.Instance);
        return (tools, runner);
    }

    [Fact]
    public async Task Create_or_update_page_is_refused_when_writes_disabled()
    {
        var (tools, runner) = Build(allowWrites: false);
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            () => tools.CreateOrUpdatePage(SamplePageXml, section: "S", page: "P"));
        Assert.Contains("Writes are disabled", ex.Message);
        Assert.Empty(runner.Commands); // nothing was executed
    }

    [Fact]
    public async Task Create_or_update_page_runs_putpage_when_writes_enabled()
    {
        var (tools, runner) = Build(allowWrites: true);
        await tools.CreateOrUpdatePage(SamplePageXml, section: "S", page: "P");
        Assert.Contains(runner.Commands, c => c.Name == "PutPage");
    }

    [Fact]
    public async Task Append_to_page_works_when_writes_disabled_and_round_trips_locally()
    {
        var (tools, runner) = Build(allowWrites: false);
        var result = await tools.AppendToPage("New note", format: "plain", notebook: "N", section: "S", page: "P", current: false);

        Assert.Equal("Appended.", result);
        // It reads the page then writes it back — GetPage before PutPage.
        Assert.Equal(new[] { "GetPage", "PutPage" }, runner.Commands.Select(c => c.Name).ToArray());
        // The written XML preserved the existing content and added the new line.
        Assert.Contains("Existing", runner.LastInfileXml);
        Assert.Contains("New note", runner.LastInfileXml);
    }

    [Fact]
    public async Task Append_to_page_requires_a_target()
    {
        var (tools, _) = Build();
        await Assert.ThrowsAsync<ArgumentException>(
            () => tools.AppendToPage("text", format: "plain", notebook: null, section: null, page: null, current: false));
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
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(
            () => tools.Export(outpath: @"C:\Elsewhere\out", format: "PDF"));
        Assert.Contains("ExportRoot", ex.Message);
    }

    [Fact]
    public async Task Add_hashtag_is_gated_by_writes()
    {
        var (tools, _) = Build(allowWrites: false);
        await Assert.ThrowsAsync<InvalidOperationException>(() => tools.AddHashtag("#todo"));
    }

    // --- Test doubles ---

    private sealed class FakeRunner : IOneMoreRunner
    {
        public List<OneMoreCommand> Commands { get; } = new();
        public string LastInfileXml { get; private set; } = "";

        public string? TryResolveCliPath() => "fake-onemore.exe";

        public Task<CliResult> RunAsync(OneMoreCommand command, CancellationToken cancellationToken = default)
        {
            Commands.Add(command);
            var argv = command.Build();

            // Capture the XML handed to PutPage via --infile so tests can assert on it.
            if (command.Name == "PutPage")
            {
                var i = argv.ToList().IndexOf("--infile");
                if (i >= 0 && i + 1 < argv.Count) LastInfileXml = File.ReadAllText(argv[i + 1]);
            }

            var stdout = command.Name == "GetPage" ? SamplePageXml : "";
            return Task.FromResult(new CliResult(0, stdout, ""));
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
