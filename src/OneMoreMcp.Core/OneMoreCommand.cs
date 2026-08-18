namespace OneMoreMcp.Core;

/// <summary>
/// Builds a <c>OneMoreCli.exe</c> invocation as a command name plus an ordered argument list, e.g.
/// <c>GetPage --section "My Section" --current</c>. Kept pure and free of any process plumbing so the
/// exact argv a tool produces can be asserted in isolation. The runner feeds <see cref="Build"/> into
/// <see cref="System.Diagnostics.ProcessStartInfo.ArgumentList"/>, which quotes each element itself —
/// so values here are stored verbatim (spaces and all), never pre-quoted.
/// </summary>
public sealed class OneMoreCommand
{
    private readonly List<string> _args = new();

    public OneMoreCommand(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("A command name is required.", nameof(name));
        Name = name;
    }

    /// <summary>The OneMore command verb, e.g. <c>GetPage</c>.</summary>
    public string Name { get; }

    /// <summary>Adds <c>--flag value</c> when <paramref name="value"/> is non-blank; otherwise a no-op.</summary>
    public OneMoreCommand Option(string flag, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            _args.Add("--" + flag);
            _args.Add(value);
        }
        return this;
    }

    /// <summary>
    /// Adds a bare <c>--flag</c> switch when <paramref name="present"/> is true; otherwise a no-op.
    /// The OneMore CLI (7.3.0+) models booleans as PowerShell-style switches: presence enables them,
    /// absence uses the default.
    /// </summary>
    public OneMoreCommand Switch(string flag, bool present)
    {
        if (present) _args.Add("--" + flag);
        return this;
    }

    /// <summary>
    /// Adds the CLI's global <c>--output &lt;file&gt;</c> option, directing command output to a file
    /// instead of stdout. Avoids console/shell encoding corruption of non-ASCII page content.
    /// Requires a OneMore build that supports <c>--output</c> (added in the CLI after 7.2.0).
    /// </summary>
    public OneMoreCommand Output(string? path) => Option("output", path);

    /// <summary>The full argument vector: the command name followed by its options, in order.</summary>
    public IReadOnlyList<string> Build()
    {
        var argv = new List<string>(_args.Count + 1) { Name };
        argv.AddRange(_args);
        return argv;
    }

    /// <summary>A shell-ish preview of the invocation, for logs and error messages (not for execution).</summary>
    public override string ToString() =>
        string.Join(' ', Build().Select(a => a.Contains(' ') ? $"\"{a}\"" : a));

    // --- Command factories (documented OneMore CLI surface) ---

    public static OneMoreCommand GetHierarchy(string? notebook = null, string? section = null, bool books = false) =>
        new OneMoreCommand("GetHierarchy").Option("notebook", notebook).Option("section", section).Switch("books", books);

    // GetPage requires --notebook together with --section/--page (the CLI docs omit --notebook, but the
    // real CLI rejects --page without it); --current needs none of them.
    public static OneMoreCommand GetPage(string? notebook, string? section, string? page, bool current) =>
        new OneMoreCommand("GetPage")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Switch("current", current);

    // Search requires --notebook (7.3.0); --section/--page narrow the scope.
    public static OneMoreCommand Search(string query, string? notebook, string? section = null, string? page = null) =>
        new OneMoreCommand("Search")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Option("query", query);

    public static OneMoreCommand SearchHashtags(string query, bool allTags = false,
        string? notebook = null, string? section = null, string? page = null) =>
        new OneMoreCommand("SearchHashtags")
            .Option("notebook", notebook).Option("section", section).Option("page", page)
            .Option("query", query).Switch("allTags", allTags);

    public static OneMoreCommand SearchTitles(string query, string? notebook = null) =>
        new OneMoreCommand("SearchTitles").Option("query", query).Option("notebook", notebook);

    public static OneMoreCommand Sync(string notebook) =>
        new OneMoreCommand("Sync").Option("notebook", notebook);

    public static OneMoreCommand PutPage(string? notebook, string? section, string? page, string infile, bool force = true) =>
        new OneMoreCommand("PutPage")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Option("infile", infile).Switch("force", force);

    // AddHashtag/RemoveHashtag require --notebook (section/page optional); a missing required option
    // makes the CLI drop into an interactive prompt, so always supply the notebook.
    public static OneMoreCommand AddHashtag(string tags, string? notebook, string? section = null, string? page = null) =>
        new OneMoreCommand("AddHashtag")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Option("tags", tags);

    public static OneMoreCommand RemoveHashtag(string tags, string? notebook, string? section = null, string? page = null) =>
        new OneMoreCommand("RemoveHashtag")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Option("tags", tags);

    public static OneMoreCommand Export(string outpath, string format, string? pageId = null, bool backup = false) =>
        new OneMoreCommand("Export").Option("outpath", outpath).Option("format", format).Option("pageId", pageId).Switch("backup", backup);

    public static OneMoreCommand Goto(string pageId, string? objectId = null) =>
        new OneMoreCommand("Goto").Option("pageId", pageId).Option("objectId", objectId);

    public static OneMoreCommand Archive(string notebook, string? section, string outfile) =>
        new OneMoreCommand("Archive").Option("notebook", notebook).Option("section", section).Option("outfile", outfile);

    public static OneMoreCommand Diagnostics(bool includeWindows = false) =>
        new OneMoreCommand("Diagnostics").Switch("windows", includeWindows);

    /// <summary>
    /// A page-maintenance command scoped to a notebook (section/page optional). These commands
    /// (ApplyStyles, RemoveEmpty, Trim, Embed, …) require <c>--notebook</c> in 7.3.0.
    /// </summary>
    public static OneMoreCommand Cleanup(string name, string? notebook, string? section = null, string? page = null) =>
        new OneMoreCommand(name).Option("notebook", notebook).Option("section", section).Option("page", page);
}
