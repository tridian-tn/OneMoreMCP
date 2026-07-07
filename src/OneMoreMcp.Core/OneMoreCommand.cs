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
    /// Adds a <c>--flag yes</c> when <paramref name="present"/> is true; otherwise a no-op. The OneMore
    /// CLI models booleans as <c>&lt;yes/no&gt;</c> options (default <c>no</c>), not bare switches, so a
    /// value must be supplied — a bare <c>--flag</c> is treated as absent.
    /// </summary>
    public OneMoreCommand Switch(string flag, bool present)
    {
        if (present) { _args.Add("--" + flag); _args.Add("yes"); }
        return this;
    }

    /// <summary>Adds <c>--flag yes|no</c> unconditionally — for booleans the CLI marks as required.</summary>
    public OneMoreCommand Bool(string flag, bool value)
    {
        _args.Add("--" + flag);
        _args.Add(value ? "yes" : "no");
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

    public static OneMoreCommand Search(string query) =>
        new OneMoreCommand("Search").Option("query", query);

    public static OneMoreCommand SearchHashtags(string query, bool allTags = false) =>
        new OneMoreCommand("SearchHashtags").Option("query", query).Switch("allTags", allTags);

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

    // InsertToc requires --notebook, --section, --page and a --refresh yes/no value.
    public static OneMoreCommand InsertToc(string? notebook, string? section, string? page, bool refresh) =>
        new OneMoreCommand("InsertToc")
            .Option("notebook", notebook).Option("section", section).Option("page", page).Bool("refresh", refresh);

    public static OneMoreCommand Export(string outpath, string format, string? pageId = null, bool backup = false) =>
        new OneMoreCommand("Export").Option("outpath", outpath).Option("format", format).Option("pageId", pageId).Switch("backup", backup);

    public static OneMoreCommand Goto(string pageId, string? objectId = null) =>
        new OneMoreCommand("Goto").Option("pageId", pageId).Option("objectId", objectId);

    /// <summary>A parameterless housekeeping command (ApplyStyles, RemoveEmpty, Trim, …).</summary>
    public static OneMoreCommand Simple(string name) => new(name);
}
