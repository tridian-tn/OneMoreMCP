using System.Reflection;

namespace OneMoreMcp.App;

/// <summary>
/// The running build's version, read from the assembly attributes MinVer and the .NET SDK stamp at
/// build time. <see cref="Informational"/> is the full string (e.g. "0.0.3-alpha.0.5+74cab42…");
/// on a tagged release commit it is a clean SemVer such as "0.0.2".
/// </summary>
internal static class VersionInfo
{
    /// <summary>Full informational version: the SemVer plus the commit hash as build metadata.</summary>
    public static string Informational { get; }

    /// <summary>The SemVer portion, without the "+commit" build metadata.</summary>
    public static string SemVer { get; }

    /// <summary>The full commit hash, or null when the build carries none (e.g. no .git on build).</summary>
    public static string? CommitHash { get; }

    static VersionInfo()
    {
        Informational = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? "0.0.0-alpha.0"; // matches MinVer's fallback when no version tag is reachable

        // The SDK appends the commit hash as SemVer build metadata: "<semver>+<hash>".
        var plus = Informational.IndexOf('+');
        if (plus < 0)
        {
            SemVer = Informational;
            CommitHash = null;
        }
        else
        {
            SemVer = Informational[..plus];
            CommitHash = Informational[(plus + 1)..];
        }
    }

    /// <summary>The commit hash shortened to 7 characters, or null when absent.</summary>
    public static string? ShortCommit => CommitHash is { Length: >= 7 } h ? h[..7] : CommitHash;

    /// <summary>Friendly one-line version, e.g. "0.0.3-alpha.0.5 (74cab42)" or "0.0.2".</summary>
    public static string Display => ShortCommit is null ? SemVer : $"{SemVer} ({ShortCommit})";
}
