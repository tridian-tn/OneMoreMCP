using Microsoft.Win32;

namespace OneMoreMcp.App;

/// <summary>
/// Manages whether the app launches when the current user logs in, via the per-user
/// <c>HKCU\Software\Microsoft\Windows\CurrentVersion\Run</c> key. Per-user (HKCU) means
/// no administrator rights are required and it only affects this user's logon.
/// </summary>
internal static class StartupManager
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "OneMoreMcp";

    /// <summary>Path to the running executable, used as the command to launch at logon.</summary>
    public static string ExecutablePath =>
        Environment.ProcessPath ?? Path.Combine(AppContext.BaseDirectory, "OneMoreMcp.exe");

    public static bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        var existing = key?.GetValue(ValueName) as string;
        if (string.IsNullOrWhiteSpace(existing)) return false;

        // Treat it as enabled only if it still points at this executable.
        return string.Equals(
            Unquote(existing), ExecutablePath, StringComparison.OrdinalIgnoreCase);
    }

    public static void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
                        ?? Registry.CurrentUser.CreateSubKey(RunKeyPath);
        if (key is null)
            throw new InvalidOperationException("Could not open the Windows startup registry key.");

        if (enabled)
            key.SetValue(ValueName, $"\"{ExecutablePath}\"", RegistryValueKind.String);
        else
            key.DeleteValue(ValueName, throwOnMissingValue: false);
    }

    private static string Unquote(string value)
    {
        value = value.Trim();
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            value = value[1..^1];
        return value;
    }
}
