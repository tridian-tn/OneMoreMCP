using System.Text;

namespace OneMoreMcp.App;

/// <summary>Well-known on-disk locations for the app's user config and logs.</summary>
internal static class AppPaths
{
    public static string Root { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OneMoreMcp");

    public static string LogDir { get; } = Path.Combine(Root, "logs");

    public static string UserConfigFile { get; } = Path.Combine(Root, "config.json");

    public static string CertificateFile { get; } = Path.Combine(Root, "onemoremcp-localhost.pfx");

    public static void EnsureCreated()
    {
        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(LogDir);
    }

    /// <summary>Writes a starter config the first time the app runs, so users have something to edit.</summary>
    public static void EnsureUserConfig()
    {
        if (File.Exists(UserConfigFile)) return;
        File.WriteAllText(UserConfigFile, DefaultConfigJson, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private const string DefaultConfigJson =
        """
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
        """;
}
