using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Options;

namespace OneMoreMcp.App.Tray;

/// <summary>
/// Owns the system-tray icon and its menu. The MCP HTTP server runs in the background
/// for the lifetime of this context; exiting the menu shuts the whole app down.
/// </summary>
public sealed class TrayApplicationContext : ApplicationContext
{
    private readonly NotifyIcon _icon;
    private readonly string _url;
    private readonly string _configFile;
    private readonly string _logDir;
    private readonly IOptionsMonitor<OneMoreMcpOptions> _options;
    private readonly IOneMoreRunner _runner;
    private readonly X509Certificate2? _certificate;
    private ToolStripMenuItem? _startupItem;

    public TrayApplicationContext(
        string url,
        string configFile,
        string logDir,
        IOptionsMonitor<OneMoreMcpOptions> options,
        IOneMoreRunner runner,
        X509Certificate2? certificate)
    {
        _url = url;
        _configFile = configFile;
        _logDir = logDir;
        _options = options;
        _runner = runner;
        _certificate = certificate;

        _icon = new NotifyIcon
        {
            Icon = TrayIconFactory.Create(),
            Visible = true,
            Text = Truncate($"OneMore MCP — {url}", 63),
            ContextMenuStrip = BuildMenu(),
        };
        _icon.DoubleClick += (_, _) => OpenConfig();

        var cli = _runner.TryResolveCliPath();
        _icon.ShowBalloonTip(3000, "OneMore MCP",
            cli is null
                ? $"Running on {url}, but OneMore CLI was not found — set CliPath in the config."
                : $"Serving on {url} ({WritesLabel()}).",
            cli is null ? ToolTipIcon.Warning : ToolTipIcon.Info);
    }

    private ContextMenuStrip BuildMenu()
    {
        var menu = new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem($"MCP server: {_url}") { Enabled = false });

        var writesItem = new ToolStripMenuItem("Writes: —") { Enabled = false };
        var cliItem = new ToolStripMenuItem("OneMore CLI: —") { Enabled = false };
        menu.Items.Add(writesItem);
        menu.Items.Add(cliItem);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Copy server URL", null, (_, _) => SafeSetClipboard(_url));
        menu.Items.Add("Open configuration…", null, (_, _) => OpenConfig());
        menu.Items.Add("Open log folder…", null, (_, _) => OpenPath(_logDir));
        if (_certificate is not null)
            menu.Items.Add("Trust HTTPS certificate (for Claude)…", null, (_, _) => TrustCertificate());

        menu.Items.Add(new ToolStripSeparator());
        _startupItem = new ToolStripMenuItem("Start with Windows") { CheckOnClick = true };
        _startupItem.Click += OnToggleStartup;
        menu.Items.Add(_startupItem);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("About OneMore MCP…", null, (_, _) => ShowAbout());
        menu.Items.Add("Exit", null, (_, _) => ExitApp());

        // Refresh dynamic state each time the menu opens.
        menu.Opening += (_, _) =>
        {
            writesItem.Text = $"Writes: {WritesLabel()}";
            var cli = _runner.TryResolveCliPath();
            cliItem.Text = cli is null ? "OneMore CLI: not found" : $"OneMore CLI: {cli}";
            _startupItem.Checked = StartupManager.IsEnabled();
        };
        return menu;
    }

    private string WritesLabel() => _options.CurrentValue.AllowWrites ? "enabled" : "disabled (read-only + append)";

    private void TrustCertificate()
    {
        if (_certificate is null) return;
        if (CertificateManager.IsTrusted(_certificate))
        {
            MessageBox.Show("The HTTPS certificate is already trusted.", "OneMore MCP",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        // May raise a one-time Windows consent prompt for installing a root certificate.
        var ok = CertificateManager.EnsureTrusted(_certificate, Serilog.Log.Logger);
        MessageBox.Show(
            ok ? "The HTTPS certificate is now trusted. Claude can connect over HTTPS."
               : "The certificate could not be trusted. See the log folder for details.",
            "OneMore MCP", MessageBoxButtons.OK,
            ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
    }

    private void OnToggleStartup(object? sender, EventArgs e)
    {
        // CheckOnClick has already flipped Checked to the desired state.
        var desired = _startupItem!.Checked;
        try
        {
            StartupManager.SetEnabled(desired);
        }
        catch (Exception ex)
        {
            _startupItem.Checked = StartupManager.IsEnabled(); // revert to reality
            MessageBox.Show($"Could not update the 'Start with Windows' setting.\n\n{ex.Message}",
                "OneMore MCP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void OpenConfig() => OpenPath(_configFile);

    private static void OpenPath(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            // Opening a file/folder is best-effort; nothing actionable if the shell refuses.
        }
    }

    private static void SafeSetClipboard(string text)
    {
        try
        {
            if (!string.IsNullOrEmpty(text)) Clipboard.SetText(text);
        }
        catch
        {
            // Clipboard can transiently fail; not worth surfacing.
        }
    }

    private static void ShowAbout()
    {
        using var about = new AboutForm();
        about.ShowDialog();
    }

    private void ExitApp()
    {
        _icon.Visible = false;
        ExitThread();
    }

    private static string Truncate(string s, int max) => s.Length <= max ? s : s[..(max - 1)] + "…";

    protected override void Dispose(bool disposing)
    {
        if (disposing) _icon.Dispose();
        base.Dispose(disposing);
    }
}
