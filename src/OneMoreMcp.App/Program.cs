using System.Net;
using System.Security.Cryptography.X509Certificates;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Serilog;
using OneMoreMcp.App.Mcp;
using OneMoreMcp.App.Tray;

namespace OneMoreMcp.App;

internal static class Program
{
    // Session-local: one running instance per logged-in user.
    private const string SingleInstanceName = @"Local\OneMoreMcp.SingleInstance";

    private static Mutex? _singleInstance;

    [STAThread]
    private static void Main(string[] args)
    {
        // One-shot, scriptable commands (autostart management, icon generation).
        if (TryHandleOneShotCommand(args)) return;

        // Enforce a single running instance.
        _singleInstance = new Mutex(initiallyOwned: true, SingleInstanceName, out var createdNew);
        if (!createdNew)
        {
            MessageBox.Show(
                "OneMore MCP is already running — look for its icon in the system tray (near the clock).",
                "OneMore MCP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ApplicationConfiguration.Initialize();

        AppPaths.EnsureCreated();
        AppPaths.EnsureUserConfig();

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(
                Path.Combine(AppPaths.LogDir, "onemoremcp-.log"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7)
            .CreateLogger();

        try
        {
            RunApp(args);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "OneMore MCP terminated unexpectedly.");
            MessageBox.Show(ex.Message, "OneMore MCP — fatal error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            Log.CloseAndFlush();
            try { _singleInstance?.ReleaseMutex(); } catch (ApplicationException) { /* not owned */ }
            _singleInstance?.Dispose();
        }
    }

    /// <summary>
    /// Handles one-shot CLI commands and exits. Returns true if a command was handled (the caller
    /// should then return without starting the tray/server).
    /// </summary>
    private static bool TryHandleOneShotCommand(string[] args)
    {
        if (args.Length == 0) return false;
        switch (args[0].Trim().ToLowerInvariant())
        {
            case "--enable-autostart":
                StartupManager.SetEnabled(true);
                return true;
            case "--disable-autostart":
                StartupManager.SetEnabled(false);
                return true;
            case "--write-icon":
                var path = Path.GetFullPath(args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "App.ico"));
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                Tray.TrayIconFactory.WriteIco(path);
                return true;
            default:
                return false;
        }
    }

    private static void RunApp(string[] args)
    {
        var builder = WebApplication.CreateBuilder(new WebApplicationOptions
        {
            Args = args,
            ContentRootPath = AppContext.BaseDirectory,
        });

        // Layer the user-editable config (in %APPDATA%) on top of the bundled appsettings.json.
        builder.Configuration.AddJsonFile(AppPaths.UserConfigFile, optional: true, reloadOnChange: true);

        builder.Services.Configure<OneMoreMcpOptions>(
            builder.Configuration.GetSection(OneMoreMcpOptions.SectionName));

        var options = builder.Configuration
            .GetSection(OneMoreMcpOptions.SectionName)
            .Get<OneMoreMcpOptions>() ?? new OneMoreMcpOptions();

        var scheme = options.UseHttps ? "https" : "http";
        var url = $"{scheme}://localhost:{options.Port}";

        builder.Host.UseSerilog();

        X509Certificate2? certificate = null;
        if (options.UseHttps)
        {
            certificate = CertificateManager.GetOrCreate(AppPaths.CertificateFile, Log.Logger);
            if (options.TrustCertificate)
                CertificateManager.EnsureTrusted(certificate, Log.Logger);
        }

        builder.WebHost.ConfigureKestrel(kestrel =>
        {
            void Configure(ListenOptions listen)
            {
                if (certificate is not null) listen.UseHttps(certificate);
            }
            // Loopback only (IPv4 + IPv6) — never exposed beyond this machine.
            kestrel.Listen(IPAddress.Loopback, options.Port, Configure);
            kestrel.Listen(IPAddress.IPv6Loopback, options.Port, Configure);
        });

        builder.Services.AddSingleton<IOneMoreRunner, OneMoreCliRunner>();
        builder.Services
            .AddMcpServer()
            .WithHttpTransport(o => o.Stateless = true)
            .WithTools<OneMoreTools>();

        var app = builder.Build();
        app.MapMcp();

        try
        {
            app.StartAsync().GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Failed to start the MCP server on {Url}.", url);
            MessageBox.Show(
                $"Could not start the MCP server on {url}.\n\n{ex.Message}\n\n" +
                "The port may already be in use; change it in the configuration file.",
                "OneMore MCP", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        Log.Information("OneMore MCP listening on {Url} (endpoint: {Url}/).", url, url);

        using var tray = new TrayApplicationContext(
            url,
            AppPaths.UserConfigFile,
            AppPaths.LogDir,
            app.Services.GetRequiredService<IOptionsMonitor<OneMoreMcpOptions>>(),
            app.Services.GetRequiredService<IOneMoreRunner>(),
            certificate);

        Application.Run(tray);

        app.StopAsync().GetAwaiter().GetResult();
    }
}
