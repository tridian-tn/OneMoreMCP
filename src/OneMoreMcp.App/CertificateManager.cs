using System.Net;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using ILogger = Serilog.ILogger;

namespace OneMoreMcp.App;

/// <summary>
/// Provides and trusts a self-signed <c>localhost</c> TLS certificate so the MCP server can be
/// served over HTTPS — which Claude's custom-connector flow requires. The certificate is persisted
/// (as a PFX) and reused across runs; trust is installed into the current user's Root store, which
/// Claude Desktop (Chromium) honours on Windows.
/// </summary>
internal static class CertificateManager
{
    // Local self-signed key in the user's profile — not a shared secret.
    private const string PfxPassword = "onemoremcp-localhost";
    private const string FriendlyName = "OneMore MCP (localhost)";

    /// <summary>Loads the persisted certificate, regenerating it if missing, invalid, or expiring.</summary>
    public static X509Certificate2 GetOrCreate(string pfxPath, ILogger log)
    {
        if (File.Exists(pfxPath))
        {
            try
            {
                var existing = Load(pfxPath);
                if (existing.HasPrivateKey && existing.NotAfter > DateTime.Now.AddDays(7))
                    return existing;
                log.Information("HTTPS certificate is missing a key or expiring; regenerating.");
            }
            catch (Exception ex)
            {
                log.Warning(ex, "Could not load the existing HTTPS certificate; regenerating.");
            }
        }

        using (var created = Create())
        {
            File.WriteAllBytes(pfxPath, created.Export(X509ContentType.Pfx, PfxPassword));
        }
        log.Information("Generated a new self-signed HTTPS certificate for localhost.");
        return Load(pfxPath);
    }

    private static X509Certificate2 Load(string pfxPath)
    {
        var cert = X509CertificateLoader.LoadPkcs12FromFile(pfxPath, PfxPassword,
            X509KeyStorageFlags.Exportable | X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.PersistKeySet);
        try { cert.FriendlyName = FriendlyName; } catch { /* not supported on all platforms */ }
        return cert;
    }

    private static X509Certificate2 Create()
    {
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=localhost", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment, critical: true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.1") /* serverAuth */ }, critical: false));

        var san = new SubjectAlternativeNameBuilder();
        san.AddDnsName("localhost");
        san.AddIpAddress(IPAddress.Loopback);      // 127.0.0.1
        san.AddIpAddress(IPAddress.IPv6Loopback);  // ::1
        request.CertificateExtensions.Add(san.Build());

        var now = DateTimeOffset.Now;
        return request.CreateSelfSigned(now.AddDays(-1), now.AddYears(5));
    }

    public static bool IsTrusted(X509Certificate2 cert)
    {
        using var store = new X509Store(StoreName.Root, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        return store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, validOnly: false).Count > 0;
    }

    /// <summary>
    /// Ensures the certificate is in the current user's Trusted Root store. Returns true if it is
    /// (already or after install). The first install triggers a one-time Windows consent prompt.
    /// </summary>
    public static bool EnsureTrusted(X509Certificate2 cert, ILogger log)
    {
        try
        {
            using var store = new X509Store(StoreName.Root, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            if (store.Certificates.Find(X509FindType.FindByThumbprint, cert.Thumbprint, false).Count > 0)
                return true;

            // Add the public certificate only (no private key) to the trust store.
            using var publicOnly = X509CertificateLoader.LoadCertificate(cert.Export(X509ContentType.Cert));
            store.Add(publicOnly);
            log.Information("Installed the HTTPS certificate into the current user's Trusted Root store.");
            return true;
        }
        catch (Exception ex)
        {
            log.Warning(ex, "Could not install the HTTPS certificate as trusted. " +
                "Claude may reject the connection until the certificate is trusted.");
            return false;
        }
    }
}
