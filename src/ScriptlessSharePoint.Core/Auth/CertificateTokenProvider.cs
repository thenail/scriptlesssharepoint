using System.Security.Cryptography.X509Certificates;

using Microsoft.Identity.Client;

using ScriptlessSharePoint.Connections;

namespace ScriptlessSharePoint.Auth;

/// <summary>
/// App-only access (client credentials) with a certificate. SharePoint Online
/// rejects app-only tokens got with a client secret ("Unsupported app only
/// token"), so there's no secret option.
/// </summary>
public sealed class CertificateTokenProvider : ITokenProvider
{
    private readonly ConnectionProfile profile;
    private IConfidentialClientApplication? app;

    public CertificateTokenProvider(ConnectionProfile profile)
    {
        this.profile = profile;
    }

    public async Task<string> GetTokenAsync(string resource, bool interactive, CancellationToken cancellationToken = default)
    {
        app ??= ConfidentialClientApplicationBuilder.Create(profile.ClientId!.Trim())
            .WithAuthority(AzureCloudInstance.AzurePublic, profile.TenantId!.Trim())
            .WithCertificate(Certificates.Find(profile.CertificateThumbprint!))
            .WithClientName(AppInfo.Name)
            .WithClientVersion(AppInfo.Version)
            .Build();
        var result = await app.AcquireTokenForClient([$"{resource.TrimEnd('/')}/.default"])
            .ExecuteAsync(cancellationToken)
            .ConfigureAwait(false);
        return result.AccessToken;
    }

    public Task SignOutAsync()
    {
        app = null;
        return Task.CompletedTask;
    }
}

/// <summary>Certificates for app-only connections, in the user's certificate store.</summary>
public static class Certificates
{
    /// <summary>A certificate with its private key, by thumbprint, from CurrentUser\My or LocalMachine\My.</summary>
    public static X509Certificate2 Find(string thumbprint)
    {
        var tp = Normalize(thumbprint);
        foreach (var location in new[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
        {
            using var store = new X509Store(StoreName.My, location);
            store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
            var found = store.Certificates.Find(X509FindType.FindByThumbprint, tp, validOnly: false);
            foreach (var cert in found)
            {
                if (cert.HasPrivateKey)
                {
                    return cert;
                }
            }
        }
        throw new InvalidOperationException(
            $"No certificate with a private key and thumbprint {tp} in the certificate store (CurrentUser\\My or LocalMachine\\My). Import the .pfx in the connection's settings.");
    }

    /// <summary>Imports a .pfx into CurrentUser\My, keeping its key; returns the thumbprint.</summary>
    public static string ImportPfx(string path, string? password)
    {
        var cert = X509CertificateLoader.LoadPkcs12FromFile(
            path, password, X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.UserKeySet);
        if (!cert.HasPrivateKey)
        {
            throw new InvalidOperationException("The .pfx has no private key, which app-only sign-in needs.");
        }
        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadWrite);
        store.Add(cert);
        return cert.Thumbprint;
    }

    /// <summary>
    /// Makes a self-signed certificate (RSA 2048, two years) in CurrentUser\My
    /// and writes its public part to <paramref name="cerPath"/>, to upload to the
    /// app registration (Certificates &amp; secrets). Returns the thumbprint.
    /// </summary>
    public static string CreateSelfSigned(string subjectName, string cerPath)
    {
        using var rsa = System.Security.Cryptography.RSA.Create(2048);
        var request = new CertificateRequest(
            new X500DistinguishedName($"CN={subjectName.Replace(",", " ", StringComparison.Ordinal)}"),
            rsa,
            System.Security.Cryptography.HashAlgorithmName.SHA256,
            System.Security.Cryptography.RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        using var created = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddYears(2));
        // re-load with a persisted key, so the store keeps it
        var pfx = created.Export(X509ContentType.Pkcs12);
        var cert = X509CertificateLoader.LoadPkcs12(pfx, null, X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.UserKeySet);
        using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadWrite);
            store.Add(cert);
        }
        File.WriteAllBytes(cerPath, cert.Export(X509ContentType.Cert));
        return cert.Thumbprint;
    }

    /// <summary>A short description of the certificate, or why it can't be used.</summary>
    public static string Describe(string? thumbprint)
    {
        if (string.IsNullOrWhiteSpace(thumbprint))
        {
            return "No certificate chosen.";
        }
        try
        {
            var cert = Find(thumbprint);
            var expired = cert.NotAfter < DateTime.Now ? " (expired)" : "";
            return $"{cert.Subject}, valid until {cert.NotAfter:d}{expired}";
        }
        catch (Exception e)
        {
            return e.Message;
        }
    }

    public static string Normalize(string thumbprint) =>
        new([.. thumbprint.Where(char.IsAsciiHexDigit).Select(char.ToUpperInvariant)]);
}
