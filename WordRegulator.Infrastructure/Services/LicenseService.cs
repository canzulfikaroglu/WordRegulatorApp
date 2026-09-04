using Microsoft.Win32;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Domain.Models;

namespace WordRegulator.Infrastructure.Services;

public class LicenseService : ILicenseService
{
    private const string SettingsFileName = "licenseSettings.json";
    private const string RegistryKeyPath = @"SOFTWARE\WordRegulator";
    private const string RegistryValueName = "LicenseKey";
    private const int GracePeriodDays = 7;
    private static readonly string CacheFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WordRegulator",
        "lc.dat");

    private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(10) };

    public async Task<LicenseCheckResult> CheckLicenseAsync(string licenseKey)
    {
        if (string.IsNullOrWhiteSpace(licenseKey))
            return LicenseCheckResult.Invalid("Lisans anahtarı boş olamaz.");

        var gistUrl = GetGistUrl();

        if (!string.IsNullOrWhiteSpace(gistUrl))
        {
            try
            {
                var manifest = await FetchManifestFromGistAsync(gistUrl);
                var result = ValidateKey(licenseKey, manifest);

                if (result.IsValid)
                    SaveCache(licenseKey); 

                return result;
            }
            catch (Exception)
            {

            }
        }

        return CheckOfflineCache(licenseKey);
    }

    public string? GetStoredLicenseKey()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
        return key?.GetValue(RegistryValueName) as string;
    }

    public void StoreLicenseKey(string licenseKey)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath, writable: true);
        key.SetValue(RegistryValueName, licenseKey);
    }

    private static string? GetGistUrl()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, SettingsFileName);
            if (!File.Exists(path)) return null;

            var json = File.ReadAllText(path);
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("GistUrl", out var prop))
            {
                return prop.GetString();
            }
        }
        catch
        {

        }

        return null;
    }

    private static async Task<LicenseManifest> FetchManifestFromGistAsync(string gistUrl)
    {
        var json = await _httpClient.GetStringAsync(gistUrl);
        return JsonSerializer.Deserialize<LicenseManifest>(json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
               ?? new LicenseManifest();
    }

    private static LicenseCheckResult ValidateKey(string key, LicenseManifest manifest)
    {
        var license = manifest.Licenses
            .FirstOrDefault(l => string.Equals(l.Key, key, StringComparison.OrdinalIgnoreCase));

        if (license is null)
            return LicenseCheckResult.Invalid("Lisans anahtarı tanınmıyor.");

        if (!license.Active)
            return LicenseCheckResult.Invalid(
                "Lisansınız devre dışı bırakıldı. Lütfen satıcı ile iletişime geçin.");

        if (license.Expires < DateTime.UtcNow)
            return LicenseCheckResult.Invalid(
                $"Lisansınızın süresi doldu ({license.Expires:dd.MM.yyyy}). Yenileme için satıcı ile iletişime geçin.");

        return LicenseCheckResult.Valid();
    }

    private record CacheEntry(string LicenseKey, DateTime LastVerified);

    private void SaveCache(string licenseKey)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(CacheFilePath)!);
            var entry = new CacheEntry(licenseKey, DateTime.UtcNow);
            var json = JsonSerializer.Serialize(entry);
            var encrypted = EncryptAes(json);
            File.WriteAllBytes(CacheFilePath, encrypted);
        }
        catch { }
    }

    private LicenseCheckResult CheckOfflineCache(string licenseKey)
    {
        try
        {
            if (!File.Exists(CacheFilePath))
                return LicenseCheckResult.Invalid("İnternet bağlantısı yok ve yerel lisans kaydı bulunamadı.");

            var encrypted = File.ReadAllBytes(CacheFilePath);
            var json = DecryptAes(encrypted);
            var entry = JsonSerializer.Deserialize<CacheEntry>(json);

            if (entry is null)
                return LicenseCheckResult.Invalid("Lisans cache dosyası okunamadı.");

            if (!string.Equals(entry.LicenseKey, licenseKey, StringComparison.OrdinalIgnoreCase))
                return LicenseCheckResult.Invalid("İnternet bağlantısı yok ve lisans anahtarı eşleşmiyor.");

            var daysSinceLastCheck = (DateTime.UtcNow - entry.LastVerified).TotalDays;
            if (daysSinceLastCheck > GracePeriodDays)
                return LicenseCheckResult.Invalid(
                    $"İnternet bağlantısı {GracePeriodDays} günden uzun süredir yok. Lisans doğrulanamıyor.");

            return LicenseCheckResult.Valid(isOffline: true);
        }
        catch
        {
            return LicenseCheckResult.Invalid("Lisans doğrulanamadı. İnternet bağlantınızı kontrol edin.");
        }
    }

    private static byte[] GetMachineKey()
    {

        var machineId = Environment.MachineName + Environment.UserName + "WR-SALT-2025";
        return SHA256.HashData(Encoding.UTF8.GetBytes(machineId));
    }

    private static byte[] EncryptAes(string plainText)
    {
        using var aes = Aes.Create();
        aes.Key = GetMachineKey();
        aes.GenerateIV();
        using var encryptor = aes.CreateEncryptor();
        var data = Encoding.UTF8.GetBytes(plainText);
        var encrypted = encryptor.TransformFinalBlock(data, 0, data.Length);
        return [.. aes.IV, .. encrypted];
    }

    private static string DecryptAes(byte[] cipherData)
    {
        using var aes = Aes.Create();
        aes.Key = GetMachineKey();
        var iv = cipherData[..16];
        var cipher = cipherData[16..];
        aes.IV = iv;
        using var decryptor = aes.CreateDecryptor();
        var decrypted = decryptor.TransformFinalBlock(cipher, 0, cipher.Length);
        return Encoding.UTF8.GetString(decrypted);
    }
}
