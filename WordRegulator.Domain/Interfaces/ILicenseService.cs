namespace WordRegulator.Domain.Interfaces;

public interface ILicenseService
{

    Task<LicenseCheckResult> CheckLicenseAsync(string licenseKey);

    string? GetStoredLicenseKey();

    void StoreLicenseKey(string licenseKey);
}

public class LicenseCheckResult
{
    public bool IsValid { get; init; }
    public string Message { get; init; } = string.Empty;
    public bool IsOfflineFallback { get; init; }

    public static LicenseCheckResult Valid(bool isOffline = false) => new()
    {
        IsValid = true,
        Message = isOffline ? "Lisans çevrimdışı doğrulandı." : "Lisans geçerli.",
        IsOfflineFallback = isOffline
    };

    public static LicenseCheckResult Invalid(string reason) => new()
    {
        IsValid = false,
        Message = reason
    };
}
