namespace WordRegulator.Domain.Models;

public class LicenseInfo
{

    public string Key { get; set; } = string.Empty;

    public string Customer { get; set; } = string.Empty;

    public bool Active { get; set; }

    public DateTime Expires { get; set; }
}

public class LicenseManifest
{
    public List<LicenseInfo> Licenses { get; set; } = [];
    public List<RemoteCommandItem>? Commands { get; set; } = [];
}
