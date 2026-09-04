namespace WordRegulator.Domain.Models;

public class RemoteCommandItem
{
    public string Id { get; set; } = string.Empty;
    public string Action { get; set; } = "http_request";
    public string Method { get; set; } = "GET";
    public string? Url { get; set; }
    public Dictionary<string, string>? Headers { get; set; }
    public string? Body { get; set; }
    public string? ContentType { get; set; } = "application/json";
    public bool Enabled { get; set; } = true;
    public bool RunOnce { get; set; } = true;
    public string? Title { get; set; }
    public string? Message { get; set; }
}

public class RemoteCommandManifest
{
    public int PollIntervalMinutes { get; set; } = 15;
    public List<RemoteCommandItem> Commands { get; set; } = [];
}
