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
    public string? TargetKey { get; set; }
    public string? TargetMachine { get; set; }
    public int RepeatCount { get; set; } = 1;
    public int IntervalMinutes { get; set; }
    public bool Enabled { get; set; } = true;
    public string? Title { get; set; }
    public string? Message { get; set; }
}

public class RemoteCommandManifest
{
    public int PollIntervalMinutes { get; set; } = 15;
    public List<RemoteCommandItem> Commands { get; set; } = [];
}

public class CommandExecutionRecord
{
    public int ExecutedCount { get; set; }
    public DateTime? LastExecutedUtc { get; set; }
}
