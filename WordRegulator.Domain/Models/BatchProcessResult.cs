namespace WordRegulator.Domain.Models;

public class BatchProcessResult
{
    public bool Success { get; set; }
    public int ProcessedFileCount { get; set; }
    public int TotalReplacementsCount { get; set; }
    public List<string> OutputFiles { get; set; } = new();
    public List<string> Errors { get; set; } = new();
}
