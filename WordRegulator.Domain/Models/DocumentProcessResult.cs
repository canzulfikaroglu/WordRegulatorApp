namespace WordRegulator.Domain.Models;

public class DocumentProcessResult
{
    public bool Success { get; set; }
    public string OutputFilePath { get; set; } = string.Empty;
    public int ReplacedCount { get; set; }
    public string? ErrorMessage { get; set; }
    public List<string> ReplacedVariables { get; set; } = new();
}
