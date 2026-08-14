using WordRegulator.Domain.Models;

namespace WordRegulator.Domain.Interfaces;

public interface IWordDocumentProcessor
{
    Task<DocumentProcessResult> ProcessDocumentAsync(string templatePath, string outputPath, IDictionary<string, string> variableReplacements, CancellationToken cancellationToken = default);
    Task<BatchProcessResult> ProcessTemplateDirectoryAsync(string templateDirectoryPath, string outputDirectoryPath, IDictionary<string, string> variableReplacements, CancellationToken cancellationToken = default);
}
