using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Models;

namespace WordRegulator.Domain.Interfaces;

public interface IWordDocumentProcessor
{
    Task<DocumentProcessResult> ProcessDocumentAsync(string templatePath, string outputPath, IDictionary<string, string> variableReplacements, IEnumerable<DrillingWellItem>? wells = null, CancellationToken cancellationToken = default);
    Task<BatchProcessResult> ProcessTemplateDirectoryAsync(string templateDirectoryPath, string outputDirectoryPath, IDictionary<string, string> variableReplacements, IEnumerable<DrillingWellItem>? wells = null, CancellationToken cancellationToken = default);
}
