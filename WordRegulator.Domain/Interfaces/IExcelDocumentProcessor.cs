using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Models;

namespace WordRegulator.Domain.Interfaces;

public interface IExcelDocumentProcessor
{
        Task<DocumentProcessResult> ProcessExcelForWellAsync(
        string templatePath,
        string outputPath,
        IDictionary<string, string> globalReplacements,
        DrillingWellItem well,
        CancellationToken cancellationToken = default);

        Task<BatchProcessResult> ProcessExcelTemplatesForWellsAsync(
        string templateDirectoryPath,
        string outputDirectoryPath,
        IDictionary<string, string> globalReplacements,
        IEnumerable<DrillingWellItem> wells,
        CancellationToken cancellationToken = default);
}
