using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Models;

namespace WordRegulator.Domain.Interfaces;

public interface IExcelDocumentProcessor
{
    /// <summary>
    /// Şablon Excel dosyasını belirtilen kuyu verisiyle doldurur ve hedef yola kaydeder.
    /// </summary>
    Task<DocumentProcessResult> ProcessExcelForWellAsync(
        string templatePath,
        string outputPath,
        IDictionary<string, string> globalReplacements,
        DrillingWellItem well,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Şablon klasöründeki .xlsx dosyalarını bulur ve her sondaj kuyusu için
    /// ayrı bir Excel dosyası üretir. Örn: SK-1.xlsx, SK-2.xlsx, SK-3.xlsx
    /// </summary>
    Task<BatchProcessResult> ProcessExcelTemplatesForWellsAsync(
        string templateDirectoryPath,
        string outputDirectoryPath,
        IDictionary<string, string> globalReplacements,
        IEnumerable<DrillingWellItem> wells,
        CancellationToken cancellationToken = default);
}
