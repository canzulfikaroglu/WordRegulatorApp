using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Domain.Models;

namespace WordRegulator.Infrastructure.Services;

public class OpenXmlWordProcessor : IWordDocumentProcessor
{
    private static readonly Regex PlaceholderRegex = new Regex(@"\{\{\s*([a-zA-Z0-9_İıŞşĞğÜüÖöÇç\-]+)\s*\}\}", RegexOptions.Compiled);

    public async Task<DocumentProcessResult> ProcessDocumentAsync(
        string templatePath,
        string outputPath,
        IDictionary<string, string> variableReplacements,
        IEnumerable<DrillingWellItem>? wells = null,
        CancellationToken cancellationToken = default)
    {
        var result = new DocumentProcessResult();

        try
        {
            if (!File.Exists(templatePath))
            {
                result.Success = false;
                result.ErrorMessage = $"Şablon dosyası bulunamadı: {templatePath}";
                return result;
            }

            var outputDirectory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            File.Copy(templatePath, outputPath, overwrite: true);

            await Task.Run(() =>
            {
                using var doc = WordprocessingDocument.Open(outputPath, isEditable: true);
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null) return;

                int replacementCount = 0;
                var replacedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (mainPart.Document?.Body != null)
                {
                    if (wells != null)
                    {
                        replacementCount += ProcessDynamicWellTables(mainPart.Document.Body, wells, replacedKeys);
                        replacementCount += ProcessDynamicWellParagraphs(mainPart.Document.Body, wells, replacedKeys);
                    }

                    replacementCount += ProcessElementContainers(mainPart.Document.Body, variableReplacements, replacedKeys);
                }

                foreach (var headerPart in mainPart.HeaderParts)
                {
                    if (headerPart.Header != null)
                    {
                        replacementCount += ProcessElementContainers(headerPart.Header, variableReplacements, replacedKeys);
                    }
                }

                foreach (var footerPart in mainPart.FooterParts)
                {
                    if (footerPart.Footer != null)
                    {
                        replacementCount += ProcessElementContainers(footerPart.Footer, variableReplacements, replacedKeys);
                    }
                }

                doc.Save();

                result.Success = true;
                result.OutputFilePath = outputPath;
                result.ReplacedCount = replacementCount;
                result.ReplacedVariables = replacedKeys.ToList();
            }, cancellationToken);

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Word dosyası işlenirken hata oluştu: {ex.Message}";
            return result;
        }
    }

    public async Task<BatchProcessResult> ProcessTemplateDirectoryAsync(
        string templateDirectoryPath,
        string outputDirectoryPath,
        IDictionary<string, string> variableReplacements,
        IEnumerable<DrillingWellItem>? wells = null,
        CancellationToken cancellationToken = default)
    {
        var batchResult = new BatchProcessResult();

        try
        {
            if (!Directory.Exists(templateDirectoryPath))
            {
                Directory.CreateDirectory(templateDirectoryPath);
            }

            if (!Directory.Exists(outputDirectoryPath))
            {
                Directory.CreateDirectory(outputDirectoryPath);
            }

            // Temp/~ lock dosyalarını hariç tutarak tüm .docx dosyalarını al
            var files = Directory.GetFiles(templateDirectoryPath, "*.docx", SearchOption.TopDirectoryOnly)
                                 .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                 .ToList();

            if (files.Count == 0)
            {
                batchResult.Success = false;
                batchResult.Errors.Add($"Şablon klasöründe (*.docx) uzantılı dosya bulunamadı: {templateDirectoryPath}");
                return batchResult;
            }

            foreach (var templateFile in files)
            {
                var fileName = Path.GetFileName(templateFile);
                var fileNameNoExt = Path.GetFileNameWithoutExtension(templateFile);
                var targetOutputPath = Path.Combine(outputDirectoryPath, $"{fileNameNoExt}_Düzenlenmiş.docx");

                var fileResult = await ProcessDocumentAsync(templateFile, targetOutputPath, variableReplacements, wells, cancellationToken);
                if (fileResult.Success)
                {
                    batchResult.ProcessedFileCount++;
                    batchResult.TotalReplacementsCount += fileResult.ReplacedCount;
                    batchResult.OutputFiles.Add(targetOutputPath);
                }
                else
                {
                    batchResult.Errors.Add($"{fileName}: {fileResult.ErrorMessage}");
                }
            }

            batchResult.Success = batchResult.ProcessedFileCount > 0;
            return batchResult;
        }
        catch (Exception ex)
        {
            batchResult.Success = false;
            batchResult.Errors.Add($"Klasör işlenirken genel hata: {ex.Message}");
            return batchResult;
        }
    }

    private int ProcessDynamicWellTables(
        Body body,
        IEnumerable<DrillingWellItem> wells,
        HashSet<string> replacedKeys)
    {
        var wellList = wells.ToList();
        int totalReplacements = 0;

        var tables = body.Descendants<Table>().ToList();
        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToList();
            foreach (var row in rows)
            {
                var rowText = row.InnerText;
                if (string.IsNullOrEmpty(rowText)) continue;

                bool isDynamicRow = rowText.Contains("{{WELL_") || rowText.Contains("{{SONDAJ_") || rowText.Contains("{{SK_");
                if (!isDynamicRow) continue;

                // Bu tablo satırı Yeraltı Suyu Ölçüm değişkenleri (tarih / su derinliği) içeriyor mu?
                bool isWaterMeasurementTable = rowText.Contains("WELL_MEASURE") ||
                                                rowText.Contains("WELL_WATER_DEPTH") ||
                                                rowText.Contains("WELL_GW_DEPTH") ||
                                                rowText.Contains("SONDAJ_OLCUM") ||
                                                rowText.Contains("SONDAJ_SU_DERINLIK");

                for (int i = 0; i < wellList.Count; i++)
                {
                    var w = wellList[i];

                    if (isWaterMeasurementTable)
                    {
                        // Ölçüm Tablosu: Her ölçüm için 1 satır çoğaltılır
                        var measurements = w.Measurements.Count > 0
                            ? w.Measurements
                            : new List<GroundwaterMeasurementItem> { new GroundwaterMeasurementItem { MeasurementDate = "", WaterDepth = w.IsCollapsed ? "Kuyu çökmüş" : "" } };

                        for (int j = 0; j < measurements.Count; j++)
                        {
                            var m = measurements[j];
                            var newRow = (TableRow)row.CloneNode(true);
                            var wellReplacements = GetWellReplacementsDictionary(w, m);

                            totalReplacements += ProcessElementContainers(newRow, wellReplacements, replacedKeys);
                            table.InsertBefore(newRow, row);
                        }
                    }
                    else
                    {
                        // Genel Sondaj Özeti Tablosu: Her kuyu için SADECE 1 satır üretilir (ölçüm sayısından bağımsız)
                        var newRow = (TableRow)row.CloneNode(true);
                        var wellReplacements = GetWellReplacementsDictionary(w);

                        totalReplacements += ProcessElementContainers(newRow, wellReplacements, replacedKeys);
                        table.InsertBefore(newRow, row);
                    }
                }

                row.Remove();
            }
        }

        return totalReplacements;
    }

    private int ProcessDynamicWellParagraphs(
        Body body,
        IEnumerable<DrillingWellItem> wells,
        HashSet<string> replacedKeys)
    {
        var wellList = wells.ToList();
        if (wellList.Count == 0) return 0;
        int totalReplacements = 0;

        // Tablo dışındaki tüm paragraflar taranır
        var allParagraphs = body.Descendants<Paragraph>()
            .Where(p => !p.Ancestors<Table>().Any())
            .ToList();

        // Ardışık dinamik paragraflar bloklar halinde gruplanır
        var blocks = new List<List<Paragraph>>();
        List<Paragraph>? currentBlock = null;

        foreach (var p in allParagraphs)
        {
            var pText = p.InnerText;
            bool isDynamic = !string.IsNullOrEmpty(pText) &&
                             (pText.Contains("{{WELL_") || pText.Contains("{{SONDAJ_") || pText.Contains("{{SK_"));

            if (isDynamic)
            {
                if (currentBlock == null)
                {
                    currentBlock = new List<Paragraph>();
                    blocks.Add(currentBlock);
                }
                currentBlock.Add(p);
            }
            else
            {
                currentBlock = null;
            }
        }

        // Her dinamik paragraf bloğu kuyu kuyu sırasıyla çoğaltılır
        foreach (var block in blocks)
        {
            if (block.Count == 0) continue;
            var anchor = block[0];
            var parent = anchor.Parent;
            if (parent == null) continue;

            for (int i = 0; i < wellList.Count; i++)
            {
                var w = wellList[i];
                var wellReplacements = GetWellReplacementsDictionary(w);

                foreach (var pInBlock in block)
                {
                    var newP = (Paragraph)pInBlock.CloneNode(true);
                    totalReplacements += ReplaceInParagraph(newP, wellReplacements, replacedKeys);
                    parent.InsertBefore(newP, anchor);
                }
            }

            foreach (var pInBlock in block)
            {
                pInBlock.Remove();
            }
        }

        return totalReplacements;
    }

    private Dictionary<string, string> GetWellReplacementsDictionary(DrillingWellItem w, GroundwaterMeasurementItem? m = null)
    {
        var targetMeasurement = m ?? w.Measurements.FirstOrDefault();
        string measureDate = targetMeasurement != null ? targetMeasurement.MeasurementDate : "";
        string waterDepth = targetMeasurement != null ? targetMeasurement.WaterDepth : (w.IsCollapsed ? "Kuyu çökmüş" : "");

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["WELL_NO"] = w.WellName,
            ["WELL_NAME"] = w.WellName,
            ["SONDAJ_NO"] = w.WellName,
            ["SK_NAME"] = w.WellName,

            ["WELL_LAT"] = w.Latitude,
            ["SONDAJ_ENLEM"] = w.Latitude,

            ["WELL_LON"] = w.Longitude,
            ["SONDAJ_BOYLAM"] = w.Longitude,

            ["WELL_GROUNDKOT"] = w.GroundKot,
            ["SONDAJ_ARAZI_KOTU"] = w.GroundKot,

            ["WELL_BOTTOMKOT"] = w.BottomKot,
            ["SONDAJ_DIP_KOTU"] = w.BottomKot,

            ["WELL_DEPTH"] = $"{w.DepthMeters:F2}m",
            ["SONDAJ_DERINLIK"] = $"{w.DepthMeters:F2}m",

            ["WELL_RESULT"] = w.GroundwaterStatus,
            ["WELL_WATER"] = w.GroundwaterStatus,
            ["SONDAJ_SONUC"] = w.GroundwaterStatus,

            ["WELL_COLLAPSED"] = w.IsCollapsed ? "Evet" : "Hayır",
            ["WELL_IS_COLLAPSED"] = w.IsCollapsed ? "Evet" : "Hayır",
            ["SONDAJ_COKME"] = w.IsCollapsed ? "Evet" : "Hayır",

            ["WELL_START"] = w.StartDate,
            ["SONDAJ_BASLANGIC"] = w.StartDate,

            ["WELL_FINISH"] = w.FinishDate,
            ["SONDAJ_BITIS"] = w.FinishDate,

            ["WELL_MEASURE_DATE"] = measureDate,
            ["WELL_MEASUREMENT_DATE"] = measureDate,
            ["SONDAJ_OLCUM_TARIHI"] = measureDate,

            ["WELL_WATER_DEPTH"] = waterDepth,
            ["WELL_GW_DEPTH"] = waterDepth,
            ["SONDAJ_SU_DERINLIK"] = waterDepth
        };
    }

    private int ProcessElementContainers(
        DocumentFormat.OpenXml.OpenXmlElement container,
        IDictionary<string, string> variableReplacements,
        HashSet<string> replacedKeys)
    {
        int count = 0;
        var paragraphs = container.Descendants<Paragraph>().ToList();
        foreach (var p in paragraphs)
        {
            count += ReplaceInParagraph(p, variableReplacements, replacedKeys);
        }
        return count;
    }

    private int ReplaceInParagraph(
        Paragraph paragraph,
        IDictionary<string, string> variableReplacements,
        HashSet<string> replacedKeys)
    {
        var fullText = paragraph.InnerText;
        if (string.IsNullOrEmpty(fullText) || !fullText.Contains("{{")) return 0;

        var matches = PlaceholderRegex.Matches(fullText);
        if (matches.Count == 0) return 0;

        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count == 0) return 0;

        int totalReplacements = 0;
        string updatedFullText = fullText;

        foreach (Match match in matches)
        {
            var keyName = match.Groups[1].Value;
            if (TryGetReplacementValue(variableReplacements, keyName, out var replacementValue))
            {
                updatedFullText = updatedFullText.Replace(match.Value, replacementValue);
                replacedKeys.Add(keyName);
                totalReplacements++;
            }
        }

        if (totalReplacements > 0)
        {
            texts[0].Text = updatedFullText;
            for (int i = 1; i < texts.Count; i++)
            {
                texts[i].Text = string.Empty;
            }
        }

        return totalReplacements;
    }

    private bool TryGetReplacementValue(IDictionary<string, string> replacements, string keyName, out string replacementValue)
    {
        if (replacements.TryGetValue(keyName, out var directValue))
        {
            replacementValue = directValue ?? string.Empty;
            return true;
        }

        var normalizedKey = NormalizeKey(keyName);
        foreach (var kvp in replacements)
        {
            var kvpNormalized = NormalizeKey(kvp.Key);
            if (string.Equals(kvpNormalized, normalizedKey, StringComparison.OrdinalIgnoreCase))
            {
                replacementValue = kvp.Value ?? string.Empty;
                return true;
            }
        }

        replacementValue = string.Empty;
        return false;
    }

    private static string NormalizeKey(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        return key.Replace('İ', 'I')
                  .Replace('ı', 'I')
                  .Replace('i', 'I')
                  .Replace('Ş', 'S')
                  .Replace('ş', 'S')
                  .Replace('Ğ', 'G')
                  .Replace('ğ', 'G')
                  .Replace('Ü', 'U')
                  .Replace('ü', 'U')
                  .Replace('Ö', 'O')
                  .Replace('ö', 'O')
                  .Replace('Ç', 'C')
                  .Replace('ç', 'C')
                  .Replace("_", string.Empty)
                  .Replace("-", string.Empty)
                  .Replace(" ", string.Empty)
                  .ToUpperInvariant();
    }
}
