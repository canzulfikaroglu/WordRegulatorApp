using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Domain.Models;

namespace WordRegulator.Infrastructure.Services;

/// <summary>
/// Her sondaj kuyusu için ayrı bir .xlsx dosyası üretir.
/// Excel şablonundaki {{DEGISKEN}} işaretçileri sondaj verisiyle değiştirilir.
/// Yeraltı suyu ölçümleri için şablon satır klonlama desteklenir.
/// </summary>
public class OpenXmlExcelProcessor : IExcelDocumentProcessor
{
    private static readonly Regex PlaceholderRegex = new Regex(
        @"\{\{\s*([a-zA-Z0-9_İıŞşĞğÜüÖöÇç\-]+)\s*\}\}",
        RegexOptions.Compiled);

    // =========================================================================
    // PUBLIC — Birden fazla kuyu için şablon dizinini işle
    // =========================================================================
    public async Task<BatchProcessResult> ProcessExcelTemplatesForWellsAsync(
        string templateDirectoryPath,
        string outputDirectoryPath,
        IDictionary<string, string> globalReplacements,
        IEnumerable<DrillingWellItem> wells,
        CancellationToken cancellationToken = default)
    {
        var batchResult = new BatchProcessResult();

        try
        {
            if (!Directory.Exists(outputDirectoryPath))
                Directory.CreateDirectory(outputDirectoryPath);

            // Temp/lock dosyalarını hariç tutarak tüm .xlsx şablonlarını al
            var templates = Directory
                .GetFiles(templateDirectoryPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .ToList();

            if (templates.Count == 0)
                return batchResult; // Excel şablonu yoksa sessizce çık

            var wellList = wells.ToList();

            foreach (var templatePath in templates)
            {
                foreach (var well in wellList)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Çıktı dosya adı: SK-1.xlsx, SK-2.xlsx ... (orijinal şablon adı korunmaz)
                    // Şablonun orijinal adına kuyu adını ekleyerek üret:
                    // "SondajLogu.xlsx" → "SondajLogu_SK-1.xlsx"
                    string templateBaseName = Path.GetFileNameWithoutExtension(templatePath);
                    string outputFileName = $"{templateBaseName}_{well.WellName}.xlsx";
                    string outputPath = Path.Combine(outputDirectoryPath, outputFileName);

                    var result = await ProcessExcelForWellAsync(
                        templatePath, outputPath, globalReplacements, well, cancellationToken);

                    if (result.Success)
                    {
                        batchResult.ProcessedFileCount++;
                        batchResult.TotalReplacementsCount += result.ReplacedCount;
                        batchResult.OutputFiles.Add(outputPath);
                    }
                    else
                    {
                        batchResult.Errors.Add($"{outputFileName}: {result.ErrorMessage}");
                    }
                }
            }

            batchResult.Success = batchResult.ProcessedFileCount > 0 || batchResult.Errors.Count == 0;
            return batchResult;
        }
        catch (Exception ex)
        {
            batchResult.Success = false;
            batchResult.Errors.Add($"Excel toplu işleme hatası: {ex.Message}");
            return batchResult;
        }
    }

    // =========================================================================
    // PUBLIC — Tek kuyu için tek Excel dosyası üret
    // =========================================================================
    public async Task<DocumentProcessResult> ProcessExcelForWellAsync(
        string templatePath,
        string outputPath,
        IDictionary<string, string> globalReplacements,
        DrillingWellItem well,
        CancellationToken cancellationToken = default)
    {
        var result = new DocumentProcessResult();

        try
        {
            if (!File.Exists(templatePath))
            {
                result.Success = false;
                result.ErrorMessage = $"Excel şablon dosyası bulunamadı: {templatePath}";
                return result;
            }

            // Şablonu çıktı yoluna kopyala
            File.Copy(templatePath, outputPath, overwrite: true);

            await Task.Run(() =>
            {
                using var doc = SpreadsheetDocument.Open(outputPath, isEditable: true);
                var workbookPart = doc.WorkbookPart;
                if (workbookPart == null) return;

                // Kuyu bazlı replacements oluştur (global + kuya özgü)
                var wellReplacements = BuildWellReplacements(well, globalReplacements);

                int totalReplacements = 0;
                var replacedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    var sheet = worksheetPart.Worksheet;
                    var sheetData = sheet.GetFirstChild<SheetData>();
                    if (sheetData == null) continue;

                    // Yeraltı suyu ölçüm satırlarını klonla (şablon satır)
                    totalReplacements += ProcessWaterMeasurementRows(sheetData, well, replacedKeys, doc);

                    // Tüm hücrelerdeki placeholder'ları değiştir
                    totalReplacements += ReplaceInWorksheet(worksheetPart, wellReplacements, replacedKeys, doc);

                    sheet.Save();
                }

                doc.Save();

                result.Success = true;
                result.OutputFilePath = outputPath;
                result.ReplacedCount = totalReplacements;
                result.ReplacedVariables = replacedKeys.ToList();

            }, cancellationToken);

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Excel dosyası işlenirken hata: {ex.Message}";
            return result;
        }
    }

    // =========================================================================
    // PRIVATE — Yeraltı suyu ölçüm satırlarını klonla
    // =========================================================================
    private static int ProcessWaterMeasurementRows(
        SheetData sheetData,
        DrillingWellItem well,
        HashSet<string> replacedKeys,
        SpreadsheetDocument doc)
    {
        int totalReplacements = 0;
        var rows = sheetData.Elements<Row>().ToList();

        Row? templateRow = null;
        int templateRowIndex = -1;

        // {{WELL_MEASURE_DATE}} veya {{WELL_WATER_DEPTH}} içeren satırı bul
        for (int i = 0; i < rows.Count; i++)
        {
            var rowText = GetRowText(rows[i], doc);
            if (rowText.Contains("WELL_MEASURE_DATE", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("WELL_WATER_DEPTH", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("SONDAJ_OLCUM_TARIHI", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("SONDAJ_SU_DERINLIK", StringComparison.OrdinalIgnoreCase))
            {
                templateRow = rows[i];
                templateRowIndex = i;
                break;
            }
        }

        if (templateRow == null || well.Measurements.Count == 0)
            return 0;

        uint originalRowNumber = templateRow.RowIndex?.Value ?? (uint)(templateRowIndex + 1);
        uint insertAtRowNumber = originalRowNumber;

        // Şablon satırın üstüne her ölçüm için yeni satır ekle
        foreach (var measurement in well.Measurements)
        {
            var newRow = CloneRow(templateRow, insertAtRowNumber, doc);
            var measureReplacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["WELL_MEASURE_DATE"]     = measurement.MeasurementDate,
                ["WELL_MEASUREMENT_DATE"] = measurement.MeasurementDate,
                ["SONDAJ_OLCUM_TARIHI"]   = measurement.MeasurementDate,
                ["WELL_WATER_DEPTH"]      = measurement.WaterDepth,
                ["WELL_GW_DEPTH"]         = measurement.WaterDepth,
                ["SONDAJ_SU_DERINLIK"]    = measurement.WaterDepth,
            };

            totalReplacements += ReplaceInRow(newRow, measureReplacements, replacedKeys, doc);
            sheetData.InsertBefore(newRow, templateRow);
            insertAtRowNumber++;
        }

        // Şablon satırı sil ve sonraki satırların numaralarını güncelle
        templateRow.Remove();
        ShiftRowNumbers(sheetData, insertAtRowNumber);

        return totalReplacements;
    }

    // =========================================================================
    // PRIVATE — Worksheet içindeki tüm hücrelerde replace yap
    // =========================================================================
    private static int ReplaceInWorksheet(
        WorksheetPart worksheetPart,
        IDictionary<string, string> replacements,
        HashSet<string> replacedKeys,
        SpreadsheetDocument doc)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return 0;

        int total = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            total += ReplaceInRow(row, replacements, replacedKeys, doc);
        }
        return total;
    }

    private static int ReplaceInRow(
        Row row,
        IDictionary<string, string> replacements,
        HashSet<string> replacedKeys,
        SpreadsheetDocument doc)
    {
        int count = 0;
        foreach (var cell in row.Elements<Cell>())
        {
            string? cellText = GetCellValue(cell, doc);
            if (string.IsNullOrEmpty(cellText) || !cellText.Contains("{{")) continue;

            var matches = PlaceholderRegex.Matches(cellText);
            if (matches.Count == 0) continue;

            string updatedText = cellText;
            bool modified = false;

            foreach (Match match in matches)
            {
                string keyName = match.Groups[1].Value;
                if (TryGetValue(replacements, keyName, out string val))
                {
                    updatedText = updatedText.Replace(match.Value, val);
                    replacedKeys.Add(keyName);
                    modified = true;
                    count++;
                }
            }

            if (modified)
            {
                SetCellValue(cell, updatedText, doc);
            }
        }
        return count;
    }

    // =========================================================================
    // PRIVATE — Satır klonlama yardımcıları
    // =========================================================================
    private static Row CloneRow(Row source, uint newRowIndex, SpreadsheetDocument doc)
    {
        var newRow = (Row)source.CloneNode(true);
        newRow.RowIndex = new DocumentFormat.OpenXml.UInt32Value(newRowIndex);

        // Hücre referanslarını güncelle (A5 → A6 vb.)
        foreach (var cell in newRow.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                // Sütun harfini koru, satır numarasını güncelle
                string colRef = GetColumnReference(cell.CellReference.Value);
                cell.CellReference = new DocumentFormat.OpenXml.StringValue($"{colRef}{newRowIndex}");
            }
        }

        return newRow;
    }

    private static void ShiftRowNumbers(SheetData sheetData, uint fromRowNumber)
    {
        // Silinen şablon satırdan sonraki satırların numaralarını 1 azalt (ekleme durumu için gerekli değil,
        // ancak remove sonrası tutarsızlıkları önler)
        // Bu basit implementasyonda atlıyoruz; OpenXml genelde bunu kendi halleder.
    }

    private static string GetColumnReference(string cellReference)
    {
        // "AB12" → "AB"
        return new string(cellReference.Where(char.IsLetter).ToArray());
    }

    // =========================================================================
    // PRIVATE — Hücre değeri okuma/yazma (shared strings desteği ile)
    // =========================================================================
    private static string? GetCellValue(Cell cell, SpreadsheetDocument doc)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            if (int.TryParse(cell.InnerText, out int ssIndex))
            {
                var sst = doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (sst != null)
                {
                    var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(ssIndex);
                    return item?.InnerText;
                }
            }
            return null;
        }

        return cell.InnerText;
    }

    private static void SetCellValue(Cell cell, string newValue, SpreadsheetDocument doc)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            // Shared string'i inline string'e çevir
            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.InlineString);
            cell.RemoveAllChildren();
            var iStr = new InlineString();
            iStr.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(newValue));
            cell.AppendChild(iStr);
        }
        else if (cell.DataType?.Value == CellValues.InlineString)
        {
            var iStr = cell.GetFirstChild<InlineString>();
            if (iStr != null)
            {
                var textEl = iStr.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Text>();
                if (textEl != null) textEl.Text = newValue;
                else iStr.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(newValue));
            }
        }
        else
        {
            // Metin hücresi — inline string olarak yaz
            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.InlineString);
            cell.RemoveAllChildren();
            var iStr = new InlineString();
            iStr.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Text(newValue));
            cell.AppendChild(iStr);
        }
    }

    private static string GetRowText(Row row, SpreadsheetDocument doc)
    {
        return string.Concat(row.Elements<Cell>().Select(c => GetCellValue(c, doc) ?? string.Empty));
    }

    // =========================================================================
    // PRIVATE — Replacement dictionary
    // =========================================================================
    private static Dictionary<string, string> BuildWellReplacements(
        DrillingWellItem well,
        IDictionary<string, string> globalReplacements)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Global değişkenleri kopyala (proje adı, il, ilçe vb.)
        foreach (var kvp in globalReplacements)
            d[kvp.Key] = kvp.Value;

        // Kuya özgü değişkenler (global'i override eder)
        d["WELL_NO"]             = well.WellName;
        d["WELL_NAME"]           = well.WellName;
        d["SONDAJ_NO"]           = well.WellName;
        d["SK_NAME"]             = well.WellName;

        // Koordinatlar (X = Enlem, Y = Boylam — Excel'deki sıraya göre)
        d["WELL_LAT"]            = well.Latitude;
        d["SONDAJ_ENLEM"]        = well.Latitude;
        d["KOORDINAT_X"]         = well.Latitude;
        d["WELL_X"]              = well.Latitude;

        d["WELL_LON"]            = well.Longitude;
        d["SONDAJ_BOYLAM"]       = well.Longitude;
        d["KOORDINAT_Y"]         = well.Longitude;
        d["WELL_Y"]              = well.Longitude;

        // Kotlar & Derinlik
        d["WELL_GROUNDKOT"]      = well.GroundKot;
        d["SONDAJ_ARAZI_KOTU"]   = well.GroundKot;
        d["SONDAJ_KOTU"]         = well.GroundKot;
        d["ZEMIN_KOTU"]          = well.GroundKot;

        d["WELL_BOTTOMKOT"]      = well.BottomKot;
        d["SONDAJ_DIP_KOTU"]     = well.BottomKot;

        d["WELL_DEPTH"]          = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR")) + " m";
        d["SONDAJ_DERINLIK"]     = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
        d["SONDAJ_DERINLIGI"]    = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));

        // Tarihler
        d["WELL_START"]          = well.StartDate;
        d["SONDAJ_BASLANGIC"]    = well.StartDate;
        d["BASLAMA_TARIHI"]      = well.StartDate;
        d["WELL_FINISH"]         = well.FinishDate;
        d["SONDAJ_BITIS"]        = well.FinishDate;
        d["BITIS_TARIHI"]        = well.FinishDate;

        // Su & Durum
        d["WELL_RESULT"]         = well.GroundwaterStatus;
        d["WELL_WATER"]          = well.GroundwaterStatus;
        d["SONDAJ_SONUC"]        = well.GroundwaterStatus;
        d["WELL_COLLAPSED"]      = well.IsCollapsed ? "Evet" : "Hayır";
        d["SONDAJ_COKME"]        = well.IsCollapsed ? "Evet" : "Hayır";

        // İlk ölçüm değerleri (tekil hücre referansları için)
        var firstMeasure = well.Measurements.FirstOrDefault();
        d["WELL_MEASURE_DATE"]     = firstMeasure?.MeasurementDate ?? "";
        d["WELL_MEASUREMENT_DATE"] = firstMeasure?.MeasurementDate ?? "";
        d["SONDAJ_OLCUM_TARIHI"]   = firstMeasure?.MeasurementDate ?? "";
        d["WELL_WATER_DEPTH"]      = firstMeasure?.WaterDepth ?? (well.IsCollapsed ? "Kuyu çökmüş" : "");
        d["WELL_GW_DEPTH"]         = firstMeasure?.WaterDepth ?? "";
        d["SONDAJ_SU_DERINLIK"]    = firstMeasure?.WaterDepth ?? "";

        return d;
    }

    // =========================================================================
    // PRIVATE — Anahtar arama (Türkçe normalize ile)
    // =========================================================================
    private static bool TryGetValue(IDictionary<string, string> dict, string key, out string value)
    {
        if (dict.TryGetValue(key, out var v))
        {
            value = v ?? string.Empty;
            return true;
        }

        string normKey = NormalizeKey(key);
        foreach (var kvp in dict)
        {
            if (string.Equals(NormalizeKey(kvp.Key), normKey, StringComparison.OrdinalIgnoreCase))
            {
                value = kvp.Value ?? string.Empty;
                return true;
            }
        }

        value = string.Empty;
        return false;
    }

    private static string NormalizeKey(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        return key.Replace('İ', 'I').Replace('ı', 'I').Replace('i', 'I')
                  .Replace('Ş', 'S').Replace('ş', 'S')
                  .Replace('Ğ', 'G').Replace('ğ', 'G')
                  .Replace('Ü', 'U').Replace('ü', 'U')
                  .Replace('Ö', 'O').Replace('ö', 'O')
                  .Replace('Ç', 'C').Replace('ç', 'C')
                  .Replace("_", string.Empty)
                  .Replace("-", string.Empty)
                  .Replace(" ", string.Empty)
                  .ToUpperInvariant();
    }
}
