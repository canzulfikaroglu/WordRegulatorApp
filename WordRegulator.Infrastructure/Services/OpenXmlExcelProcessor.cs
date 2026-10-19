using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Domain.Models;

namespace WordRegulator.Infrastructure.Services;

public class OpenXmlExcelProcessor : IExcelDocumentProcessor
{
    private static readonly Regex PlaceholderRegex = new Regex(
        @"\{\{\s*([a-zA-Z0-9_İıŞşĞğÜüÖöÇç\-]+)\s*\}\}",
        RegexOptions.Compiled);

    
    
    
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

            
            var templates = Directory
                .GetFiles(templateDirectoryPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .ToList();

            if (templates.Count == 0)
                return batchResult; 

            var wellList = wells.ToList();

            foreach (var templatePath in templates)
            {
                foreach (var well in wellList)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    
                    
                    
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

            
            File.Copy(templatePath, outputPath, overwrite: true);

            await Task.Run(() =>
            {
                using var doc = SpreadsheetDocument.Open(outputPath, isEditable: true);
                var workbookPart = doc.WorkbookPart;
                if (workbookPart == null) return;

                
                var wellReplacements = BuildWellReplacements(well, globalReplacements);

                int totalReplacements = 0;
                var replacedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    var sheet = worksheetPart.Worksheet;
                    var sheetData = sheet.GetFirstChild<SheetData>();
                    if (sheetData == null) continue;

                    
                    totalReplacements += ProcessWaterMeasurementRows(sheet, sheetData, well, replacedKeys, doc);

                    
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

    
    
    
    private static int ProcessWaterMeasurementRows(
        Worksheet sheet,
        SheetData sheetData,
        DrillingWellItem well,
        HashSet<string> replacedKeys,
        SpreadsheetDocument doc)
    {
        int totalReplacements = 0;
        var rows = sheetData.Elements<Row>().ToList();

        Row? templateRow = null;
        int templateRowIndex = -1;

        for (int i = 0; i < rows.Count; i++)
        {
            var rowText = GetRowText(rows[i], doc);
            if (rowText.Contains("WELL_MEASURE_DATE", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("WELL_WATER_DEPTH", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("SONDAJ_OLCUM_TARIHI", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("SONDAJ_SU_DERINLIK", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("{{DAY}}", StringComparison.OrdinalIgnoreCase) ||
                rowText.Contains("{{GUN}}", StringComparison.OrdinalIgnoreCase))
            {
                templateRow = rows[i];
                templateRowIndex = i;
                break;
            }
        }

        if (templateRow == null)
            return 0;

        uint originalRowNumber = templateRow.RowIndex?.Value ?? (uint)(templateRowIndex + 1);

        // Detect column letters for Depth, Date, and Day from templateRow
        string depthCol = "N";
        string dateCol = "O";
        string dayCol = "Q";

        Cell? cellDepth1 = null;
        Cell? cellDate1 = null;
        Cell? cellDay1 = null;

        foreach (var cell in templateRow.Elements<Cell>())
        {
            string? text = GetCellValue(cell, doc);
            if (string.IsNullOrEmpty(text)) continue;

            string col = GetColumnReference(cell.CellReference?.Value ?? "");
            if (string.IsNullOrEmpty(col)) continue;

            if (text.Contains("WELL_WATER_DEPTH", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("SONDAJ_SU_DERINLIK", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("WELL_GW_DEPTH", StringComparison.OrdinalIgnoreCase))
            {
                depthCol = col;
                cellDepth1 = cell;
            }
            else if (text.Contains("WELL_MEASURE_DATE", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("WELL_MEASUREMENT_DATE", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("SONDAJ_OLCUM_TARIHI", StringComparison.OrdinalIgnoreCase))
            {
                dateCol = col;
                cellDate1 = cell;
            }
            else if (text.Contains("{{DAY}}", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("{{GUN}}", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("WELL_DAY", StringComparison.OrdinalIgnoreCase) ||
                     text.Contains("ACIKLAMA", StringComparison.OrdinalIgnoreCase))
            {
                dayCol = col;
                cellDay1 = cell;
            }
        }

        cellDepth1 ??= GetOrCreateCell(templateRow, depthCol, originalRowNumber);
        cellDate1  ??= GetOrCreateCell(templateRow, dateCol, originalRowNumber);
        cellDay1   ??= GetOrCreateCell(templateRow, dayCol, originalRowNumber);

        uint? depthStyle = cellDepth1.StyleIndex?.Value;
        uint? dateStyle  = cellDate1.StyleIndex?.Value;
        uint? dayStyle   = cellDay1.StyleIndex?.Value;

        int measureCount = well.Measurements.Count;

        if (measureCount == 0)
        {
            string emptyDepth = well.IsCollapsed ? "Kuyu çökmüş" : string.Empty;
            SetCellValue(cellDepth1, emptyDepth, doc);
            SetCellValue(cellDate1, string.Empty, doc);
            SetCellValue(cellDay1, string.Empty, doc);
            replacedKeys.Add("WELL_WATER_DEPTH");
            replacedKeys.Add("WELL_MEASURE_DATE");
            replacedKeys.Add("DAY");
            return 3;
        }

        // 1. Measurement: Row 13 (templateRow in-place)
        SetCellValue(cellDepth1, well.Measurements[0].WaterDepth ?? string.Empty, doc);
        SetCellValue(cellDate1, well.Measurements[0].MeasurementDate ?? string.Empty, doc);
        SetCellValue(cellDay1, well.Measurements[0].Day ?? string.Empty, doc);
        replacedKeys.Add("WELL_WATER_DEPTH");
        replacedKeys.Add("WELL_MEASURE_DATE");
        replacedKeys.Add("DAY");
        totalReplacements += 3;

        // 2. Measurement: Row 14 (in-place in pre-allocated row slot)
        if (measureCount > 1)
        {
            uint row2Num = originalRowNumber + 1;
            var row2 = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == row2Num);
            if (row2 != null)
            {
                var cDepth2 = GetOrCreateCell(row2, depthCol, row2Num, depthStyle);
                var cDate2  = GetOrCreateCell(row2, dateCol, row2Num, dateStyle);
                var cDay2   = GetOrCreateCell(row2, dayCol, row2Num, dayStyle);

                SetCellValue(cDepth2, well.Measurements[1].WaterDepth ?? string.Empty, doc);
                SetCellValue(cDate2, well.Measurements[1].MeasurementDate ?? string.Empty, doc);
                SetCellValue(cDay2, well.Measurements[1].Day ?? string.Empty, doc);
                EnsureMergeCell(sheet, $"{dateCol}{row2Num}:P{row2Num}");
                EnsureMergeCell(sheet, $"{dayCol}{row2Num}:R{row2Num}");
                totalReplacements += 3;
            }
        }

        // 3. Measurement: Row 15 (in-place in pre-allocated row slot)
        if (measureCount > 2)
        {
            uint row3Num = originalRowNumber + 2;
            var row3 = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == row3Num);
            if (row3 != null)
            {
                var cDepth3 = GetOrCreateCell(row3, depthCol, row3Num, depthStyle);
                var cDate3  = GetOrCreateCell(row3, dateCol, row3Num, dateStyle);
                var cDay3   = GetOrCreateCell(row3, dayCol, row3Num, dayStyle);

                SetCellValue(cDepth3, well.Measurements[2].WaterDepth ?? string.Empty, doc);
                SetCellValue(cDate3, well.Measurements[2].MeasurementDate ?? string.Empty, doc);
                SetCellValue(cDay3, well.Measurements[2].Day ?? string.Empty, doc);
                EnsureMergeCell(sheet, $"{dateCol}{row3Num}:P{row3Num}");
                EnsureMergeCell(sheet, $"{dayCol}{row3Num}:R{row3Num}");
                totalReplacements += 3;
            }
        }

        // 4. and subsequent measurements: insert clean rows below Row 15 without duplicating left-side columns
        if (measureCount > 3)
        {
            int extraCount = measureCount - 3;
            uint shiftStart = originalRowNumber + 3;
            ShiftRowNumbers(sheetData, shiftStart, extraCount);
            ShiftMergeCells(sheet, shiftStart, extraCount);

            for (int i = 3; i < measureCount; i++)
            {
                uint extraRowNum = (uint)(originalRowNumber + i);
                var extraRow = new Row { RowIndex = extraRowNum };

                var cDepth = new Cell { CellReference = $"{depthCol}{extraRowNum}" };
                if (depthStyle.HasValue) cDepth.StyleIndex = depthStyle.Value;
                SetCellValue(cDepth, well.Measurements[i].WaterDepth ?? string.Empty, doc);
                extraRow.AppendChild(cDepth);

                var cDate = new Cell { CellReference = $"{dateCol}{extraRowNum}" };
                if (dateStyle.HasValue) cDate.StyleIndex = dateStyle.Value;
                SetCellValue(cDate, well.Measurements[i].MeasurementDate ?? string.Empty, doc);
                extraRow.AppendChild(cDate);

                var cDay = new Cell { CellReference = $"{dayCol}{extraRowNum}" };
                if (dayStyle.HasValue) cDay.StyleIndex = dayStyle.Value;
                SetCellValue(cDay, well.Measurements[i].Day ?? string.Empty, doc);
                extraRow.AppendChild(cDay);

                var afterRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value > extraRowNum);
                if (afterRow != null)
                    sheetData.InsertBefore(extraRow, afterRow);
                else
                    sheetData.AppendChild(extraRow);

                EnsureMergeCell(sheet, $"{dateCol}{extraRowNum}:P{extraRowNum}");
                EnsureMergeCell(sheet, $"{dayCol}{extraRowNum}:R{extraRowNum}");
                totalReplacements += 3;
            }
        }

        return totalReplacements;
    }

    private static Cell GetOrCreateCell(Row row, string colLetter, uint rowIndex, uint? fallbackStyle = null)
    {
        string cellRef = $"{colLetter}{rowIndex}";
        var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell != null) return cell;

        var newCell = new Cell { CellReference = cellRef };
        if (fallbackStyle.HasValue)
        {
            newCell.StyleIndex = fallbackStyle.Value;
        }

        Cell? refCell = null;
        int targetColIdx = ColumnLetterToIndex(colLetter);
        foreach (var existing in row.Elements<Cell>())
        {
            string existingCol = GetColumnReference(existing.CellReference?.Value ?? "");
            if (ColumnLetterToIndex(existingCol) > targetColIdx)
            {
                refCell = existing;
                break;
            }
        }

        if (refCell != null)
            row.InsertBefore(newCell, refCell);
        else
            row.AppendChild(newCell);

        return newCell;
    }

    private static int ColumnLetterToIndex(string col)
    {
        int index = 0;
        foreach (char c in col.ToUpperInvariant())
        {
            if (c >= 'A' && c <= 'Z')
                index = index * 26 + (c - 'A' + 1);
        }
        return index;
    }

    private static void EnsureMergeCell(Worksheet sheet, string reference)
    {
        var mergeCells = sheet.GetFirstChild<MergeCells>();
        if (mergeCells == null)
        {
            mergeCells = new MergeCells();
            sheet.InsertAfter(mergeCells, sheet.GetFirstChild<SheetData>());
        }

        if (!mergeCells.Elements<MergeCell>().Any(mc => string.Equals(mc.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)))
        {
            mergeCells.AppendChild(new MergeCell { Reference = reference });
        }
    }

    
    
    
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

    
    
    
    private static Row CloneRow(Row source, uint newRowIndex, SpreadsheetDocument doc)
    {
        var newRow = (Row)source.CloneNode(true);
        newRow.RowIndex = new DocumentFormat.OpenXml.UInt32Value(newRowIndex);

        
        foreach (var cell in newRow.Elements<Cell>())
        {
            if (cell.CellReference?.Value != null)
            {
                
                string colRef = GetColumnReference(cell.CellReference.Value);
                cell.CellReference = new DocumentFormat.OpenXml.StringValue($"{colRef}{newRowIndex}");
            }
        }

        return newRow;
    }

    private static void ShiftRowNumbers(SheetData sheetData, uint fromRowNumber, int delta)
    {
        if (delta <= 0) return;

        
        var subsequentRows = sheetData.Elements<Row>()
            .Where(r => r.RowIndex != null && r.RowIndex.Value >= fromRowNumber)
            .OrderByDescending(r => r.RowIndex!.Value)
            .ToList();

        foreach (var row in subsequentRows)
        {
            uint oldIndex = row.RowIndex!.Value;
            uint newIndex = (uint)(oldIndex + delta);
            row.RowIndex = new DocumentFormat.OpenXml.UInt32Value(newIndex);

            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                {
                    string colRef = GetColumnReference(cell.CellReference.Value);
                    cell.CellReference = new DocumentFormat.OpenXml.StringValue($"{colRef}{newIndex}");
                }
            }
        }
    }

    private static void ShiftMergeCells(Worksheet sheet, uint fromRowNumber, int delta)
    {
        if (delta <= 0) return;
        var mergeCells = sheet.GetFirstChild<MergeCells>();
        if (mergeCells == null) return;

        foreach (var mergeCell in mergeCells.Elements<MergeCell>())
        {
            if (mergeCell.Reference?.Value == null) continue;
            string refVal = mergeCell.Reference.Value;
            var parts = refVal.Split(':');
            if (parts.Length != 2) continue;

            string startCol = GetColumnReference(parts[0]);
            string endCol = GetColumnReference(parts[1]);

            string startRowNumStr = new string(parts[0].Where(char.IsDigit).ToArray());
            string endRowNumStr = new string(parts[1].Where(char.IsDigit).ToArray());

            if (uint.TryParse(startRowNumStr, out uint startRow) && uint.TryParse(endRowNumStr, out uint endRow))
            {
                if (startRow >= fromRowNumber)
                {
                    startRow = (uint)(startRow + delta);
                    endRow = (uint)(endRow + delta);
                    mergeCell.Reference = $"{startCol}{startRow}:{endCol}{endRow}";
                }
            }
        }
    }

    private static string GetColumnReference(string cellReference)
    {
        
        return new string(cellReference.Where(char.IsLetter).ToArray());
    }

    
    
    
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

    
    
    
    private static Dictionary<string, string> BuildWellReplacements(
        DrillingWellItem well,
        IDictionary<string, string> globalReplacements)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        
        foreach (var kvp in globalReplacements)
            d[kvp.Key] = kvp.Value;

        
        d["WELL_NO"]             = well.WellName;
        d["WELL_NAME"]           = well.WellName;
        d["SONDAJ_NO"]           = well.WellName;
        d["SK_NAME"]             = well.WellName;

        
        d["WELL_LAT"]            = well.Latitude;
        d["SONDAJ_ENLEM"]        = well.Latitude;
        d["KOORDINAT_X"]         = well.Latitude;
        d["WELL_X"]              = well.Latitude;

        d["WELL_LON"]            = well.Longitude;
        d["SONDAJ_BOYLAM"]       = well.Longitude;
        d["KOORDINAT_Y"]         = well.Longitude;
        d["WELL_Y"]              = well.Longitude;

        
        d["WELL_GROUNDKOT"]      = well.GroundKot;
        d["SONDAJ_ARAZI_KOTU"]   = well.GroundKot;
        d["SONDAJ_KOTU"]         = well.GroundKot;
        d["ZEMIN_KOTU"]          = well.GroundKot;

        d["WELL_BOTTOMKOT"]      = well.BottomKot;
        d["SONDAJ_DIP_KOTU"]     = well.BottomKot;

        d["WELL_DEPTH"]          = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR")) + " m";
        d["SONDAJ_DERINLIK"]     = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
        d["SONDAJ_DERINLIGI"]    = well.DepthMeters.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));

        
        d["WELL_START"]          = well.StartDate;
        d["SONDAJ_BASLANGIC"]    = well.StartDate;
        d["BASLAMA_TARIHI"]      = well.StartDate;
        d["WELL_FINISH"]         = well.FinishDate;
        d["SONDAJ_BITIS"]        = well.FinishDate;
        d["BITIS_TARIHI"]        = well.FinishDate;

        
        d["WELL_RESULT"]         = well.GroundwaterStatus;
        d["WELL_WATER"]          = well.GroundwaterStatus;
        d["SONDAJ_SONUC"]        = well.GroundwaterStatus;
        d["WELL_COLLAPSED"]      = well.IsCollapsed ? "Evet" : "Hayır";
        d["SONDAJ_COKME"]        = well.IsCollapsed ? "Evet" : "Hayır";

        
        var firstMeasure = well.Measurements.FirstOrDefault();
        d["WELL_MEASURE_DATE"]     = firstMeasure?.MeasurementDate ?? "";
        d["WELL_MEASUREMENT_DATE"] = firstMeasure?.MeasurementDate ?? "";
        d["SONDAJ_OLCUM_TARIHI"]   = firstMeasure?.MeasurementDate ?? "";
        d["WELL_WATER_DEPTH"]      = firstMeasure?.WaterDepth ?? (well.IsCollapsed ? "Kuyu çökmüş" : "");
        d["WELL_GW_DEPTH"]         = firstMeasure?.WaterDepth ?? "";
        d["SONDAJ_SU_DERINLIK"]    = firstMeasure?.WaterDepth ?? "";
        d["DAY"]                   = firstMeasure?.Day ?? "";
        d["GUN"]                   = firstMeasure?.Day ?? "";
        d["WELL_DAY"]              = firstMeasure?.Day ?? "";
        d["ACIKLAMA"]              = firstMeasure?.Day ?? "";

        return d;
    }

    
    
    
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
