using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordRegulator.Application.Services;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Helpers;
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
        IEnumerable<Vs30SeriesItem>? vs30Series = null,
        IEnumerable<FormationDefinition>? formations = null,
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
                        replacementCount += ProcessDynamicPressuremeterTables(mainPart.Document.Body, wells, replacedKeys);
                    }

                    if (vs30Series != null)
                    {
                        replacementCount += ProcessDynamicVs30Tables(mainPart.Document.Body, vs30Series, replacedKeys);
                        replacementCount += ProcessDynamicVs30Paragraphs(mainPart.Document.Body, vs30Series, replacedKeys);
                        replacementCount += ProcessDynamicSeismicTables(mainPart.Document.Body, vs30Series, replacedKeys);
                    }

                    replacementCount += ProcessDynamicMicrotremorTables(mainPart.Document.Body, variableReplacements, replacedKeys);

                    if (formations != null && formations.Any())
                    {
                        replacementCount += ProcessDynamicFormationParagraphs(mainPart.Document.Body, formations, replacedKeys);
                    }
                    else
                    {
                        var inferred = InferFormationsFromReplacements(variableReplacements);
                        if (inferred.Count > 0)
                        {
                            replacementCount += ProcessDynamicFormationParagraphs(mainPart.Document.Body, inferred, replacedKeys);
                        }
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
        IEnumerable<Vs30SeriesItem>? vs30Series = null,
        IEnumerable<FormationDefinition>? formations = null,
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

                var fileResult = await ProcessDocumentAsync(templateFile, targetOutputPath, variableReplacements, wells, vs30Series, formations, cancellationToken);
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

                
                
                bool isDynamicRow = (rowText.Contains("{{WELL_") || rowText.Contains("{{SONDAJ_") || rowText.Contains("{{SK_"))
                                 && !rowText.Contains("{{SONDAJ_PM_") && !rowText.Contains("{{SK_PM_");
                if (!isDynamicRow) continue;

                
                bool isWaterMeasurementTable = rowText.Contains("WELL_MEASURE") ||
                                                rowText.Contains("WELL_WATER_DEPTH") ||
                                                rowText.Contains("WELL_GW_DEPTH") ||
                                                rowText.Contains("SONDAJ_OLCUM") ||
                                                rowText.Contains("SONDAJ_SU_DERINLIK") ||
                                                rowText.Contains("{{DAY}}") ||
                                                rowText.Contains("{{GUN}}");

                for (int i = 0; i < wellList.Count; i++)
                {
                    var w = wellList[i];

                    if (isWaterMeasurementTable)
                    {
                        
                        var measurements = w.Measurements.Count > 0
                            ? w.Measurements
                            : new List<GroundwaterMeasurementItem> { new GroundwaterMeasurementItem { MeasurementDate = "", WaterDepth = w.IsCollapsed ? "Kuyu çökmüş" : "" } };

                        for (int j = 0; j < measurements.Count; j++)
                        {
                            var m = measurements[j];
                            var newRow = (TableRow)row.CloneNode(true);
                            var wellReplacements = GetWellReplacementsDictionary(w, m, i + 1);

                            totalReplacements += ProcessElementContainers(newRow, wellReplacements, replacedKeys);
                            table.InsertBefore(newRow, row);
                        }
                    }
                    else
                    {
                        
                        var newRow = (TableRow)row.CloneNode(true);
                        var wellReplacements = GetWellReplacementsDictionary(w, null, i + 1);

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

        
        var allParagraphs = body.Descendants<Paragraph>()
            .Where(p => !p.Ancestors<Table>().Any())
            .ToList();

        
        var blocks = new List<List<Paragraph>>();
        List<Paragraph>? currentBlock = null;

        foreach (var p in allParagraphs)
        {
            var pText = p.InnerText;
            
            bool isDynamic = !string.IsNullOrEmpty(pText) &&
                             (pText.Contains("{{WELL_") || pText.Contains("{{SONDAJ_") || pText.Contains("{{SK_"))
                             && !pText.Contains("{{SONDAJ_PM_") && !pText.Contains("{{SK_PM_");

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

        
        foreach (var block in blocks)
        {
            if (block.Count == 0) continue;
            var anchor = block[0];
            var parent = anchor.Parent;
            if (parent == null) continue;

            for (int i = 0; i < wellList.Count; i++)
            {
                var w = wellList[i];
                var wellReplacements = GetWellReplacementsDictionary(w, null, i + 1);

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

    private int ProcessDynamicVs30Tables(
        Body body,
        IEnumerable<Vs30SeriesItem> vs30Series,
        HashSet<string> replacedKeys)
    {
        var seriesList = vs30Series.ToList();
        if (seriesList.Count == 0) return 0;
        int totalReplacements = 0;

        
        var bodyChildren = body.ChildElements.ToList();

        foreach (var bodyChild in bodyChildren)
        {
            if (bodyChild is not Table table) continue;

            var rows = table.Elements<TableRow>().ToList();
            if (rows.Count == 0) continue;

            
            bool hasLayerTemplate = rows.Any(r =>
                r.InnerText.Contains("{{H}}") ||
                r.InnerText.Contains("{{LAYER_") ||
                r.InnerText.Contains("{{TABAKA_"));

            if (!hasLayerTemplate) continue;

            
            
            Paragraph? titleParagraph = null;
            var nextSib = table.NextSibling();
            while (nextSib != null)
            {
                if (nextSib is Paragraph nextPara)
                {
                    string pText = nextPara.InnerText;
                    if (pText.Contains("{{SERIES_") || pText.Contains("{{SERIM_"))
                        titleParagraph = nextPara;
                    break;
                }
                nextSib = nextSib.NextSibling();
            }

            
            foreach (var s in seriesList)
            {
                
                var clonedTable = (Table)table.CloneNode(true);
                var clonedRows = clonedTable.Elements<TableRow>().ToList();

                
                var templateDataRow = clonedRows.FirstOrDefault(r =>
                    r.InnerText.Contains("{{H}}") ||
                    r.InnerText.Contains("{{LAYER_") ||
                    r.InnerText.Contains("{{TABAKA_"));

                if (templateDataRow != null)
                {
                    bool isHTable = templateDataRow.InnerText.Contains("{{H}}");
                    bool isVsurTable = !isHTable && templateDataRow.InnerText.Contains("{{LAYER_VSUR}}");

                    if (isHTable)
                    {
                        // First table: {{H}} / {{LAYER_VS}} / {{SERIES_RESULT}} — uses Layers with Vs30 h-calc
                        double sumThickness = 0;
                        bool firstValidInserted = false;
                        var insertedHRows = new List<TableRow>();

                        // Identify which column index contains the Vs30 result before cloning
                        int vs30ColIndex = -1;
                        var tplCells = templateDataRow.Elements<TableCell>().ToList();
                        for (int ci = 0; ci < tplCells.Count; ci++)
                        {
                            string ct = tplCells[ci].InnerText;
                            if (ct.Contains("{{SERIES_RESULT") || ct.Contains("{{SERIES_VS30") || ct.Contains("{{SERIES_A"))
                            {
                                vs30ColIndex = ci;
                                break;
                            }
                        }

                        for (int j = 0; j < s.Layers.Count; j++)
                        {
                            var l = s.Layers[j];
                            double h_i;

                            if (j < s.Layers.Count - 1)
                            {
                                // effectiveStart: if Df falls within this layer, start from Df; otherwise layer's own start
                                double effectiveStart = Math.Max(s.Df, l.StartDepth);
                                h_i = Math.Max(0, l.EndDepth - effectiveStart);
                            }
                            else
                            {
                                h_i = Math.Max(0, 30.0 - sumThickness);
                            }

                            sumThickness += h_i;

                            // Skip layers that contribute nothing
                            if (h_i <= 0) continue;

                            bool isFirstRow = !firstValidInserted;
                            firstValidInserted = true;

                            var newRow = (TableRow)templateDataRow.CloneNode(true);
                            var layerReplacements = GetSeriesReplacementsDictionary(s, l, h_i, isFirstRow);
                            totalReplacements += ProcessElementContainers(newRow, layerReplacements, replacedKeys);
                            clonedTable.InsertBefore(newRow, templateDataRow);
                            insertedHRows.Add(newRow);
                        }

                        // Apply vertical merge on the Vs30 column so the value spans all rows
                        if (vs30ColIndex >= 0 && insertedHRows.Count > 1)
                            ApplyVerticalMerge(insertedHRows, vs30ColIndex);
                    }
                    else if (isVsurTable)
                    {
                        // Second table: {{LAYER_RANGE}} / {{LAYER_VSUR}} — uses Layers2
                        var layers2 = s.Layers2.Count > 0 ? s.Layers2 : s.Layers;
                        for (int j = 0; j < layers2.Count; j++)
                        {
                            var l = layers2[j];
                            var newRow = (TableRow)templateDataRow.CloneNode(true);
                            var layerReplacements = GetSeriesReplacementsDictionary(s, l, 0, false);
                            totalReplacements += ProcessElementContainers(newRow, layerReplacements, replacedKeys);
                            clonedTable.InsertBefore(newRow, templateDataRow);
                        }
                    }
                    else
                    {
                        // Generic LAYER_ table — uses Layers
                        for (int j = 0; j < s.Layers.Count; j++)
                        {
                            var l = s.Layers[j];
                            var newRow = (TableRow)templateDataRow.CloneNode(true);
                            var layerReplacements = GetSeriesReplacementsDictionary(s, l, 0, false);
                            totalReplacements += ProcessElementContainers(newRow, layerReplacements, replacedKeys);
                            clonedTable.InsertBefore(newRow, templateDataRow);
                        }
                    }

                    
                    templateDataRow.Remove();
                }

                
                var seriesOnlyReplacements = GetSeriesReplacementsDictionary(s);
                foreach (var row in clonedTable.Elements<TableRow>())
                {
                    
                    totalReplacements += ProcessElementContainers(row, seriesOnlyReplacements, replacedKeys);
                }

                
                body.InsertBefore(clonedTable, table);
                totalReplacements++;

                
                if (titleParagraph != null)
                {
                    var clonedPara = (Paragraph)titleParagraph.CloneNode(true);
                    totalReplacements += ReplaceInParagraph(clonedPara, seriesOnlyReplacements, replacedKeys);
                    body.InsertBefore(clonedPara, table);
                }
            }

            
            table.Remove();

            
            if (titleParagraph != null)
                titleParagraph.Remove();
        }

        return totalReplacements;
    }

    private int ProcessDynamicVs30Paragraphs(
        Body body,
        IEnumerable<Vs30SeriesItem> vs30Series,
        HashSet<string> replacedKeys)
    {
        var seriesList = vs30Series.ToList();
        if (seriesList.Count == 0) return 0;
        int totalReplacements = 0;

        var allParagraphs = body.Descendants<Paragraph>()
            .Where(p => !p.Ancestors<Table>().Any())
            .ToList();

        var blocks = new List<List<Paragraph>>();
        List<Paragraph>? currentBlock = null;

        foreach (var p in allParagraphs)
        {
            var pText = p.InnerText;
            bool isDynamic = !string.IsNullOrEmpty(pText) &&
                             (pText.Contains("{{SERIES_") || pText.Contains("{{SERIM_"));

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

        foreach (var block in blocks)
        {
            if (block.Count == 0) continue;
            var anchor = block[0];
            var parent = anchor.Parent;
            if (parent == null) continue;

            for (int i = 0; i < seriesList.Count; i++)
            {
                var s = seriesList[i];
                var seriesReplacements = GetSeriesReplacementsDictionary(s);

                foreach (var pInBlock in block)
                {
                    var newP = (Paragraph)pInBlock.CloneNode(true);
                    totalReplacements += ReplaceInParagraph(newP, seriesReplacements, replacedKeys);
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

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    private int ProcessDynamicSeismicTables(
        Body body,
        IEnumerable<Vs30SeriesItem> vs30Series,
        HashSet<string> replacedKeys)
    {
        var seriesList = vs30Series.ToList();
        if (seriesList.Count == 0) return 0;

        int totalReplacements = 0;

        
        
        static bool IsSeismicTable(Table t)
        {
            if (t.Descendants<Table>().Any()) return false;

            string text = t.InnerText;
            if (string.IsNullOrEmpty(text)) return false;

            return text.Contains("{{SS_1_") || text.Contains("{{SS_2_")
                || text.Contains("{{SEİSMİCCOORDİNATESS-1") || text.Contains("{{SEİSMİCCOORDİNATESS-2")
                || text.Contains("{{POİSSONRATİO1S1") || text.Contains("{{POİSSONRATİO1S2")
                || text.Contains("{{SHEARMODULUS1S1") || text.Contains("{{SHEARMODULUS1S2")
                || text.Contains("{{BULKMODULE1S1") || text.Contains("{{BULKMODULE1S2")
                || text.Contains("{{ELASTİCİTYMODULUS1S1") || text.Contains("{{ELASTİCİTYMODULUS1S2");
        }

        var tables = body.Descendants<Table>().Where(IsSeismicTable).ToList();
        if (tables.Count == 0) return 0;

        
        var pairs = new List<(Vs30SeriesItem? s1, Vs30SeriesItem? s2)>();
        for (int i = 0; i < seriesList.Count; i += 2)
        {
            pairs.Add((
                i < seriesList.Count ? seriesList[i] : null,
                i + 1 < seriesList.Count ? seriesList[i + 1] : null
            ));
        }

        foreach (var table in tables)
        {
            var parent = table.Parent;
            if (parent == null) continue;

            for (int pairIdx = 0; pairIdx < pairs.Count; pairIdx++)
            {
                var (s1, s2) = pairs[pairIdx];
                int num1 = s1?.SeriesNumber ?? (pairIdx * 2 + 1);
                int num2 = s2?.SeriesNumber ?? (pairIdx * 2 + 2);

                
                var cloned = (Table)table.CloneNode(true);

                
                
                var tblPr = cloned.GetFirstChild<TableProperties>();
                if (tblPr != null)
                {
                    tblPr.GetFirstChild<TablePositionProperties>()?.Remove();
                }

                
                // Normalize any split placeholders in all paragraphs of this cloned table
                foreach (var p in cloned.Descendants<Paragraph>())
                {
                    NormalizeSplitPlaceholders(p);
                }

                if (pairIdx > 0)
                {
                    ReplaceLiteralHeaders(cloned, num1, num2, s2 != null);
                }

                // If odd number of series and s2 is null, trim the right-hand side (unused 2nd series)
                if (s2 == null)
                {
                    TrimSeismicTableForSingleSeries(cloned);
                }

                var pairReplacements = BuildSeismicReplacements(s1, num1, s2, num2);
                foreach (var row in cloned.Elements<TableRow>())
                {
                    totalReplacements += ProcessElementContainers(row, pairReplacements, replacedKeys);
                }

                parent.InsertBefore(cloned, table);

                
                if (pairIdx < pairs.Count - 1)
                {
                    var separator = new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines { Before = "240", After = "240", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                        ),
                        new Run(
                            new RunProperties(new FontSize { Val = "22" }),
                            new Text(" ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }
                        )
                    );
                    parent.InsertBefore(separator, table);
                }

                totalReplacements++;
            }

            
            table.Remove();
        }

        return totalReplacements;
    }

    /// <summary>
    /// Applies OpenXML vertical cell merge to a specific column across a list of rows.
    /// The first row gets VerticalMerge(Restart) and its content is vertically centered.
    /// Subsequent rows get VerticalMerge() (Continue) — their content is hidden by the merge.
    /// </summary>
    private static void ApplyVerticalMerge(List<TableRow> rows, int colIndex)
    {
        for (int r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Elements<TableCell>().ToList();
            if (colIndex >= cells.Count) continue;

            var cell = cells[colIndex];

            // Ensure TableCellProperties exists
            var tcp = cell.GetFirstChild<TableCellProperties>();
            if (tcp == null)
            {
                tcp = new TableCellProperties();
                cell.PrependChild(tcp);
            }

            // Remove any existing VerticalMerge
            foreach (var vm in tcp.Elements<VerticalMerge>().ToList())
                vm.Remove();

            if (r == 0)
            {
                // First row: restart the merge span, content is kept
                tcp.AppendChild(new VerticalMerge { Val = MergedCellValues.Restart });

                // Vertically center the content in the merged cell
                var vAlign = tcp.GetFirstChild<TableCellVerticalAlignment>();
                if (vAlign == null)
                    tcp.AppendChild(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
                else
                    vAlign.Val = TableVerticalAlignmentValues.Center;
            }
            else
            {
                // Subsequent rows: continue the merge (content is suppressed by Word)
                tcp.AppendChild(new VerticalMerge());

                // Clear paragraph content so Word doesn't show stale text in merged cells
                foreach (var para in cell.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>().ToList())
                        run.Remove();
                }
            }
        }
    }

        private static void ReplaceLiteralHeaders(Table table, int num1, int num2, bool hasS2)
    {
        foreach (var textEl in table.Descendants<Text>())
        {
            var text = textEl.Text;
            if (string.IsNullOrEmpty(text)) continue;

            
            if (text.Contains("{{")) continue;

            bool modified = false;

            if (text.Contains("SS-1", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("SS-1", $"SS-{num1}", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("SS-2", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("SS-2", hasS2 ? $"SS-{num2}" : "—", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("SS - 1", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("SS - 1", $"SS - {num1}", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("SS - 2", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("SS - 2", hasS2 ? $"SS - {num2}" : "—", StringComparison.OrdinalIgnoreCase); modified = true; }

            if (text.Contains("1.Serim", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("1.Serim", $"{num1}.Serim", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("2.Serim", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("2.Serim", hasS2 ? $"{num2}.Serim" : "—", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("1. Serim", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("1. Serim", $"{num1}. Serim", StringComparison.OrdinalIgnoreCase); modified = true; }
            if (text.Contains("2. Serim", StringComparison.OrdinalIgnoreCase)) { text = text.Replace("2. Serim", hasS2 ? $"{num2}. Serim" : "—", StringComparison.OrdinalIgnoreCase); modified = true; }

            if (modified)
            {
                textEl.Text = text;
            }
        }
    }

    /// <summary>
    /// When there is an odd number of seismic series (e.g. 1 or 3 series),
    /// the last table clone for the odd series (where s2 == null) is trimmed so that
    /// the right-hand columns (allocated for the non-existent even series) are deleted.
    /// </summary>
    private static void TrimSeismicTableForSingleSeries(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        var grid = table.GetFirstChild<TableGrid>();
        var gridCols = grid?.Elements<GridColumn>().ToList() ?? new List<GridColumn>();

        // Check whether this is a 5-column modulus table (or similar) or a 4-column coordinate table
        int maxCells = rows.Max(r => r.Elements<TableCell>().Count());

        if (gridCols.Count == 5 || maxCells >= 5)
        {
            // Modulus table (Poisson, Shear, Bulk, Elasticity):
            // Row 0 has 3 cells: C0 (label/empty), C1 (span 2 for Series 1), C2 (span 2 for Series 2).
            // Rows 1-3 have 5 cells: C0 (label), C1-C2 (Tabaka 1-2 for Series 1), C3-C4 (Tabaka 1-2 for Series 2).
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count >= 5)
                {
                    cells[cells.Count - 1].Remove();
                    cells[cells.Count - 2].Remove();
                }
                else if (cells.Count == 3)
                {
                    // Header row: delete C2 (Series 2)
                    cells[cells.Count - 1].Remove();
                }
                else if (cells.Count == 1)
                {
                    // Full row span if any
                    var gs = cells[0].GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>();
                    if (gs?.Val != null && gs.Val.Value > 2)
                    {
                        gs.Val = gs.Val.Value - 2;
                    }
                }
            }

            // Remove last 2 grid columns
            int removedColsWidth = 0;
            if (gridCols.Count >= 2)
            {
                for (int i = gridCols.Count - 1; i >= gridCols.Count - 2; i--)
                {
                    if (int.TryParse(gridCols[i].Width?.Value, out int cw))
                        removedColsWidth += cw;
                    gridCols[i].Remove();
                }
            }

            // Adjust table width if specified in dxa
            var tblW = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableWidth>();
            if (tblW?.Type?.Value == TableWidthUnitValues.Dxa && int.TryParse(tblW.Width?.Value, out int tw))
            {
                if (tw > removedColsWidth && removedColsWidth > 0)
                {
                    tblW.Width = (tw - removedColsWidth).ToString();
                }
            }
        }
        else if (gridCols.Count == 4 || maxCells == 4)
        {
            // Coordinate table (Table 6):
            // Row 0 has 3 cells: C0 (span 2), C1 (SS-1), C2 (SS-2).
            // Rows 1-2 have 4 cells: C0 (label), C1 (sub-label), C2 (SS-1), C3 (SS-2).
            // Rows 3-5 have 3 cells: C0 (span 2), C1 (SS-1), C2 (SS-2).
            // In all rows, the last cell belongs to Series 2.
            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > 0)
                {
                    cells[cells.Count - 1].Remove();
                }
            }

            // Remove last grid column
            int removedColsWidth = 0;
            if (gridCols.Count >= 1)
            {
                var lastCol = gridCols.Last();
                if (int.TryParse(lastCol.Width?.Value, out int cw))
                    removedColsWidth += cw;
                lastCol.Remove();
            }

            // Adjust table width if specified in dxa
            var tblW = table.GetFirstChild<TableProperties>()?.GetFirstChild<TableWidth>();
            if (tblW?.Type?.Value == TableWidthUnitValues.Dxa && int.TryParse(tblW.Width?.Value, out int tw))
            {
                if (tw > removedColsWidth && removedColsWidth > 0)
                {
                    tblW.Width = (tw - removedColsWidth).ToString();
                }
            }
        }
    }

    private static Dictionary<string, string> BuildSeismicReplacements(
        Vs30SeriesItem? s1, int num1,
        Vs30SeriesItem? s2, int num2)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        d["SS_1_NAME"] = $"SS-{num1}";
        d["SS_2_NAME"] = s2 != null ? $"SS-{num2}" : "—";
        d["SERIM_1_NAME"] = $"{num1}. Serim";
        d["SERIM_2_NAME"] = s2 != null ? $"{num2}. Serim" : "—";
        d["SERIM_1_NO"] = num1.ToString();
        d["SERIM_2_NO"] = s2 != null ? num2.ToString() : "—";

        // Aliases with Turkish İ and alternate formats
        d["SERİM_1_NAME"] = $"{num1}. Serim";
        d["SERİM_2_NAME"] = s2 != null ? $"{num2}. Serim" : "—";
        d["SERİM_1_NO"] = num1.ToString();
        d["SERİM_2_NO"] = s2 != null ? num2.ToString() : "—";
        d["SERIM_1"] = $"{num1}. Serim";
        d["SERIM_2"] = s2 != null ? $"{num2}. Serim" : "—";
        d["SERİM_1"] = $"{num1}. Serim";
        d["SERİM_2"] = s2 != null ? $"{num2}. Serim" : "—";

        
        string xStart1  = s1?.XStart  ?? "";
        string yStart1  = s1?.YStart  ?? "";
        string xFinish1 = s1?.XFinish ?? "";
        string yFinish1 = s1?.YFinish ?? "";
        string p1_1     = s1?.PoissonLayer1 ?? "";
        string p2_1     = s1?.PoissonLayer2 ?? "";
        string sh1_1    = s1?.ShearModulusLayer1 ?? "";
        string sh2_1    = s1?.ShearModulusLayer2 ?? "";
        string b1_1     = s1?.BulkModulusLayer1 ?? "";
        string b2_1     = s1?.BulkModulusLayer2 ?? "";
        string e1_1     = s1?.ElasticityModulusLayer1 ?? "";
        string e2_1     = s1?.ElasticityModulusLayer2 ?? "";

        string p1_1_eval  = SeismicEvaluationHelper.EvaluatePoisson(p1_1);
        string p2_1_eval  = SeismicEvaluationHelper.EvaluatePoisson(p2_1);
        string sh1_1_eval = SeismicEvaluationHelper.EvaluateShearModulus(sh1_1);
        string sh2_1_eval = SeismicEvaluationHelper.EvaluateShearModulus(sh2_1);
        string b1_1_eval  = SeismicEvaluationHelper.EvaluateBulkModulus(b1_1);
        string b2_1_eval  = SeismicEvaluationHelper.EvaluateBulkModulus(b2_1);
        string e1_1_eval  = SeismicEvaluationHelper.EvaluateElasticityModulus(e1_1);
        string e2_1_eval  = SeismicEvaluationHelper.EvaluateElasticityModulus(e2_1);

        
        d["SS_1_XSTART"]      = xStart1;
        d["SS_1_YSTART"]      = yStart1;
        d["SS_1_XFINISH"]     = xFinish1;
        d["SS_1_YFINISH"]     = yFinish1;
        d["SS_1_POISSON1"]    = p1_1;
        d["SS_1_POISSON2"]    = p2_1;
        d["SS_1_SHEAR1"]      = sh1_1;
        d["SS_1_SHEAR2"]      = sh2_1;
        d["SS_1_BULK1"]       = b1_1;
        d["SS_1_BULK2"]       = b2_1;
        d["SS_1_ELASTICITY1"] = e1_1;
        d["SS_1_ELASTICITY2"] = e2_1;

        // Değerlendirmeler (SS_1)
        d["SS_1_POISSON1_EVAL"]    = p1_1_eval;
        d["SS_1_POISSON2_EVAL"]    = p2_1_eval;
        d["SS_1_SHEAR1_EVAL"]      = sh1_1_eval;
        d["SS_1_SHEAR2_EVAL"]      = sh2_1_eval;
        d["SS_1_BULK1_EVAL"]       = b1_1_eval;
        d["SS_1_BULK2_EVAL"]       = b2_1_eval;
        d["SS_1_ELASTICITY1_EVAL"] = e1_1_eval;
        d["SS_1_ELASTICITY2_EVAL"] = e2_1_eval;

        // Alternatif aliaslar (SS_1)
        d["SS_1_POISSON1_DESC"] = p1_1_eval;
        d["SS_1_POISSON2_DESC"] = p2_1_eval;
        d["SS_1_POISSON1_DEG"]  = p1_1_eval;
        d["SS_1_POISSON2_DEG"]  = p2_1_eval;
        d["SS_1_POISSON1_DEGERLENDIRME"] = p1_1_eval;
        d["SS_1_POISSON2_DEGERLENDIRME"] = p2_1_eval;
        d["SS_1_POISSON1_DEĞERLENDİRME"] = p1_1_eval;
        d["SS_1_POISSON2_DEĞERLENDİRME"] = p2_1_eval;

        d["SS_1_SHEAR1_DESC"] = sh1_1_eval;
        d["SS_1_SHEAR2_DESC"] = sh2_1_eval;
        d["SS_1_SHEAR1_DEG"]  = sh1_1_eval;
        d["SS_1_SHEAR2_DEG"]  = sh2_1_eval;
        d["SS_1_SHEAR1_DEGERLENDIRME"] = sh1_1_eval;
        d["SS_1_SHEAR2_DEGERLENDIRME"] = sh2_1_eval;
        d["SS_1_SHEAR1_DEĞERLENDİRME"] = sh1_1_eval;
        d["SS_1_SHEAR2_DEĞERLENDİRME"] = sh2_1_eval;

        d["SS_1_BULK1_DESC"]  = b1_1_eval;
        d["SS_1_BULK2_DESC"]  = b2_1_eval;
        d["SS_1_BULK1_DEG"]   = b1_1_eval;
        d["SS_1_BULK2_DEG"]   = b2_1_eval;
        d["SS_1_BULK1_SINIF"] = b1_1_eval;
        d["SS_1_BULK2_SINIF"] = b2_1_eval;
        d["SS_1_BULK1_CLASS"] = b1_1_eval;
        d["SS_1_BULK2_CLASS"] = b2_1_eval;
        d["SS_1_BULK1_DEGERLENDIRME"] = b1_1_eval;
        d["SS_1_BULK2_DEGERLENDIRME"] = b2_1_eval;
        d["SS_1_BULK1_DEĞERLENDİRME"] = b1_1_eval;
        d["SS_1_BULK2_DEĞERLENDİRME"] = b2_1_eval;

        d["SS_1_ELASTICITY1_DESC"] = e1_1_eval;
        d["SS_1_ELASTICITY2_DESC"] = e2_1_eval;
        d["SS_1_ELASTICITY1_DEG"]  = e1_1_eval;
        d["SS_1_ELASTICITY2_DEG"]  = e2_1_eval;
        d["SS_1_ELASTICITY1_DEGERLENDIRME"] = e1_1_eval;
        d["SS_1_ELASTICITY2_DEGERLENDIRME"] = e2_1_eval;
        d["SS_1_ELASTICITY1_DEĞERLENDİRME"] = e1_1_eval;
        d["SS_1_ELASTICITY2_DEĞERLENDİRME"] = e2_1_eval;

        
        d["SEİSMİCCOORDİNATESS-1XSTART"]  = xStart1;
        d["SEİSMİCCOORDİNATESS-1YSTART"]  = yStart1;
        d["SEİSMİCCOORDİNATESS-1XFİNİSH"] = xFinish1;
        d["SEİSMİCCOORDİNATESS-1YFİNİSH"] = yFinish1;
        d["POİSSONRATİO1S1"]              = p1_1;
        d["POİSSONRATİO2S1"]              = p2_1;
        d["SHEARMODULUS1S1"]              = sh1_1;
        d["SHEARMODULUS2S1"]              = sh2_1;
        d["BULKMODULE1S1"]                = b1_1;
        d["BULKMODULE2S1"]                = b2_1;
        d["ELASTİCİTYMODULUS1S1"]          = e1_1;
        d["ELASTİCİTYMODULUS2S1"]          = e2_1;
        d["POİSSONRATİO1S1_EVAL"]         = p1_1_eval;
        d["POİSSONRATİO2S1_EVAL"]         = p2_1_eval;
        d["SHEARMODULUS1S1_EVAL"]         = sh1_1_eval;
        d["SHEARMODULUS2S1_EVAL"]         = sh2_1_eval;
        d["BULKMODULE1S1_EVAL"]           = b1_1_eval;
        d["BULKMODULE2S1_EVAL"]           = b2_1_eval;
        d["ELASTİCİTYMODULUS1S1_EVAL"]     = e1_1_eval;
        d["ELASTİCİTYMODULUS2S1_EVAL"]     = e2_1_eval;

        
        string xStart2  = s2?.XStart  ?? "";
        string yStart2  = s2?.YStart  ?? "";
        string xFinish2 = s2?.XFinish ?? "";
        string yFinish2 = s2?.YFinish ?? "";
        string p1_2     = s2?.PoissonLayer1 ?? "";
        string p2_2     = s2?.PoissonLayer2 ?? "";
        string sh1_2    = s2?.ShearModulusLayer1 ?? "";
        string sh2_2    = s2?.ShearModulusLayer2 ?? "";
        string b1_2     = s2?.BulkModulusLayer1 ?? "";
        string b2_2     = s2?.BulkModulusLayer2 ?? "";
        string e1_2     = s2?.ElasticityModulusLayer1 ?? "";
        string e2_2     = s2?.ElasticityModulusLayer2 ?? "";

        string p1_2_eval  = SeismicEvaluationHelper.EvaluatePoisson(p1_2);
        string p2_2_eval  = SeismicEvaluationHelper.EvaluatePoisson(p2_2);
        string sh1_2_eval = SeismicEvaluationHelper.EvaluateShearModulus(sh1_2);
        string sh2_2_eval = SeismicEvaluationHelper.EvaluateShearModulus(sh2_2);
        string b1_2_eval  = SeismicEvaluationHelper.EvaluateBulkModulus(b1_2);
        string b2_2_eval  = SeismicEvaluationHelper.EvaluateBulkModulus(b2_2);
        string e1_2_eval  = SeismicEvaluationHelper.EvaluateElasticityModulus(e1_2);
        string e2_2_eval  = SeismicEvaluationHelper.EvaluateElasticityModulus(e2_2);

        
        d["SS_2_XSTART"]      = xStart2;
        d["SS_2_YSTART"]      = yStart2;
        d["SS_2_XFINISH"]     = xFinish2;
        d["SS_2_YFINISH"]     = yFinish2;
        d["SS_2_POISSON1"]    = p1_2;
        d["SS_2_POISSON2"]    = p2_2;
        d["SS_2_SHEAR1"]      = sh1_2;
        d["SS_2_SHEAR2"]      = sh2_2;
        d["SS_2_BULK1"]       = b1_2;
        d["SS_2_BULK2"]       = b2_2;
        d["SS_2_ELASTICITY1"] = e1_2;
        d["SS_2_ELASTICITY2"] = e2_2;

        // Değerlendirmeler (SS_2)
        d["SS_2_POISSON1_EVAL"]    = p1_2_eval;
        d["SS_2_POISSON2_EVAL"]    = p2_2_eval;
        d["SS_2_SHEAR1_EVAL"]      = sh1_2_eval;
        d["SS_2_SHEAR2_EVAL"]      = sh2_2_eval;
        d["SS_2_BULK1_EVAL"]       = b1_2_eval;
        d["SS_2_BULK2_EVAL"]       = b2_2_eval;
        d["SS_2_ELASTICITY1_EVAL"] = e1_2_eval;
        d["SS_2_ELASTICITY2_EVAL"] = e2_2_eval;

        // Alternatif aliaslar (SS_2)
        d["SS_2_POISSON1_DESC"] = p1_2_eval;
        d["SS_2_POISSON2_DESC"] = p2_2_eval;
        d["SS_2_POISSON1_DEG"]  = p1_2_eval;
        d["SS_2_POISSON2_DEG"]  = p2_2_eval;
        d["SS_2_POISSON1_DEGERLENDIRME"] = p1_2_eval;
        d["SS_2_POISSON2_DEGERLENDIRME"] = p2_2_eval;
        d["SS_2_POISSON1_DEĞERLENDİRME"] = p1_2_eval;
        d["SS_2_POISSON2_DEĞERLENDİRME"] = p2_2_eval;

        d["SS_2_SHEAR1_DESC"] = sh1_2_eval;
        d["SS_2_SHEAR2_DESC"] = sh2_2_eval;
        d["SS_2_SHEAR1_DEG"]  = sh1_2_eval;
        d["SS_2_SHEAR2_DEG"]  = sh2_2_eval;
        d["SS_2_SHEAR1_DEGERLENDIRME"] = sh1_2_eval;
        d["SS_2_SHEAR2_DEGERLENDIRME"] = sh2_2_eval;
        d["SS_2_SHEAR1_DEĞERLENDİRME"] = sh1_2_eval;
        d["SS_2_SHEAR2_DEĞERLENDİRME"] = sh2_2_eval;

        d["SS_2_BULK1_DESC"]  = b1_2_eval;
        d["SS_2_BULK2_DESC"]  = b2_2_eval;
        d["SS_2_BULK1_DEG"]   = b1_2_eval;
        d["SS_2_BULK2_DEG"]   = b2_2_eval;
        d["SS_2_BULK1_SINIF"] = b1_2_eval;
        d["SS_2_BULK2_SINIF"] = b2_2_eval;
        d["SS_2_BULK1_CLASS"] = b1_2_eval;
        d["SS_2_BULK2_CLASS"] = b2_2_eval;
        d["SS_2_BULK1_DEGERLENDIRME"] = b1_2_eval;
        d["SS_2_BULK2_DEGERLENDIRME"] = b2_2_eval;
        d["SS_2_BULK1_DEĞERLENDİRME"] = b1_2_eval;
        d["SS_2_BULK2_DEĞERLENDİRME"] = b2_2_eval;

        d["SS_2_ELASTICITY1_DESC"] = e1_2_eval;
        d["SS_2_ELASTICITY2_DESC"] = e2_2_eval;
        d["SS_2_ELASTICITY1_DEG"]  = e1_2_eval;
        d["SS_2_ELASTICITY2_DEG"]  = e2_2_eval;
        d["SS_2_ELASTICITY1_DEGERLENDIRME"] = e1_2_eval;
        d["SS_2_ELASTICITY2_DEGERLENDIRME"] = e2_2_eval;
        d["SS_2_ELASTICITY1_DEĞERLENDİRME"] = e1_2_eval;
        d["SS_2_ELASTICITY2_DEĞERLENDİRME"] = e2_2_eval;

        
        d["SEİSMİCCOORDİNATESS-2XSTART"]  = xStart2;
        d["SEİSMİCCOORDİNATESS-2YSTART"]  = yStart2;
        d["SEİSMİCCOORDİNATESS-2XFİNİSH"] = xFinish2;
        d["SEİSMİCCOORDİNATESS-2YFİNİSH"] = yFinish2;
        d["POİSSONRATİO1S2"]              = p1_2;
        d["POİSSONRATİO2S2"]              = p2_2;
        d["SHEARMODULUS1S2"]              = sh1_2;
        d["SHEARMODULUS2S2"]              = sh2_2;
        d["BULKMODULE1S2"]                = b1_2;
        d["BULKMODULE2S2"]                = b2_2;
        d["ELASTİCİTYMODULUS1S2"]          = e1_2;
        d["ELASTİCİTYMODULUS2S2"]          = e2_2;
        d["POİSSONRATİO1S2_EVAL"]         = p1_2_eval;
        d["POİSSONRATİO2S2_EVAL"]         = p2_2_eval;
        d["SHEARMODULUS1S2_EVAL"]         = sh1_2_eval;
        d["SHEARMODULUS2S2_EVAL"]         = sh2_2_eval;
        d["BULKMODULE1S2_EVAL"]           = b1_2_eval;
        d["BULKMODULE2S2_EVAL"]           = b2_2_eval;
        d["ELASTİCİTYMODULUS1S2_EVAL"]     = e1_2_eval;
        d["ELASTİCİTYMODULUS2S2_EVAL"]     = e2_2_eval;

        
        d[$"SS_{num1}_NAME"]        = $"SS-{num1}";
        d[$"SS_{num1}_XSTART"]      = xStart1;
        d[$"SS_{num1}_YSTART"]      = yStart1;
        d[$"SS_{num1}_XFINISH"]     = xFinish1;
        d[$"SS_{num1}_YFINISH"]     = yFinish1;
        d[$"SS_{num1}_POISSON1"]    = p1_1;
        d[$"SS_{num1}_POISSON2"]    = p2_1;
        d[$"SS_{num1}_SHEAR1"]      = sh1_1;
        d[$"SS_{num1}_SHEAR2"]      = sh2_1;
        d[$"SS_{num1}_BULK1"]       = b1_1;
        d[$"SS_{num1}_BULK2"]       = b2_1;
        d[$"SS_{num1}_ELASTICITY1"] = e1_1;
        d[$"SS_{num1}_ELASTICITY2"] = e2_1;

        d[$"SS_{num1}_POISSON1_EVAL"]    = p1_1_eval;
        d[$"SS_{num1}_POISSON2_EVAL"]    = p2_1_eval;
        d[$"SS_{num1}_SHEAR1_EVAL"]      = sh1_1_eval;
        d[$"SS_{num1}_SHEAR2_EVAL"]      = sh2_1_eval;
        d[$"SS_{num1}_BULK1_EVAL"]       = b1_1_eval;
        d[$"SS_{num1}_BULK2_EVAL"]       = b2_1_eval;
        d[$"SS_{num1}_ELASTICITY1_EVAL"] = e1_1_eval;
        d[$"SS_{num1}_ELASTICITY2_EVAL"] = e2_1_eval;

        d[$"SS_{num1}_POISSON1_DESC"] = p1_1_eval;
        d[$"SS_{num1}_POISSON2_DESC"] = p2_1_eval;
        d[$"SS_{num1}_POISSON1_DEG"]  = p1_1_eval;
        d[$"SS_{num1}_POISSON2_DEG"]  = p2_1_eval;
        d[$"SS_{num1}_SHEAR1_DESC"] = sh1_1_eval;
        d[$"SS_{num1}_SHEAR2_DESC"] = sh2_1_eval;
        d[$"SS_{num1}_SHEAR1_DEG"]  = sh1_1_eval;
        d[$"SS_{num1}_SHEAR2_DEG"]  = sh2_1_eval;
        d[$"SS_{num1}_BULK1_DESC"]  = b1_1_eval;
        d[$"SS_{num1}_BULK2_DESC"]  = b2_1_eval;
        d[$"SS_{num1}_BULK1_DEG"]   = b1_1_eval;
        d[$"SS_{num1}_BULK2_DEG"]   = b2_1_eval;
        d[$"SS_{num1}_BULK1_SINIF"] = b1_1_eval;
        d[$"SS_{num1}_BULK2_SINIF"] = b2_1_eval;
        d[$"SS_{num1}_BULK1_CLASS"] = b1_1_eval;
        d[$"SS_{num1}_BULK2_CLASS"] = b2_1_eval;
        d[$"SS_{num1}_ELASTICITY1_DESC"] = e1_1_eval;
        d[$"SS_{num1}_ELASTICITY2_DESC"] = e2_1_eval;
        d[$"SS_{num1}_ELASTICITY1_DEG"]  = e1_1_eval;
        d[$"SS_{num1}_ELASTICITY2_DEG"]  = e2_1_eval;

        if (s2 != null)
        {
            d[$"SS_{num2}_NAME"]        = $"SS-{num2}";
            d[$"SS_{num2}_XSTART"]      = xStart2;
            d[$"SS_{num2}_YSTART"]      = yStart2;
            d[$"SS_{num2}_XFINISH"]     = xFinish2;
            d[$"SS_{num2}_YFINISH"]     = yFinish2;
            d[$"SS_{num2}_POISSON1"]    = p1_2;
            d[$"SS_{num2}_POISSON2"]    = p2_2;
            d[$"SS_{num2}_SHEAR1"]      = sh1_2;
            d[$"SS_{num2}_SHEAR2"]      = sh2_2;
            d[$"SS_{num2}_BULK1"]       = b1_2;
            d[$"SS_{num2}_BULK2"]       = b2_2;
            d[$"SS_{num2}_ELASTICITY1"] = e1_2;
            d[$"SS_{num2}_ELASTICITY2"] = e2_2;

            d[$"SS_{num2}_POISSON1_EVAL"]    = p1_2_eval;
            d[$"SS_{num2}_POISSON2_EVAL"]    = p2_2_eval;
            d[$"SS_{num2}_SHEAR1_EVAL"]      = sh1_2_eval;
            d[$"SS_{num2}_SHEAR2_EVAL"]      = sh2_2_eval;
            d[$"SS_{num2}_BULK1_EVAL"]       = b1_2_eval;
            d[$"SS_{num2}_BULK2_EVAL"]       = b2_2_eval;
            d[$"SS_{num2}_ELASTICITY1_EVAL"] = e1_2_eval;
            d[$"SS_{num2}_ELASTICITY2_EVAL"] = e2_2_eval;

            d[$"SS_{num2}_POISSON1_DESC"] = p1_2_eval;
            d[$"SS_{num2}_POISSON2_DESC"] = p2_2_eval;
            d[$"SS_{num2}_POISSON1_DEG"]  = p1_2_eval;
            d[$"SS_{num2}_POISSON2_DEG"]  = p2_2_eval;
            d[$"SS_{num2}_SHEAR1_DESC"] = sh1_2_eval;
            d[$"SS_{num2}_SHEAR2_DESC"] = sh2_2_eval;
            d[$"SS_{num2}_SHEAR1_DEG"]  = sh1_2_eval;
            d[$"SS_{num2}_SHEAR2_DEG"]  = sh2_2_eval;
            d[$"SS_{num2}_BULK1_DESC"]  = b1_2_eval;
            d[$"SS_{num2}_BULK2_DESC"]  = b2_2_eval;
            d[$"SS_{num2}_BULK1_DEG"]   = b1_2_eval;
            d[$"SS_{num2}_BULK2_DEG"]   = b2_2_eval;
            d[$"SS_{num2}_BULK1_SINIF"] = b1_2_eval;
            d[$"SS_{num2}_BULK2_SINIF"] = b2_2_eval;
            d[$"SS_{num2}_BULK1_CLASS"] = b1_2_eval;
            d[$"SS_{num2}_BULK2_CLASS"] = b2_2_eval;
            d[$"SS_{num2}_ELASTICITY1_DESC"] = e1_2_eval;
            d[$"SS_{num2}_ELASTICITY2_DESC"] = e2_2_eval;
            d[$"SS_{num2}_ELASTICITY1_DEG"]  = e1_2_eval;
            d[$"SS_{num2}_ELASTICITY2_DEG"]  = e2_2_eval;
        }

        return d;
    }

    
    
    
    
    
    private int ProcessDynamicPressuremeterTables(
        Body body,
        IEnumerable<DrillingWellItem> wells,
        HashSet<string> replacedKeys)
    {
        var wellList = wells.Where(w => w.HasPressuremeterTest && w.PressuremeterRows.Count > 0).ToList();
        int totalReplacements = 0;

        var tables = body.Descendants<Table>().ToList();
        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToList();
            TableRow? templateRow = null;

            foreach (var row in rows)
            {
                var rowText = row.InnerText;
                if (rowText.Contains("{{SONDAJ_PM_") || rowText.Contains("{{SK_PM_"))
                {
                    templateRow = row;
                    break;
                }
            }

            if (templateRow == null) continue;

            
            if (wellList.Count == 0)
            {
                var emptyReplacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["SONDAJ_PM_NAME"] = "",
                    ["SONDAJ_PM_ZEMIN_KOTU"] = "",
                    ["SONDAJ_PM_DEPTH"] = "",
                    ["SONDAJ_PM_PL"] = "",
                    ["SONDAJ_PM_PL_NET"] = "",
                    ["SONDAJ_PM_EM"] = "",
                    ["SK_PM_NAME"] = "",
                    ["SK_PM_ZEMIN_KOTU"] = "",
                    ["SK_PM_DEPTH"] = "",
                    ["SK_PM_PL"] = "",
                    ["SK_PM_PL_NET"] = "",
                    ["SK_PM_EM"] = ""
                };

                
                foreach (Match match in PlaceholderRegex.Matches(templateRow.InnerText))
                {
                    string key = match.Groups[1].Value;
                    if (key.StartsWith("SONDAJ_PM_", StringComparison.OrdinalIgnoreCase) ||
                        key.StartsWith("SK_PM_", StringComparison.OrdinalIgnoreCase))
                    {
                        emptyReplacements[key] = "";
                    }
                }

                totalReplacements += ProcessElementContainers(templateRow, emptyReplacements, replacedKeys);
                continue;
            }

            
            foreach (var w in wellList)
            {
                bool isFirstRowOfWell = true;
                foreach (var pmRow in w.PressuremeterRows)
                {
                    var newRow = (TableRow)templateRow.CloneNode(true);
                    var pmReplacements = GetPressuremeterReplacementsDictionary(w, pmRow, isFirstRowOfWell);
                    totalReplacements += ProcessElementContainers(newRow, pmReplacements, replacedKeys);
                    table.InsertBefore(newRow, templateRow);
                    isFirstRowOfWell = false;
                }
            }

            
            templateRow.Remove();
        }

        return totalReplacements;
    }

    private Dictionary<string, string> GetPressuremeterReplacementsDictionary(
        DrillingWellItem w,
        PressuremeterRowItem row,
        bool isFirstRowOfWell)
    {
        string depth = row.Depth;
        string pl    = row.LimitPressure;
        string plNet = row.NetLimitPressure;
        string em    = row.ElasticityModulus;

        
        string wellName   = isFirstRowOfWell ? w.WellName : string.Empty;
        string groundKot  = isFirstRowOfWell ? w.GroundKot : string.Empty;

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["SONDAJ_PM_NAME"]       = wellName,
            ["SONDAJ_PM_NO"]         = wellName,
            ["SK_PM_NAME"]           = wellName,
            ["SK_PM_NO"]             = wellName,
            ["SONDAJ_PM_ZEMIN_KOTU"] = groundKot,
            ["SK_PM_ZEMIN_KOTU"]     = groundKot,
            ["SONDAJ_PM_DEPTH"]      = depth,
            ["SONDAJ_PM_DERINLIK"]   = depth,
            ["SK_PM_DEPTH"]          = depth,
            ["SK_PM_DERINLIK"]       = depth,
            ["SONDAJ_PM_PL"]         = pl,
            ["SK_PM_PL"]             = pl,
            ["SONDAJ_PM_PL_NET"]     = plNet,
            ["SK_PM_PL_NET"]         = plNet,
            ["SONDAJ_PM_EM"]         = em,
            ["SK_PM_EM"]             = em,
        };
    }

    private static readonly System.Globalization.CultureInfo TrCulture = new System.Globalization.CultureInfo("tr-TR");

    private Dictionary<string, string> GetSeriesReplacementsDictionary(Vs30SeriesItem s, Vs30LayerItem? layer = null, double layerH = 0, bool isFirstValidRow = false)
    {
        
        string surfaceVs30Raw = string.IsNullOrWhiteSpace(s.SurfaceVs30) ? s.Vs30Result : s.SurfaceVs30.Trim();
        string surfaceVs30WithMsn = surfaceVs30Raw;
        if (!string.IsNullOrWhiteSpace(surfaceVs30WithMsn))
        {
            if (surfaceVs30WithMsn.EndsWith("m/s", StringComparison.OrdinalIgnoreCase) && !surfaceVs30WithMsn.EndsWith("m/sn", StringComparison.OrdinalIgnoreCase))
                surfaceVs30WithMsn = surfaceVs30WithMsn[..^3] + "m/sn";
            else if (!surfaceVs30WithMsn.EndsWith("m/sn", StringComparison.OrdinalIgnoreCase))
                surfaceVs30WithMsn = $"{surfaceVs30WithMsn} m/sn";
        }

        string t0Formatted = !string.IsNullOrWhiteSpace(s.T0) ? s.T0 : "—";
        string seismicSummaryText = $"Sismik Serim-{s.SeriesNumber} için; yüzeyden Vs30= {surfaceVs30WithMsn}, Bu serimden alınan kayıtlardan yapılan hesaplamalar sonucu zemin hakim titreşim periyodunun değeri T0 = {t0Formatted} sn olarak hesaplanmıştır.";

        string aStr = s.AValue.ToString("F2", TrCulture);

        
        string vs30Digits = new string(s.Vs30Result.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
        string vs30CleanNum = double.TryParse(vs30Digits, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsedVs)
            ? Math.Round(parsedVs).ToString(TrCulture)
            : s.Vs30Result;

        string vs30WithMsn = s.Vs30Result;
        if (!string.IsNullOrWhiteSpace(vs30WithMsn))
        {
            if (vs30WithMsn.EndsWith("m/s", StringComparison.OrdinalIgnoreCase) && !vs30WithMsn.EndsWith("m/sn", StringComparison.OrdinalIgnoreCase))
                vs30WithMsn = vs30WithMsn[..^3] + "m/sn";
            else if (!vs30WithMsn.EndsWith("m/sn", StringComparison.OrdinalIgnoreCase))
                vs30WithMsn = $"{vs30WithMsn} m/sn";
        }

        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["SERIES_NO"] = s.SeriesNumber.ToString(),
            ["SERIM_NO"] = s.SeriesNumber.ToString(),
            ["SERIM_NUMARASI"] = s.SeriesNumber.ToString(),

            ["SERIES_NAME"] = s.SeriesName,
            ["SERIM_ADI"] = s.SeriesName,
            ["SERIM_ISMI"] = s.SeriesName,

            ["SERIES_KOTSTART"] = s.KotStart.ToString("F2", TrCulture),
            ["SERIM_KOT_BASLANGIC"] = s.KotStart.ToString("F2", TrCulture),
            ["SERIM_KOTSTART"] = s.KotStart.ToString("F2", TrCulture),

            ["SERIES_KOTEND"] = s.KotEnd.ToString("F2", TrCulture),
            ["SERIM_KOT_BITIS"] = s.KotEnd.ToString("F2", TrCulture),
            ["SERIM_KOTEND"] = s.KotEnd.ToString("F2", TrCulture),

            ["SERIES_KOTAVG"] = s.KotAverage.ToString("F2", TrCulture),
            ["SERIM_KOT_ORTALAMA"] = s.KotAverage.ToString("F2", TrCulture),
            ["SERIM_KOTAVG"] = s.KotAverage.ToString("F2", TrCulture),

            ["SERIES_DF"] = s.Df.ToString("F2", TrCulture),
            ["SERIM_DF"] = s.Df.ToString("F2", TrCulture),
            ["SERIES_DF_M"] = s.Df.ToString("F2", TrCulture),

            ["SERIES_H1"] = s.H1.ToString("F2", TrCulture),
            ["SERIM_H1"] = s.H1.ToString("F2", TrCulture),

            ["SERIES_FORMULA"] = s.Formula,
            ["SERIM_FORMUL"] = s.Formula,
            ["SERIES_FORMULA_FULL"] = s.FormulaFull,
            ["SERIM_FORMUL_TAM"] = s.FormulaFull,

            ["SERIES_RESULT"] = s.Vs30Result,
            ["SERIM_SONUC"] = s.Vs30Result,
            ["SERIES_VS30"] = s.Vs30Result,
            ["SERIM_VS30"] = s.Vs30Result,
            ["SERIES_VS30_NUM"] = vs30CleanNum,
            ["SERIES_VS30_VAL"] = vs30CleanNum,

            ["SERIES_A"] = aStr,
            ["SERIM_A"] = aStr,
            ["A"] = aStr,
            ["SERIES_A_FORMULA"] = $"68x{vs30CleanNum}^(-0,6) = {aStr}",
            ["SERIES_A_FULL"] = $"A{s.SeriesNumber} = 68x{vs30CleanNum}^(-0,6) = {aStr}",

            ["T0"] = s.T0,
            ["SERIES_T0"] = s.T0,
            ["SERIM_T0"] = s.T0,

            ["SERIES_SEISMIC_SUMMARY"] = seismicSummaryText,
            ["SERIES_T0_SUMMARY"] = seismicSummaryText,
            ["SERIM_T0_OZET"] = seismicSummaryText,
            ["SERIES_PERIOD_SUMMARY"] = seismicSummaryText
        };

        if (layer != null)
        {
            string startStr = layer.StartDepth.ToString("F2", TrCulture);
            string endStr = layer.EndDepth.ToString("F2", TrCulture);
            string vsStr = layer.Velocity.ToString("F0", TrCulture);
            string hStr = layerH.ToString("F2", TrCulture);

            dict["LAYER_NO"] = layer.LayerNumber.ToString();
            dict["TABAKA_NO"] = layer.LayerNumber.ToString();

            dict["LAYER_START"] = startStr;
            dict["TABAKA_BASLANGIC"] = startStr;
            dict["TABAKA_ALT"] = startStr;

            dict["LAYER_END"] = endStr;
            dict["TABAKA_BITIS"] = endStr;
            dict["TABAKA_UST"] = endStr;

            dict["LAYER_RANGE"] = $"{startStr} – {endStr}";
            dict["TABAKA_ARALIK"] = $"{startStr} – {endStr}";
            dict["TABAKA_DERINLIK"] = $"{startStr} – {endStr}";

            dict["LAYER_VS"] = vsStr;
            dict["TABAKA_VS"] = vsStr;
            dict["LAYER_VELOCITY"] = vsStr;
            dict["TABAKA_HIZ"] = vsStr;
            // LAYER_VSUR is an alias used in the second (Vs Ortalama) table
            dict["LAYER_VSUR"] = vsStr;

            
            dict["H"] = hStr;
            dict["LAYER_H"] = hStr;
            dict["TABAKA_H"] = hStr;

            
            if (!isFirstValidRow)
            {
                dict["SERIES_RESULT"] = string.Empty;
                dict["SERIM_SONUC"] = string.Empty;
                dict["SERIES_VS30"] = string.Empty;
                dict["SERIM_VS30"] = string.Empty;
                dict["SERIES_VS30_NUM"] = string.Empty;
                dict["SERIES_VS30_VAL"] = string.Empty;
            }
        }

        return dict;
    }

    private Dictionary<string, string> GetWellReplacementsDictionary(DrillingWellItem w, GroundwaterMeasurementItem? m = null, int index = 0)
    {
        var targetMeasurement = m ?? w.Measurements.FirstOrDefault();
        string measureDate = targetMeasurement != null ? targetMeasurement.MeasurementDate : "";
        string waterDepth = targetMeasurement != null ? targetMeasurement.WaterDepth : (w.IsCollapsed ? "Kuyu çökmüş" : "");

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["WELL_INDEX"] = index > 0 ? index.ToString() : "",
            ["SONDAJ_SIRASI"] = index > 0 ? index.ToString() : "",
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
            ["SONDAJ_SU_DERINLIK"] = waterDepth,

            ["DAY"] = targetMeasurement?.Day ?? "",
            ["GUN"] = targetMeasurement?.Day ?? "",
            ["WELL_DAY"] = targetMeasurement?.Day ?? "",
            ["OLCUM_GUN"] = targetMeasurement?.Day ?? ""
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

    /// <summary>
    /// Ensures the given RunProperties explicitly specify Times New Roman 12pt (24 half-points).
    /// This prevents Word from inheriting a different paragraph style when new runs are inserted.
    /// </summary>
    private static void EnsureFontSize12(RunProperties rpr)
    {
        // Font size: 24 half-points = 12pt
        var sz = rpr.GetFirstChild<FontSize>();
        if (sz == null)
        {
            sz = new FontSize();
            rpr.AppendChild(sz);
        }
        sz.Val = "24";

        var szCs = rpr.GetFirstChild<FontSizeComplexScript>();
        if (szCs == null)
        {
            szCs = new FontSizeComplexScript();
            rpr.AppendChild(szCs);
        }
        szCs.Val = "24";

        // Font family: Times New Roman
        var fonts = rpr.GetFirstChild<RunFonts>();
        if (fonts == null)
        {
            fonts = new RunFonts();
            rpr.AppendChild(fonts);
        }
        fonts.Ascii = "Times New Roman";
        fonts.HighAnsi = "Times New Roman";
        fonts.ComplexScript = "Times New Roman";
    }

    /// <summary>
    /// Normalizes placeholders that have been split across multiple Text (<w:t>) elements
    /// within runs in the same paragraph (common in Word documents edited by users).
    /// </summary>
    private static void NormalizeSplitPlaceholders(Paragraph paragraph)
    {
        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count <= 1) return;

        bool hasOpen = false;
        for (int i = 0; i < texts.Count; i++)
        {
            var t = texts[i].Text;
            if (t != null && t.Contains('{'))
            {
                hasOpen = true;
                break;
            }
        }
        if (!hasOpen) return;

        string fullText = string.Concat(texts.Select(t => t.Text ?? ""));
        if (!fullText.Contains("{{") || !fullText.Contains("}}")) return;

        var matches = PlaceholderRegex.Matches(fullText);
        if (matches.Count == 0) return;

        // Process matches in reverse order so character offsets before each match remain valid
        for (int m = matches.Count - 1; m >= 0; m--)
        {
            var match = matches[m];
            int mStart = match.Index;
            int mEnd = match.Index + match.Length;

            int currentOffset = 0;
            int firstIdx = -1;
            int lastIdx = -1;
            int firstStartOffset = 0;
            int lastStartOffset = 0;

            for (int i = 0; i < texts.Count; i++)
            {
                int textLen = texts[i].Text?.Length ?? 0;
                int textStart = currentOffset;
                int textEnd = currentOffset + textLen;

                if (firstIdx == -1 && mStart >= textStart && mStart < textEnd)
                {
                    firstIdx = i;
                    firstStartOffset = textStart;
                }

                if (lastIdx == -1 && (mEnd - 1) >= textStart && (mEnd - 1) < textEnd)
                {
                    lastIdx = i;
                    lastStartOffset = textStart;
                    break;
                }

                currentOffset += textLen;
            }

            // Only normalize if the placeholder is actually split across different Text nodes
            if (firstIdx != -1 && lastIdx != -1 && firstIdx < lastIdx)
            {
                string firstOriginal = texts[firstIdx].Text ?? "";
                string lastOriginal = texts[lastIdx].Text ?? "";

                int prefixLen = mStart - firstStartOffset;
                string prefix = prefixLen > 0 ? firstOriginal.Substring(0, prefixLen) : "";

                int suffixStart = mEnd - lastStartOffset;
                string suffix = suffixStart < lastOriginal.Length ? lastOriginal.Substring(suffixStart) : "";

                texts[firstIdx].Text = prefix + match.Value;

                for (int k = firstIdx + 1; k < lastIdx; k++)
                {
                    texts[k].Text = "";
                }

                texts[lastIdx].Text = suffix;
            }
        }
    }

    private int ReplaceInParagraph(
        Paragraph paragraph,
        IDictionary<string, string> variableReplacements,
        HashSet<string> replacedKeys)
    {
        int totalReplacements = 0;

        NormalizeSplitPlaceholders(paragraph);

        foreach (var textNode in paragraph.Descendants<Text>().ToList())
        {
            var text = textNode.Text;
            if (string.IsNullOrEmpty(text) || !text.Contains("{{")) continue;

            var matches = PlaceholderRegex.Matches(text);
            if (matches.Count == 0) continue;

            string updatedText = text;
            foreach (Match match in matches)
            {
                var keyName = match.Groups[1].Value;
                if (TryGetReplacementValue(variableReplacements, keyName, out var replacementValue))
                {
                    updatedText = updatedText.Replace(match.Value, replacementValue);
                    replacedKeys.Add(keyName);
                    totalReplacements++;
                }
            }

            if (updatedText != text)
            {
                if (updatedText.Contains('\n'))
                {
                    var parts = updatedText.Split('\n');
                    textNode.Text = parts[0];

                    var currentRun = textNode.Parent as Run;
                    if (currentRun != null && currentRun.Parent != null)
                    {
                        var insertAnchor = currentRun.NextSibling();
                        var parentEl = currentRun.Parent;
                        var runProps = currentRun.GetFirstChild<RunProperties>()?.CloneNode(true) as RunProperties;

                        for (int p = 1; p < parts.Length; p++)
                        {
                            var brRun = new Run(new Break());
                            var brProps = (runProps?.CloneNode(true) as RunProperties) ?? new RunProperties();
                            EnsureFontSize12(brProps);
                            brRun.PrependChild(brProps);

                            var newTextRun = new Run(new Text(parts[p]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
                            var textProps = (runProps?.CloneNode(true) as RunProperties) ?? new RunProperties();
                            EnsureFontSize12(textProps);
                            newTextRun.PrependChild(textProps);

                            if (insertAnchor == null)
                            {
                                parentEl.AppendChild(brRun);
                                parentEl.AppendChild(newTextRun);
                            }
                            else
                            {
                                parentEl.InsertBefore(brRun, insertAnchor);
                                parentEl.InsertBefore(newTextRun, insertAnchor);
                            }
                        }
                    }
                }
                else
                {
                    textNode.Text = updatedText;
                }
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
    
    
    
    
    
    
    private int ProcessDynamicMicrotremorTables(
        Body body,
        IDictionary<string, string> variableReplacements,
        HashSet<string> replacedKeys)
    {
        
        int mtCount = 0;
        if (variableReplacements.TryGetValue("MICROTREMOR_COUNT", out var mcStr) ||
            variableReplacements.TryGetValue("MT_COUNT", out mcStr))
        {
            int.TryParse(mcStr, out mtCount);
        }
        if (mtCount <= 0) return 0;

        int totalReplacements = 0;
        var tables = body.Descendants<Table>().ToList();

        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToList();

            
            TableRow? templateRow = rows.FirstOrDefault(r =>
            {
                string text = r.InnerText;
                return text.Contains("{{MT_NO}}") ||
                       text.Contains("{{MT_T0}}") ||
                       text.Contains("{{MT_TA}}") ||
                       text.Contains("{{MT_TB}}") ||
                       text.Contains("{{MT_FREQ}}") ||
                       text.Contains("{{MT_DATE}}") ||
                       text.Contains("{{MT_TARIH}}") ||
                       text.Contains("{{MT_X}}") ||
                       text.Contains("{{MT_Y}}") ||
                       text.Contains("{{MT_YATAY}}") ||
                       text.Contains("{{MT_DUZEY}}");
            });

            if (templateRow == null) continue;

            
            for (int n = 1; n <= mtCount; n++)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                
                var mtReplacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["MT_NO"]     = $"MT-{n}",
                    ["MT_NUM"]    = n.ToString(),
                    ["MT_T0"]     = variableReplacements.TryGetValue($"MT_{n}_T0",    out var v1) ? v1 : "",
                    ["MT_TA"]     = variableReplacements.TryGetValue($"MT_{n}_TA",    out var v2) ? v2 : "",
                    ["MT_TB"]     = variableReplacements.TryGetValue($"MT_{n}_TB",    out var v3) ? v3 : "",
                    ["MT_FREQ"]   = variableReplacements.TryGetValue($"MT_{n}_FREQ",  out var v4) ? v4 : "",
                    ["MT_FREKANS"]= variableReplacements.TryGetValue($"MT_{n}_FREQ",  out var v4b) ? v4b : "",
                    ["MT_DATE"]   = variableReplacements.TryGetValue($"MT_{n}_DATE",  out var v5) ? v5 : "",
                    ["MT_TARIH"]  = variableReplacements.TryGetValue($"MT_{n}_DATE",  out var v5b) ? v5b : "",
                    ["MT_X"]      = variableReplacements.TryGetValue($"MT_{n}_X",     out var v6) ? v6 : "",
                    ["MT_YATAY"]  = variableReplacements.TryGetValue($"MT_{n}_X",     out var v6b) ? v6b : "",
                    ["MT_Y"]      = variableReplacements.TryGetValue($"MT_{n}_Y",     out var v7) ? v7 : "",
                    ["MT_DUZEY"]  = variableReplacements.TryGetValue($"MT_{n}_Y",     out var v7b) ? v7b : "",
                };

                totalReplacements += ProcessElementContainers(newRow, mtReplacements, replacedKeys);
                table.InsertBefore(newRow, templateRow);
            }

            
            templateRow.Remove();
        }

        return totalReplacements;
    }

    private sealed class FormationBlock
    {
        public Paragraph? TypeParagraph { get; set; }
        public Paragraph TextParagraph { get; set; } = null!;
        public List<Paragraph> AllParagraphs { get; set; } = new();
        public bool HasSeparateTypeParagraph { get; set; }
    }

    private static bool IsFormationTypePlaceholder(string text) =>
        text.IndexOf("FORMATİONTYPE", StringComparison.OrdinalIgnoreCase) >= 0 ||
        text.IndexOf("FORMAŞİONTYPE", StringComparison.OrdinalIgnoreCase) >= 0 ||
        text.IndexOf("FORMATIONTYPE", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool IsFormationTextPlaceholder(string text) =>
        text.IndexOf("FORMATİONTEXT", StringComparison.OrdinalIgnoreCase) >= 0 ||
        text.IndexOf("FORMAŞİONTEXT", StringComparison.OrdinalIgnoreCase) >= 0 ||
        text.IndexOf("FORMATIONTEXT", StringComparison.OrdinalIgnoreCase) >= 0;

    private static List<FormationDefinition> InferFormationsFromReplacements(IDictionary<string, string> replacements)
    {
        var list = new List<FormationDefinition>();
        string? val = null;
        if (replacements.TryGetValue("FORMATİONTYPE", out var v1) && !string.IsNullOrWhiteSpace(v1)) val = v1;
        else if (replacements.TryGetValue("FORMAŞİONTYPE", out var v2) && !string.IsNullOrWhiteSpace(v2)) val = v2;
        else if (replacements.TryGetValue("FORMATIONTYPE", out var v3) && !string.IsNullOrWhiteSpace(v3)) val = v3;

        if (string.IsNullOrWhiteSpace(val)) return list;

        var parts = val.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var trimmed = p.Trim();
            var repoMatch = FormationRepository.All
                .FirstOrDefault(f => string.Equals(f.Name.Trim(), trimmed, StringComparison.OrdinalIgnoreCase));
            if (repoMatch != null)
            {
                list.Add(repoMatch);
            }
            else
            {
                list.Add(new FormationDefinition(trimmed, ""));
            }
        }
        return list;
    }

    private int ProcessDynamicFormationParagraphs(
        Body body,
        IEnumerable<FormationDefinition> formations,
        HashSet<string> replacedKeys)
    {
        var formationList = formations.ToList();
        if (formationList.Count == 0) return 0;

        int totalReplacements = 0;

        var allParagraphs = body.Descendants<Paragraph>().ToList();

        var blocks = new List<FormationBlock>();
        var visited = new HashSet<Paragraph>();

        for (int i = 0; i < allParagraphs.Count; i++)
        {
            var p = allParagraphs[i];
            if (visited.Contains(p)) continue;

            string pText = p.InnerText ?? "";
            bool hasType = IsFormationTypePlaceholder(pText);
            bool hasText = IsFormationTextPlaceholder(pText);

            if (!hasType && !hasText) continue;

            // Pattern 1: Both placeholders in the SAME paragraph
            if (hasType && hasText)
            {
                blocks.Add(new FormationBlock
                {
                    TypeParagraph = p,
                    TextParagraph = p,
                    AllParagraphs = new List<Paragraph> { p },
                    HasSeparateTypeParagraph = false
                });
                visited.Add(p);
                continue;
            }

            // Pattern 2: FORMATİONTYPE followed shortly by FORMATİONTEXT in separate paragraphs
            if (hasType)
            {
                Paragraph? foundTextP = null;
                var candidateBlock = new List<Paragraph> { p };

                for (int j = i + 1; j < allParagraphs.Count && j <= i + 5; j++)
                {
                    var nextP = allParagraphs[j];
                    if (nextP.Parent != p.Parent) break;

                    string nextText = nextP.InnerText ?? "";
                    if (IsFormationTextPlaceholder(nextText))
                    {
                        foundTextP = nextP;
                        candidateBlock.Add(nextP);
                        break;
                    }
                    else if (string.IsNullOrWhiteSpace(nextText))
                    {
                        candidateBlock.Add(nextP);
                    }
                    else
                    {
                        break;
                    }
                }

                if (foundTextP != null)
                {
                    int nextIdx = allParagraphs.IndexOf(foundTextP) + 1;
                    if (nextIdx < allParagraphs.Count &&
                        allParagraphs[nextIdx].Parent == p.Parent &&
                        string.IsNullOrWhiteSpace(allParagraphs[nextIdx].InnerText))
                    {
                        candidateBlock.Add(allParagraphs[nextIdx]);
                    }

                    blocks.Add(new FormationBlock
                    {
                        TypeParagraph = p,
                        TextParagraph = foundTextP,
                        AllParagraphs = candidateBlock,
                        HasSeparateTypeParagraph = true
                    });

                    foreach (var cp in candidateBlock)
                    {
                        visited.Add(cp);
                    }
                    continue;
                }
                else
                {
                    // Isolated FORMATİONTYPE elsewhere in the report
                    continue;
                }
            }

            // Pattern 3: Isolated FORMATİONTEXT without any preceding FORMATİONTYPE
            if (hasText)
            {
                blocks.Add(new FormationBlock
                {
                    TypeParagraph = null,
                    TextParagraph = p,
                    AllParagraphs = new List<Paragraph> { p },
                    HasSeparateTypeParagraph = false
                });
                visited.Add(p);
                continue;
            }
        }

        foreach (var block in blocks)
        {
            if (block.AllParagraphs.Count == 0) continue;
            var anchor = block.AllParagraphs[0];
            var parent = anchor.Parent;
            if (parent == null) continue;

            for (int fIndex = 0; fIndex < formationList.Count; fIndex++)
            {
                var f = formationList[fIndex];
                string fName = f.Name;
                string fDesc = f.Description.Trim().Replace("\r\n", "\n").Replace("\n", " ");

                var formReplacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["FORMATİONTYPE"] = fName,
                    ["FORMAŞİONTYPE"] = fName,
                    ["FORMATIONTYPE"] = fName,
                    ["FORMATİONTEXT"] = block.HasSeparateTypeParagraph ? fDesc : $"{fName}\n{fDesc}",
                    ["FORMAŞİONTEXT"] = block.HasSeparateTypeParagraph ? fDesc : $"{fName}\n{fDesc}",
                    ["FORMATIONTEXT"] = block.HasSeparateTypeParagraph ? fDesc : $"{fName}\n{fDesc}",
                };

                foreach (var pInBlock in block.AllParagraphs)
                {
                    var newP = (Paragraph)pInBlock.CloneNode(true);
                    totalReplacements += ReplaceInParagraph(newP, formReplacements, replacedKeys);
                    parent.InsertBefore(newP, anchor);
                }
            }

            foreach (var pInBlock in block.AllParagraphs)
            {
                pInBlock.Remove();
            }
        }

        return totalReplacements;
    }
}
