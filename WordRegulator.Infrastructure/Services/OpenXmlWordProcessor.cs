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
        IEnumerable<Vs30SeriesItem>? vs30Series = null,
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

                var fileResult = await ProcessDocumentAsync(templateFile, targetOutputPath, variableReplacements, wells, vs30Series, cancellationToken);
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

                // Presiyometre satırları ({{SONDAJ_PM_}} / {{SK_PM_}}) bu metodun kapsamı dışında;
                // ayrı ProcessDynamicPressuremeterTables ile işlenir.
                bool isDynamicRow = (rowText.Contains("{{WELL_") || rowText.Contains("{{SONDAJ_") || rowText.Contains("{{SK_"))
                                 && !rowText.Contains("{{SONDAJ_PM_") && !rowText.Contains("{{SK_PM_");
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
            // Presiyometre paragrafları ({{SONDAJ_PM_}} / {{SK_PM_}}) bu metodun kapsamı dışında.
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

    private int ProcessDynamicVs30Tables(
        Body body,
        IEnumerable<Vs30SeriesItem> vs30Series,
        HashSet<string> replacedKeys)
    {
        var seriesList = vs30Series.ToList();
        if (seriesList.Count == 0) return 0;
        int totalReplacements = 0;

        // Body'nin doğrudan çocuklarını liste olarak al (tablolar ve paragraflar karışık)
        var bodyChildren = body.ChildElements.ToList();

        foreach (var bodyChild in bodyChildren)
        {
            if (bodyChild is not Table table) continue;

            var rows = table.Elements<TableRow>().ToList();
            if (rows.Count == 0) continue;

            // Bu tabloda {{H}} veya {{LAYER_}} içeren tabaka şablonu var mı?
            bool hasLayerTemplate = rows.Any(r =>
                r.InnerText.Contains("{{H}}") ||
                r.InnerText.Contains("{{LAYER_") ||
                r.InnerText.Contains("{{TABAKA_"));

            if (!hasLayerTemplate) continue;

            // Tablonun ardından gelen başlık paragrafını bul
            // Örn: "Tablo 18b: Masw{{SERIES_NO}} Vs ve Vs30 değerlendirme sonuç tablosu"
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

            // Her serim için tam tabloyu + başlık paragrafını klonla
            foreach (var s in seriesList)
            {
                // --- 1. Tabloyu klonla ---
                var clonedTable = (Table)table.CloneNode(true);
                var clonedRows = clonedTable.Elements<TableRow>().ToList();

                // Şablon tabaka satırını bul ({{H}} veya {{LAYER_}} içeren)
                var templateDataRow = clonedRows.FirstOrDefault(r =>
                    r.InnerText.Contains("{{H}}") ||
                    r.InnerText.Contains("{{LAYER_") ||
                    r.InnerText.Contains("{{TABAKA_"));

                if (templateDataRow != null)
                {
                    // Tablo tipi belirleme: {{H}} içeriyorsa h tablosu (df kontrolü yapılır),
                    // içermiyorsa LAYER_RANGE tablosu (tüm tabakalar her zaman gösterilir)
                    bool isHTable = templateDataRow.InnerText.Contains("{{H}}");

                    if (isHTable)
                    {
                        // H tablosu: df >= layer1.EndDepth ise 1. tabaka satır üretilmez
                        double sumThickness = 0;
                        bool firstValidInserted = false;

                        for (int j = 0; j < s.Layers.Count; j++)
                        {
                            var l = s.Layers[j];
                            double h_i;

                            if (j == 0)
                                h_i = Math.Max(0, l.EndDepth - s.Df);
                            else if (j < s.Layers.Count - 1)
                                h_i = Math.Max(0, l.EndDepth - s.Layers[j - 1].EndDepth);
                            else
                                h_i = Math.Max(0, 30.0 - sumThickness);

                            sumThickness += h_i;

                            // İlk tabaka hariç tutuluyorsa satır üretme
                            if (j == 0 && h_i <= 0) continue;

                            bool isFirstRow = !firstValidInserted;
                            firstValidInserted = true;

                            var newRow = (TableRow)templateDataRow.CloneNode(true);
                            var layerReplacements = GetSeriesReplacementsDictionary(s, l, h_i, isFirstRow);
                            totalReplacements += ProcessElementContainers(newRow, layerReplacements, replacedKeys);
                            clonedTable.InsertBefore(newRow, templateDataRow);
                        }
                    }
                    else
                    {
                        // LAYER_RANGE / LAYER_VS tablosu: df'den bağımsız, kullanıcının girdiği
                        // tüm tabakalar her zaman gösterilir (1. tabaka da dahil)
                        for (int j = 0; j < s.Layers.Count; j++)
                        {
                            var l = s.Layers[j];
                            var newRow = (TableRow)templateDataRow.CloneNode(true);
                            var layerReplacements = GetSeriesReplacementsDictionary(s, l, 0, false);
                            totalReplacements += ProcessElementContainers(newRow, layerReplacements, replacedKeys);
                            clonedTable.InsertBefore(newRow, templateDataRow);
                        }
                    }

                    // Şablon satırını sil (artık genişletildi)
                    templateDataRow.Remove();
                }

                // Serim düzeyi değişkenlerini klonlanan tablonun geri kalanına uygula (başlık satırı vs.)
                var seriesOnlyReplacements = GetSeriesReplacementsDictionary(s);
                foreach (var row in clonedTable.Elements<TableRow>())
                {
                    // Tabaka satırları zaten dolduruldu; sadece kalan satırları işle
                    totalReplacements += ProcessElementContainers(row, seriesOnlyReplacements, replacedKeys);
                }

                // Klonlanmış tabloyu orijinalin önüne ekle
                body.InsertBefore(clonedTable, table);
                totalReplacements++;

                // --- 2. Başlık paragrafını klonla ---
                if (titleParagraph != null)
                {
                    var clonedPara = (Paragraph)titleParagraph.CloneNode(true);
                    totalReplacements += ReplaceInParagraph(clonedPara, seriesOnlyReplacements, replacedKeys);
                    body.InsertBefore(clonedPara, table);
                }
            }

            // Orijinal şablon tablosunu kaldır
            table.Remove();

            // Orijinal başlık paragrafını kaldır
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

    // =====================================================================
    // Sismik Koordinat / Poisson / Kayma / Bulk / Elastisite Tabloları
    // Dinamik Çift-Serim Klonlama
    //
    // Şablon tabloda {{SS_1_XSTART}} ve {{SS_2_XSTART}} gibi işaretler varsa,
    // processor tabloyu serim ÇIFTLERI için klonlar:
    //   Serim 1-2 → orijinal tabloyu doldur
    //   Serim 3-4 → tablo klonla, 1→3, 2→4 diye değiştir
    //   Serim 5-6 → tablo klonla, 1→5, 2→6 diye değiştir vb.
    //
    // Aynı mekanizma şu prefix'ler için çalışır:
    //   SS_n_XSTART, SS_n_YSTART, SS_n_XFINISH, SS_n_YFINISH
    //   SS_n_POISSON1, SS_n_POISSON2
    //   SS_n_SHEAR1, SS_n_SHEAR2
    //   SS_n_BULK1, SS_n_BULK2
    //   SS_n_ELASTICITY1, SS_n_ELASTICITY2
    // =====================================================================
    private int ProcessDynamicSeismicTables(
        Body body,
        IEnumerable<Vs30SeriesItem> vs30Series,
        HashSet<string> replacedKeys)
    {
        var seriesList = vs30Series.ToList();
        if (seriesList.Count == 0) return 0;

        int totalReplacements = 0;

        // Yalnızca en içteki sismik tabloları seç (içinde başka tablo olmayan yaprak tablolar)
        // Böylece dış tablo / layout tabloları yanlışlıkla klonlanmaz
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

        // Serimleri 2'şerli grupla: [(1,2), (3,4), (5,6) ...]
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

                // Tabloyu OpenXML yapısını ve tüm hücre formatlarını birebir koruyarak klonla
                var cloned = (Table)table.CloneNode(true);

                // Floating (metin kaydırma / mutlak konumlandırma) özelliklerini kaldır:
                // Word'ün iki tabloyu üst üste (iç içe) bindirmesini kesin olarak engeller
                var tblPr = cloned.GetFirstChild<TableProperties>();
                if (tblPr != null)
                {
                    tblPr.GetFirstChild<TablePositionProperties>()?.Remove();
                }

                // 2. ve sonraki çiftlerde literal başlıktaki "SS-1" -> "SS-3", "1.Serim" -> "3.Serim" metinlerini güncelle
                if (pairIdx > 0)
                {
                    ReplaceLiteralHeaders(cloned, num1, num2, s2 != null);
                }

                // Bu çiftin (örn: 3 ve 4. serim) verilerini şablondaki SS_1_ ve SS_2_ etiketlerine eşle
                var pairReplacements = BuildSeismicReplacements(s1, num1, s2, num2);
                foreach (var row in cloned.Elements<TableRow>())
                {
                    totalReplacements += ProcessElementContainers(row, pairReplacements, replacedKeys);
                }

                // Klonlanan tabloyu orijinal tablonun hemen öncesine yerleştir
                parent.InsertBefore(cloned, table);

                // Tabloların Word'de birleşmemesi ve görsel olarak ayrılması için araya belirgin bir boş satır (paragraf) koy
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

            // Orijinal şablon tabloyu güvenle kaldır
            table.Remove();
        }

        return totalReplacements;
    }

    /// <summary>
    /// Klon tablolardaki sabit metin başlıklarını (SS-1, SS-2, 1.Serim, 2.Serim) serim numaralarına göre günceller.
    /// Hücre içi formatları ve {{...}} değişkenlerini bozmaz.
    /// </summary>
    private static void ReplaceLiteralHeaders(Table table, int num1, int num2, bool hasS2)
    {
        foreach (var textEl in table.Descendants<Text>())
        {
            var text = textEl.Text;
            if (string.IsNullOrEmpty(text)) continue;

            // Değişken etiketlerini elleme (onları ProcessElementContainers yapacak)
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
    /// Bir serim çifti (örneğin Serim 1-2 veya Serim 3-4) için replacements haritası üretir.
    /// Şablondaki SS_1_ / SS_2_ etiketlerini bu çiftin ilgili verileriyle eşler.
    /// </summary>
    private static Dictionary<string, string> BuildSeismicReplacements(
        Vs30SeriesItem? s1, int num1,
        Vs30SeriesItem? s2, int num2)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Başlık değişkenleri
        d["SS_1_NAME"] = $"SS-{num1}";
        d["SS_2_NAME"] = s2 != null ? $"SS-{num2}" : "—";
        d["SERIM_1_NAME"] = $"{num1}. Serim";
        d["SERIM_2_NAME"] = s2 != null ? $"{num2}. Serim" : "—";
        d["SERIM_1_NO"] = num1.ToString();
        d["SERIM_2_NO"] = s2 != null ? num2.ToString() : "—";

        // 1. Sütun Verileri (s1: Serim 1, 3, 5...)
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

        // Dinamik SS_1_ prefix'leri
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

        // Eski değişken isimleri için geriye dönük uyumluluk (1. sütun)
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

        // 2. Sütun Verileri (s2: Serim 2, 4, 6...)
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

        // Dinamik SS_2_ prefix'leri
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

        // Eski değişken isimleri için geriye dönük uyumluluk (2. sütun)
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

        // Şablonda doğrudan SS_3_ veya SS_4_ yazılmışsa onlar da çalışsın
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
        }

        return d;
    }

    // =====================================================================
    // Presiyometre Deneyi Tablosu Dinamik İşleme
    // Şablon tabloşunda {{SONDAJ_PM_}} prefix'li değişken varsa,
    // presiyometre verisi olan her kuyu için satırlar çoğaltılır.
    // =====================================================================
    private int ProcessDynamicPressuremeterTables(
        Body body,
        IEnumerable<DrillingWellItem> wells,
        HashSet<string> replacedKeys)
    {
        var wellList = wells.Where(w => w.HasPressuremeterTest && w.PressuremeterRows.Count > 0).ToList();
        if (wellList.Count == 0) return 0;

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

            // Şablon satır öncesine tüm presiyometre verisi olan kuyuların satırlarını ekle
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

            // Orijinal şablon satırını kaldır
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

        // Sondaj Adı ve Zemin Kotu sadece ilk satırda yazılır (tablo birleşim görünümü için)
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
        // Yüzeyde elde edilen VS30 (elle girilmiş); boşsa hesaplanan değeri fallback olarak kullan
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

        // Hesaplanan VS30 — sayısal kısmını temizle
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

            // {{H}} — tabaka yüksekliği (Görev 6)
            dict["H"] = hStr;
            dict["LAYER_H"] = hStr;
            dict["TABAKA_H"] = hStr;

            // SERIES_RESULT ve VS30 yalnızca ilk geçerli tabaka satırında dolu (Görev 6)
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
            if (updatedFullText.Contains('\n'))
            {
                // Çok satırlı replacement: \n'leri Word satır sonu (w:br) olarak ekliyoruz
                var parts = updatedFullText.Split('\n');

                // texts[0]'a ilk parçayı yaz, diğerlerini temizle
                texts[0].Text = parts[0];
                for (int i = 1; i < texts.Count; i++)
                    texts[i].Text = string.Empty;

                // Birinci Text'in içinde bulunduğu Run (veya ilk ataş)
                var firstRun = texts[0].Parent;
                if (firstRun != null)
                {
                    // Mevcut run'dan sonraki kardeşi referans olarak al (insert point)
                    var insertAnchor = firstRun.NextSibling();
                    var parentEl   = firstRun.Parent;

                    // Her ek parça için: <w:br/> run + metin run
                    for (int p = 1; p < parts.Length; p++)
                    {
                        // Satır sonu runı: <w:r><w:br/></w:r>
                        var brRun = new Run(new Break());

                        // İsteğe bağlı: mevcut run'un formatını (bold/font vb.) kopyala
                        var existingRpr = firstRun.Elements<RunProperties>().FirstOrDefault();
                        if (existingRpr != null)
                            brRun.PrependChild((RunProperties)existingRpr.CloneNode(true));

                        // Metin runı: <w:r><w:t xml:space="preserve">parça</w:t></w:r>
                        var textRun = new Run(
                            new Text(parts[p]) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
                        if (existingRpr != null)
                            textRun.PrependChild((RunProperties)existingRpr.CloneNode(true));

                        if (parentEl != null)
                        {
                            if (insertAnchor == null)
                            {
                                parentEl.AppendChild(brRun);
                                parentEl.AppendChild(textRun);
                            }
                            else
                            {
                                parentEl.InsertBefore(brRun, insertAnchor);
                                parentEl.InsertBefore(textRun, insertAnchor);
                            }
                        }
                    }
                }
            }
            else
            {
                texts[0].Text = updatedFullText;
                for (int i = 1; i < texts.Count; i++)
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
    // =====================================================================
    // Mikrotremör Ölçüm Tablosu Dinamik İşleme
    // Şablon tablosunda {{MT_NO}}, {{MT_T0}}, {{MT_TA}} vb. varsa,
    // replacements içindeki MT_1_, MT_2_ ... verilerini kullanarak
    // her ölçüm için satırlar çoğaltanır.
    // =====================================================================
    private int ProcessDynamicMicrotremorTables(
        Body body,
        IDictionary<string, string> variableReplacements,
        HashSet<string> replacedKeys)
    {
        // Kaç mikrotremör ölçümü var?
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

            // {{MT_NO}} veya {{MT_T0}} gibi mikrotremör Şablon satırını bul
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

            // Her ölçüm için satır klonla
            for (int n = 1; n <= mtCount; n++)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                // Bu ölçümün değişkenlerini replacements'tan al
                var mtReplacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["MT_NO"]     = n.ToString(),
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

            // Şablon satırını sil
            templateRow.Remove();
        }

        return totalReplacements;
    }
}
