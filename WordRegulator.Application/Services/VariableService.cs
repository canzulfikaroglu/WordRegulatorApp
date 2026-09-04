using System.Globalization;
using WordRegulator.Application.Interfaces;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Enums;

namespace WordRegulator.Application.Services;

public class VariableService : IVariableService
{
    private static readonly CultureInfo TrCulture = new CultureInfo("tr-TR");

    public IReadOnlyList<VariableDefinition> GetDefaultDefinitions()
    {
        var now = DateTime.Now;
        var monthName = now.ToString("MMMM", TrCulture);
        var yearString = now.Year.ToString();

        var list = new List<VariableDefinition>
        {
            
            new("PROVİNCE", "İl", "İl adı", VariableCategory.OnKapak, VariableType.Independent, "İstanbul", order: 1),
            new("DİSTRİCT", "İlçe", "İlçe adı", VariableCategory.OnKapak, VariableType.Independent, "Kartal", order: 2),
            new("NEİGHBOURHOOD", "Mahalle", "Mahalle adı", VariableCategory.OnKapak, VariableType.Independent, "Çavuşoğlu", order: 3),
            new("SHEET", "Pafta", "Pafta numarası / kodu", VariableCategory.OnKapak, VariableType.Independent, "250", order: 4),
            new("İSLAND", "Ada", "Ada numarası", VariableCategory.OnKapak, VariableType.Independent, "2249", order: 5),
            new("PARCEL", "Parsel", "Parsel numarası", VariableCategory.OnKapak, VariableType.Independent, "85", order: 6),
            new("PROPRİETOR", "Mal Sahibi", "Mal sahibi adı / soyadı", VariableCategory.OnKapak, VariableType.Independent, "", order: 7),
            
            new("MONTH", "Ay", "Ay bilgisi", VariableCategory.OnKapak, VariableType.Independent, monthName, order: 8),
            new("YEAR", "Yıl", "Yıl bilgisi", VariableCategory.OnKapak, VariableType.Independent, yearString, order: 9),

            
            new("SHEETİSLANDPARCEL", "Pafta Ada Parsel (Birleşik)", "G22A136MD PAFTA 2249 ADA 85 PARSEL", VariableCategory.VeriKapak, VariableType.Independent, "G22A136MD 250 PAFTA 2249 ADA 85 PARSEL", formula: "{SHEET} PAFTA {İSLAND} ADA {PARCEL} PARSEL", order: 10),
            
            new("LOCALSOİLCLASS", "Yerel Zemin Sınıfı", "Yerel Zemin Sınıfı (Örn: ZF, ZC)", VariableCategory.VeriKapak, VariableType.Independent, "ZF", order: 14),
            new("GROUNDWATERLEVEL", "Yeraltı Suyu Seviyesi", "Sondaj kuyuları su durumu", VariableCategory.VeriKapak, VariableType.Dependent, "Yeraltı suyuna rastlanmamıştır.", order: 15),

            
            new("LATİTUDE", "Enlem", "Enlem koordinatı", VariableCategory.Rapor, VariableType.Independent, "", order: 20),
            new("LONGİTUDE", "Boylam", "Boylam koordinatı", VariableCategory.Rapor, VariableType.Independent, "", order: 21),
            new("DRİLLİNGWELLCOUNT", "Sondaj Kuyusu Sayısı", "Sondaj kuyusu sayısı (Sondaj kuyuları panelinden otomatik)", VariableCategory.Rapor, VariableType.Dependent, "3", order: 22),
            new("TOTALDRİLLİNG", "Toplam Sondaj", "Toplam sondaj metrajı (Otomatik)", VariableCategory.Rapor, VariableType.Dependent, "45.00", order: 23),
            new("MAXKOT", "En Yüksek Kot", "En yüksek kot", VariableCategory.Rapor, VariableType.Independent, "39.50", order: 24),
            new("MİNKOT", "En Düşük Kot", "En düşük kot", VariableCategory.Rapor, VariableType.Independent, "36.20", order: 25),
            new("AREAOFSTUDY", "İnceleme Alanı (m²)", "İnceleme alanı (m²)", VariableCategory.Rapor, VariableType.Independent, "500", order: 26),
            new("HOUSINGPLAN", "Yapı Kat Planı", "Bodrum + Zemin + 5 Normal Kat + Çatı Arası", VariableCategory.Rapor, VariableType.Independent, "Bodrum + Zemin + 5 Normal Kat + Çatı Arası", order: 28),
            new("SS", "Ss İvme Katsayısı", "Ss ivme parametresi", VariableCategory.Rapor, VariableType.Independent, "", order: 29),
            new("S1", "S1 İvme Katsayısı", "S1 ivme parametresi", VariableCategory.Rapor, VariableType.Independent, "", order: 30),
            new("PGA", "PGA Değeri", "PGA yer ivmesi", VariableCategory.Rapor, VariableType.Independent, "", order: 31),
            new("PGV", "PGV Değeri", "PGV yer hızı", VariableCategory.Rapor, VariableType.Independent, "", order: 32),
            new("LİTHOLOGİCALARRANGEMENT", "Zemin Birimi", "Zemin birimi yapısı (örn: kumtaşı kiltaşı ardalanması)", VariableCategory.Rapor, VariableType.Independent, "kumtaşı kiltaşı ardalanması", order: 33),
            new("SDS", "SDS Katsayısı", "Tasarım ivme katsayısı SDS", VariableCategory.Rapor, VariableType.Independent, "", order: 34),
            new("SD1", "SD1 Katsayısı", "Tasarım ivme katsayısı SD1", VariableCategory.Rapor, VariableType.Independent, "", order: 35),
            new("SEATİNGAREA", "Oturum Alanı (m²)", "Oturum alanı (m²)", VariableCategory.Rapor, VariableType.Independent, "250", order: 36),
            new("ZEROPOİNTZEROKOT", "±0.00 Kotu", "±0.00 kotu değeri", VariableCategory.Rapor, VariableType.Independent, "38,81", order: 37),
            new("FROMTHEBASETOPLEVEL", "Temel Üst Kotu Bilgisi", "Temel üst kotundan mesafe/bilgi", VariableCategory.Rapor, VariableType.Dependent, "-3,24(35,57)", order: 38),
            new("FORMATİONTYPE", "Formasyon Tipi", "Formasyon adı (Dropdown'dan seçilir)", VariableCategory.Rapor, VariableType.Independent, "Pendik Formasyonu (Dp)", order: 39),
            new("FORMATİONTEXT", "Formasyon Açıklaması", "Seçilen formasyonun rapor metni (Otomatik)", VariableCategory.Rapor, VariableType.Dependent, "", order: 39),
            new("DRILLER", "Sondör Adı Soyadı", "Sondör adı soyadı", VariableCategory.Rapor, VariableType.Independent, "VURAL ÖNCÜ", order: 40),
            new("DRİLLİNGDATE", "Sondaj Tarihi", "Sondaj yapılma tarihi", VariableCategory.Rapor, VariableType.Independent, "", order: 41),
        };

        return list.OrderBy(x => x.Order).ToList();
    }

    public IDictionary<string, string> RecalculateDependentVariables(IDictionary<string, string> currentValues, HashSet<string> userOverriddenKeys)
    {
        var result = new Dictionary<string, string>(currentValues, StringComparer.OrdinalIgnoreCase);

        var sheet = result.TryGetValue("SHEET", out var s) ? s : "";
        var island = result.TryGetValue("İSLAND", out var isl) ? isl : (result.TryGetValue("ISLAND", out var isl2) ? isl2 : "");
        var parcel = result.TryGetValue("PARCEL", out var p) ? p : "";

        if (!userOverriddenKeys.Contains("SHEETİSLANDPARCEL"))
        {
            if (string.IsNullOrWhiteSpace(sheet) && string.IsNullOrWhiteSpace(island) && string.IsNullOrWhiteSpace(parcel))
            {
                result["SHEETİSLANDPARCEL"] = "";
            }
            else
            {
                result["SHEETİSLANDPARCEL"] = $"{sheet} PAFTA {island} ADA {parcel} PARSEL".Trim();
            }
        }

        return result;
    }

    public string CalculateVs30(IEnumerable<Vs30LayerItem> layers, double df)
    {
        var layerList = layers.OrderBy(l => l.LayerNumber).ToList();
        if (layerList.Count == 0) return "0 m/s";

        double sumThickness = 0;
        double sumTime = 0;

        for (int i = 0; i < layerList.Count; i++)
        {
            var layer = layerList[i];
            double h_i = 0;

            if (i == 0)
            {
                
                h_i = Math.Max(0, layer.EndDepth - df);
            }
            else if (i < layerList.Count - 1)
            {
                
                h_i = Math.Max(0, layer.EndDepth - layerList[i - 1].EndDepth);
            }
            else
            {
                
                h_i = Math.Max(0, 30.0 - sumThickness);
            }

            sumThickness += h_i;

            if (layer.Velocity > 0)
            {
                sumTime += (h_i / layer.Velocity);
            }
        }

        if (sumTime <= 0) return "0 m/s";

        double vs30 = 30.0 / sumTime;
        return $"{Math.Round(vs30)} m/s";
    }

    public (string totalDrilling, string groundwaterLevel) SummarizeDrillingWells(IEnumerable<DrillingWellItem> wells)
    {
        var wellList = wells.ToList();
        if (wellList.Count == 0)
        {
            return ("0.00m", "Yeraltı suyuna rastlanmamıştır.");
        }

        double sumDepth = wellList.Sum(w => w.DepthMeters);
        string totalDrillingStr = $"{sumDepth:F2}";

        var statusParts = new List<string>();

        foreach (var well in wellList)
        {
            string wellEntry;
            if (well.IsCollapsed)
            {
                wellEntry = $"{well.WellName}: Kuyu çökmüş";
            }
            else
            {
                
                var lastMeasurement = well.Measurements.LastOrDefault();
                string lastDepthRaw = lastMeasurement?.WaterDepth ?? string.Empty;
                var normalizedLast = lastDepthRaw.Replace(',', '.');
                bool isNumeric = double.TryParse(normalizedLast, NumberStyles.Any, CultureInfo.InvariantCulture, out var lastDepthVal);

                if (isNumeric)
                {
                    string formattedDepth = lastDepthVal.ToString("F2", TrCulture);
                    wellEntry = $"{well.WellName}: {formattedDepth} m.";
                }
                else if (!string.IsNullOrWhiteSpace(well.GroundwaterStatus))
                {
                    wellEntry = $"{well.WellName}: {well.GroundwaterStatus}";
                }
                else
                {
                    wellEntry = $"{well.WellName}: Yeraltı suyuna rastlanmamıştır";
                }
            }
            statusParts.Add(wellEntry);
        }

        string groundwaterStr = statusParts.Count > 0
            ? string.Join(", ", statusParts)
            : "Yeraltı suyuna rastlanmamıştır.";

        return (totalDrillingStr, groundwaterStr);
    }

    public IDictionary<string, string> GenerateDrillingWellReplacements(IEnumerable<DrillingWellItem> wells, double baseTopLevel = 0.0)
    {
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var wellList = wells.ToList();

        
        var skBaseDeepParts = new List<string>();
        for (int i = 0; i < wellList.Count; i++)
        {
            var w = wellList[i];
            bool isLast = (i == wellList.Count - 1);

            
            string groundKotRaw = w.GroundKot.Replace(',', '.');
            double.TryParse(groundKotRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var groundKotVal);

            double depth = groundKotVal - baseTopLevel;
            if (depth < 0) depth = 0;

            string groundKotFormatted = groundKotVal.ToString("F2", TrCulture);
            string depthFormatted = depth.ToString("F2", TrCulture);

            string suffix = isLast ? " derinliğe inilecek ve bina temelleri dolgu birimine oturtulacaktır." 
                                   : " derinliğe inilecek ve bina temelleri dolgu birimine,";

            skBaseDeepParts.Add($"{w.WellName} sondajında {groundKotFormatted} kotundan temel üst kotu dikkate alınarak {depthFormatted} metre{suffix}");
        }
        replacements["SKBASEDEEP"] = string.Join("\n", skBaseDeepParts);

        
        for (int i = 0; i < wellList.Count; i++)
        {
            var w = wellList[i];
            int idx = i + 1;
            replacements[$"SK{idx}_NAME"] = w.WellName;
            replacements[$"SK{idx}_LAT"] = w.Latitude;
            replacements[$"SK{idx}_LON"] = w.Longitude;
            replacements[$"SK{idx}_GROUNDKOT"] = w.GroundKot;
            replacements[$"SK{idx}_BOTTOMKOT"] = w.BottomKot;
            replacements[$"SK{idx}_DEPTH"] = $"{w.DepthMeters:F2}";
            replacements[$"SK{idx}_WATER"] = w.GroundwaterStatus;
            replacements[$"SK{idx}_RESULT"] = w.GroundwaterStatus;
            replacements[$"SK{idx}_COLLAPSED"] = w.IsCollapsed ? "Evet" : "Hayır";
            replacements[$"SK{idx}_IS_COLLAPSED"] = w.IsCollapsed ? "Evet" : "Hayır";
            replacements[$"SK{idx}_START"] = w.StartDate;
            replacements[$"SK{idx}_FINISH"] = w.FinishDate;

            if (w.Measurements.Count > 0)
            {
                
                var lastM = w.Measurements.Last();
                replacements[$"SK{idx}_MEASURE_DATE"] = lastM.MeasurementDate;
                replacements[$"SK{idx}_WATER_DEPTH"] = lastM.WaterDepth;

                
                for (int m = 0; m < w.Measurements.Count; m++)
                {
                    int mIdx = m + 1;
                    replacements[$"SK{idx}_MEASURE{mIdx}_DATE"] = w.Measurements[m].MeasurementDate;
                    replacements[$"SK{idx}_MEASURE{mIdx}_DEPTH"] = w.Measurements[m].WaterDepth;
                }
            }
            else
            {
                replacements[$"SK{idx}_MEASURE_DATE"] = string.Empty;
                replacements[$"SK{idx}_WATER_DEPTH"] = w.IsCollapsed ? "Kuyu çökmüş" : string.Empty;
            }

            
            replacements[$"SK{idx}_PM_ZEMIN_KOTU"] = w.GroundKot;
            replacements[$"SK{idx}_PM_COUNT"] = w.PressuremeterRows.Count.ToString();
        }
        
        var pmWells = wellList.Where(w => w.HasPressuremeterTest).Select(w => w.WellName).ToList();
        if (pmWells.Count > 1)
        {
            var last = pmWells.Last();
            var others = string.Join(", ", pmWells.Take(pmWells.Count - 1));
            replacements["PM_WELL_NAMES"] = $"{others} ve {last}";
        }
        else if (pmWells.Count == 1)
        {
            replacements["PM_WELL_NAMES"] = pmWells.First();
        }
        else
        {
            replacements["PM_WELL_NAMES"] = "";
        }
        replacements["PM_WELLS"] = replacements["PM_WELL_NAMES"];

        int seismicCount = Math.Max(0, wellList.Count - 1);
        replacements["SEISMIC_MEASUREMENT_COUNT"] = seismicCount.ToString();
        replacements["SEISMIC_COUNT"] = seismicCount.ToString();

        return replacements;
    }

    public IDictionary<string, string> GenerateVs30Replacements(IEnumerable<Vs30LayerItem> layers, double df)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var layerList = layers.OrderBy(l => l.LayerNumber).ToList();

        
        result["VS30_DF"] = df.ToString("F2", TrCulture);

        double sumThickness = 0;
        double sumTime = 0;
        var formulaParts = new List<string>();

        for (int i = 0; i < layerList.Count; i++)
        {
            var layer = layerList[i];
            int num = i + 1;
            double h_i;

            if (i == 0)
                h_i = Math.Max(0, layer.EndDepth - df);
            else if (i < layerList.Count - 1)
                h_i = Math.Max(0, layer.EndDepth - layerList[i - 1].EndDepth);
            else
                h_i = Math.Max(0, 30.0 - sumThickness);

            sumThickness += h_i;
            if (layer.Velocity > 0)
                sumTime += (h_i / layer.Velocity);

            
            string startStr = layer.StartDepth.ToString("F2", TrCulture);
            string endStr = layer.EndDepth.ToString("F2", TrCulture);
            string vsStr = layer.Velocity.ToString("F0", TrCulture);
            string hStr = h_i.ToString("F2", TrCulture);

            result[$"VS30_START{num}"] = startStr;
            result[$"VS30_END{num}"] = endStr;
            result[$"VS30_RANGE{num}"] = $"{startStr} - {endStr}";
            result[$"VS30_H{num}"] = hStr;
            result[$"VS30_VS{num}"] = vsStr;

            
            formulaParts.Add($"({hStr}/{vsStr})");
        }

        
        string formulaBody = string.Join("+", formulaParts);
        result["VS30_FORMULA"] = $"30/[{formulaBody}]";

        
        double vs30Val = sumTime > 0 ? 30.0 / sumTime : 0;
        string vs30Str = $"{Math.Round(vs30Val)} m/s";

        
        result["VS30_FORMULA_FULL"] = $"Vs30 = 30/[{formulaBody}] = {vs30Str}";

        return result;
    }
}
