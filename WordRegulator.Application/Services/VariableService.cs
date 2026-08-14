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
            // ================= 1. Ön Kapak =================
            new("PROVİNCE", "İl", "İl adı", VariableCategory.OnKapak, VariableType.Independent, "İstanbul", order: 1),
            new("DİSTRİCT", "İlçe", "İlçe adı", VariableCategory.OnKapak, VariableType.Independent, "Kartal", order: 2),
            new("NEİGHBOURHOOD", "Mahalle", "Mahalle adı", VariableCategory.OnKapak, VariableType.Independent, "Çavuşoğlu", order: 3),
            new("SHEET", "Pafta", "Pafta numarası / kodu", VariableCategory.OnKapak, VariableType.Independent, "G22A136MD", order: 4),
            new("İSLAND", "Ada", "Ada numarası", VariableCategory.OnKapak, VariableType.Independent, "2249", order: 5),
            new("PARCEL", "Parsel", "Parsel numarası", VariableCategory.OnKapak, VariableType.Independent, "85", order: 6),
            new("PROPRİETOR", "Mal Sahibi", "Mal sahibi adı / soyadı", VariableCategory.OnKapak, VariableType.Independent, "", order: 7),
            new("MONTH", "Ay", "Cari Ay (Otomatik)", VariableCategory.OnKapak, VariableType.Dependent, monthName, formula: "{AUTO_MONTH}", order: 8),
            new("YEAR", "Yıl", "Cari Yıl (Otomatik)", VariableCategory.OnKapak, VariableType.Dependent, yearString, formula: "{AUTO_YEAR}", order: 9),

            // ================= 1.2 KARTAL--Veri Kapak =================
            new("SHEETİSLANDPARCEL", "Pafta Ada Parsel (Birleşik)", "G22A136MD PAFTA 2249 ADA 85 PARSEL", VariableCategory.VeriKapak, VariableType.Independent, "", formula: "{SHEET} PAFTA {İSLAND} ADA {PARCEL} PARSEL", order: 10),
            new("ARCHİTECTURELEVEL", "Mimari 0.00 Kotu", "Mimari 0.00 Kotu", VariableCategory.VeriKapak, VariableType.Independent, "38,81", order: 11),
            new("BASEUPPERLEVELACCORDİNGTOLEVEL", "Temel Üst Kotu", "0.00 Kotuna Göre Temel Üst Kotu", VariableCategory.VeriKapak, VariableType.Independent, "-3,24(35,57)", order: 12),
            new("VS30", "VS30 Değeri", "VS30 hızı / değeri", VariableCategory.VeriKapak, VariableType.Independent, "", order: 13),
            new("LOCALSOİLCLASS", "Yerel Zemin Sınıfı", "Yerel Zemin Sınıfı (Örn: ZF, ZC)", VariableCategory.VeriKapak, VariableType.Independent, "ZF", order: 14),
            new("GROUNDWATERLEVEL", "Yeraltı Suyu Seviyesi", "Yeraltı Suyu Seviyesi", VariableCategory.VeriKapak, VariableType.Independent, "Yeraltı suyuna rastlanmamıştır.", order: 15),

            // ================= 3. Rapor =================
            new("LATİTUDE", "Enlem", "Enlem koordinatı", VariableCategory.Rapor, VariableType.Independent, "", order: 20),
            new("LONGİTUDE", "Boylam", "Boylam koordinatı", VariableCategory.Rapor, VariableType.Independent, "", order: 21),
            new("DRİLLİNGWELLCOUNT", "Sondaj Kuyusu Sayısı", "Sondaj kuyusu sayısı", VariableCategory.Rapor, VariableType.Independent, "3", order: 22),
            new("TOTALDRİLLİNG", "Toplam Sondaj", "Toplam sondaj metrajı", VariableCategory.Rapor, VariableType.Independent, "45.00m", order: 23),
            new("MAXKOT", "En Yüksek Kot", "En yüksek kot", VariableCategory.Rapor, VariableType.Independent, "39.50", order: 24),
            new("MİNKOT", "En Düşük Kot", "En düşük kot", VariableCategory.Rapor, VariableType.Independent, "36.20", order: 25),
            new("AREAOFSTUDY", "İnceleme Alanı (m²)", "İnceleme alanı (m² cinsinden)", VariableCategory.Rapor, VariableType.Independent, "500", order: 26),
            new("MİCROZONATİONTEXT", "Mikrobölgeleme Metni", "Mikrobölgeleme rapor açıklaması", VariableCategory.Rapor, VariableType.Independent, "Pendik Formasyonu-Kartal Üyesi (Dpk) ve UA (Uygun Alan)", order: 27),
            new("HOUSINGPLAN", "Yapı Kat Planı", "Bodrum + Zemin + 5 Normal Kat + Çatı Arası", VariableCategory.Rapor, VariableType.Independent, "Bodrum + Zemin + 5 Normal Kat + Çatı Arası", order: 28),
            new("SS", "Ss İvme Katsayısı", "Ss ivme parametresi", VariableCategory.Rapor, VariableType.Independent, "", order: 29),
            new("S1", "S1 İvme Katsayısı", "S1 ivme parametresi", VariableCategory.Rapor, VariableType.Independent, "", order: 30),
            new("PGA", "PGA Değeri", "PGA yer ivmesi", VariableCategory.Rapor, VariableType.Independent, "", order: 31),
            new("PGV", "PGV Değeri", "PGV yer hızı", VariableCategory.Rapor, VariableType.Independent, "", order: 32),
            new("LİTHOLOGİCALARRANGEMENT", "Litolojik Dizilim", "Litolojik yapı (örn: kumtaşı kiltaşı ardalanması)", VariableCategory.Rapor, VariableType.Independent, "kumtaşı kiltaşı ardalanması", order: 33),
            new("SDS", "SDS Katsayısı", "Tasarım ivme katsayısı SDS", VariableCategory.Rapor, VariableType.Independent, "", order: 34),
            new("SD1", "SD1 Katsayısı", "Tasarım ivme katsayısı SD1", VariableCategory.Rapor, VariableType.Independent, "", order: 35),
            new("SEATİNGAREA", "Oturum Alanı (m²)", "Oturum alanı (m² cinsinden)", VariableCategory.Rapor, VariableType.Independent, "250", order: 36),
            new("ZEROPOİNTZEROKOT", "±0.00 Kotu", "±0.00 kotu değeri", VariableCategory.Rapor, VariableType.Independent, "38,81", order: 37),
            new("FROMTHEBASETOPLEVEL", "Temel Üst Kotu Bilgisi", "Temel üst kotundan mesafe/bilgi", VariableCategory.Rapor, VariableType.Independent, "-3,24(35,57)", order: 38),
            new("FORMATİONTYPE", "Formasyon Tipi", "Formasyon tanımı", VariableCategory.Rapor, VariableType.Independent, "Pendik Formasyonu (Dp)", order: 39),
            new("DRILLER", "Sondör Adı Soyadı", "Sondör adı soyadı", VariableCategory.Rapor, VariableType.Independent, "VURAL ÖNCÜ", order: 40),
            new("DRİLLİNGDATE", "Sondaj Tarihi", "Sondaj yapılma tarihi", VariableCategory.Rapor, VariableType.Independent, "", order: 41),

            // Sismik Koordinatlar
            new("SEİSMİCCOORDİNATESS-1XSTART", "Sismik SS-1 X Başlangıç", "Sismik koordinat SS-1 X başlangıç", VariableCategory.Rapor, VariableType.Independent, "", order: 50),
            new("SEİSMİCCOORDİNATESS-1YSTART", "Sismik SS-1 Y Başlangıç", "Sismik koordinat SS-1 Y başlangıç", VariableCategory.Rapor, VariableType.Independent, "", order: 51),
            new("SEİSMİCCOORDİNATESS-2XSTART", "Sismik SS-2 X Başlangıç", "Sismik koordinat SS-2 X başlangıç", VariableCategory.Rapor, VariableType.Independent, "", order: 52),
            new("SEİSMİCCOORDİNATESS-2YSTART", "Sismik SS-2 Y Başlangıç", "Sismik koordinat SS-2 Y başlangıç", VariableCategory.Rapor, VariableType.Independent, "", order: 53),
            new("SEİSMİCCOORDİNATESS-1XFİNİSH", "Sismik SS-1 X Bitiş", "Sismik koordinat SS-1 X bitiş", VariableCategory.Rapor, VariableType.Independent, "", order: 54),
            new("SEİSMİCCOORDİNATESS-1YFİNİSH", "Sismik SS-1 Y Bitiş", "Sismik koordinat SS-1 Y bitiş", VariableCategory.Rapor, VariableType.Independent, "", order: 55),
            new("SEİSMİCCOORDİNATESS-2XFİNİSH", "Sismik SS-2 X Bitiş", "Sismik koordinat SS-2 X bitiş", VariableCategory.Rapor, VariableType.Independent, "", order: 56),
            new("SEİSMİCCOORDİNATESS-2YFİNİSH", "Sismik SS-2 Y Bitiş", "Sismik koordinat SS-2 Y bitiş", VariableCategory.Rapor, VariableType.Independent, "", order: 57),

            // Poisson Oranları
            new("POİSSONRATİO1S1", "Poisson Oranı 1. Tabaka 1. Serim", "Poisson oranı (1. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 60),
            new("POİSSONRATİO2S1", "Poisson Oranı 2. Tabaka 1. Serim", "Poisson oranı (2. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 61),
            new("POİSSONRATİO1S2", "Poisson Oranı 1. Tabaka 2. Serim", "Poisson oranı (1. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 62),
            new("POİSSONRATİO2S2", "Poisson Oranı 2. Tabaka 2. Serim", "Poisson oranı (2. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 63),

            // Kayma Modülü (Shear Modulus)
            new("SHEARMODULUS1S1", "Kayma Modülü 1. Tabaka 1. Serim", "Kayma modülü (1. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 70),
            new("SHEARMODULUS2S1", "Kayma Modülü 2. Tabaka 1. Serim", "Kayma modülü (2. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 71),
            new("SHEARMODULUS1S2", "Kayma Modülü 1. Tabaka 2. Serim", "Kayma modülü (1. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 72),
            new("SHEARMODULUS2S2", "Kayma Modülü 2. Tabaka 2. Serim", "Kayma modülü (2. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 73),

            // Bulk Modülü
            new("BULKMODULE1S1", "Bulk Modülü 1. Tabaka 1. Serim", "Bulk modülü (1. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 80),
            new("BULKMODULE2S1", "Bulk Modülü 2. Tabaka 1. Serim", "Bulk modülü (2. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 81),
            new("BULKMODULE1S2", "Bulk Modülü 1. Tabaka 2. Serim", "Bulk modülü (1. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 82),
            new("BULKMODULE2S2", "Bulk Modülü 2. Tabaka 2. Serim", "Bulk modülü (2. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 83),

            // Elastisite Modülü
            new("ELASTİCİTYMODULUS1S1", "Elastisite Modülü 1. Tabaka 1. Serim", "Elastisite modülü (1. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 90),
            new("ELASTİCİTYMODULUS2S1", "Elastisite Modülü 2. Tabaka 1. Serim", "Elastisite modülü (2. tabaka 1. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 91),
            new("ELASTİCİTYMODULUS1S2", "Elastisite Modülü 1. Tabaka 2. Serim", "Elastisite modülü (1. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 92),
            new("ELASTİCİTYMODULUS2S2", "Elastisite Modülü 2. Tabaka 2. Serim", "Elastisite modülü (2. tabaka 2. serim)", VariableCategory.Rapor, VariableType.Independent, "", order: 93),

            // Serim T0
            new("S1T0", "1. Serim T0", "1. Serim T0 değeri", VariableCategory.Rapor, VariableType.Independent, "", order: 100),
            new("S2T0", "2. Serim T0", "2. Serim T0 değeri", VariableCategory.Rapor, VariableType.Independent, "", order: 101)
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
            var combined = $"{sheet} PAFTA {island} ADA {PARCEL_FORMAT(parcel)}".Trim();
            result["SHEETİSLANDPARCEL"] = string.IsNullOrWhiteSpace(sheet) && string.IsNullOrWhiteSpace(island) && string.IsNullOrWhiteSpace(parcel)
                ? ""
                : $"{sheet} PAFTA {island} ADA {parcel} PARSEL".Trim();
        }

        var now = DateTime.Now;
        if (!userOverriddenKeys.Contains("MONTH"))
        {
            result["MONTH"] = now.ToString("MMMM", TrCulture);
        }

        if (!userOverriddenKeys.Contains("YEAR"))
        {
            result["YEAR"] = now.Year.ToString();
        }

        return result;

        string PARCEL_FORMAT(string prc) => prc;
    }
}
