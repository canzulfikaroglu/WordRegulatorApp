namespace WordRegulator.Domain.Entities;

public class MicrotremorMeasurementItem
{
    /// <summary>Ölçüm sıra numarası (1, 2, 3...)</summary>
    public int MeasurementNo { get; set; }

    // ===== Tablo 1: Ölçüm Veri Tablosu =====
    /// <summary>Hakim periyot değeri (sn)</summary>
    public string T0 { get; set; } = string.Empty;
    /// <summary>Alt periyot değeri</summary>
    public string Ta { get; set; } = string.Empty;
    /// <summary>Üst periyot değeri</summary>
    public string Tb { get; set; } = string.Empty;
    /// <summary>Frekans değeri (Hz)</summary>
    public string Frequency { get; set; } = string.Empty;

    // ===== Tablo 2: Konum Tablosu =====
    /// <summary>Ölçüm tarihi</summary>
    public string Date { get; set; } = string.Empty;
    /// <summary>Yatay koordinat (X)</summary>
    public string HorizontalCoordinate { get; set; } = string.Empty;
    /// <summary>Düşey koordinat (Y)</summary>
    public string VerticalCoordinate { get; set; } = string.Empty;
}
