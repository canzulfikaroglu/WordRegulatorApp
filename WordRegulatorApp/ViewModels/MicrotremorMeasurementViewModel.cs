using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

/// <summary>
/// Tek bir mikrotremör ölçümünün UI modeli.
/// Tablo 1: Ölçü no, T0, Ta, Tb, Frekans
/// Tablo 2: Ölçüm No, Tarih, Yatay koordinat, Düşey koordinat
/// </summary>
public partial class MicrotremorMeasurementViewModel : ObservableObject
{
    public int MeasurementNo { get; set; }

    public string DisplayName => $"Ölçüm {MeasurementNo}";

    // ===== Tablo 1: Ölçüm Veri Tablosu =====
    [ObservableProperty] private string _t0 = string.Empty;
    [ObservableProperty] private string _ta = string.Empty;
    [ObservableProperty] private string _tb = string.Empty;
    [ObservableProperty] private string _frequency = string.Empty;

    // ===== Tablo 2: Konum Tablosu =====
    [ObservableProperty] private string _date = string.Empty;
    [ObservableProperty] private string _horizontalCoordinate = string.Empty;
    [ObservableProperty] private string _verticalCoordinate = string.Empty;

    public MicrotremorMeasurementViewModel(int measurementNo)
    {
        MeasurementNo = measurementNo;
    }
}
