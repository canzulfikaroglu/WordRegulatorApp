using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class MicrotremorMeasurementViewModel : ObservableObject
{
    public int MeasurementNo { get; set; }

    public string DisplayName => $"Ölçüm {MeasurementNo}";

    
    [ObservableProperty] private string _t0 = string.Empty;
    [ObservableProperty] private string _ta = string.Empty;
    [ObservableProperty] private string _tb = string.Empty;
    [ObservableProperty] private string _frequency = string.Empty;

    
    [ObservableProperty] private string _date = string.Empty;
    [ObservableProperty] private string _horizontalCoordinate = string.Empty;
    [ObservableProperty] private string _verticalCoordinate = string.Empty;

    public MicrotremorMeasurementViewModel(int measurementNo)
    {
        MeasurementNo = measurementNo;
    }
}
