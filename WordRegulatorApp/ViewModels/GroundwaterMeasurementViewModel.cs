using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class GroundwaterMeasurementViewModel : ObservableObject
{
    public int Index { get; set; }

    [ObservableProperty]
    private string _measurementDate = string.Empty;

    [ObservableProperty]
    private string _waterDepth = "2,70";

    [ObservableProperty]
    private bool _isEditable = true;

    public event EventHandler? MeasurementChanged;

    public GroundwaterMeasurementViewModel(int index, string measurementDate, string waterDepth = "2,70", bool isEditable = true)
    {
        Index = index;
        _measurementDate = measurementDate;
        _waterDepth = waterDepth;
        _isEditable = isEditable;
    }

    partial void OnMeasurementDateChanged(string value) => MeasurementChanged?.Invoke(this, EventArgs.Empty);
    partial void OnWaterDepthChanged(string value) => MeasurementChanged?.Invoke(this, EventArgs.Empty);
}
