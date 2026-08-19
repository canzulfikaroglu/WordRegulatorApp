using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class DrillingWellViewModel : ObservableObject
{
    public string WellName { get; set; } = string.Empty;

    [ObservableProperty]
    private string _latitude = string.Empty;

    [ObservableProperty]
    private string _longitude = string.Empty;

    [ObservableProperty]
    private string _groundKot = string.Empty;

    [ObservableProperty]
    private string _bottomKot = string.Empty;

    [ObservableProperty]
    private string _depthInput = "15.00";

    [ObservableProperty]
    private string _groundwaterStatus = "Yeraltı suyuna rastlanmamıştır.";

    [ObservableProperty]
    private string _startDate = DateTime.Now.ToString("dd.MM.yyyy");

    [ObservableProperty]
    private string _finishDate = DateTime.Now.ToString("dd.MM.yyyy");

    public double DepthMeters => double.TryParse(DepthInput.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0;

    public event EventHandler? WellChanged;

    public DrillingWellViewModel(string wellName, double defaultDepth = 15.0)
    {
        WellName = wellName;
        _depthInput = defaultDepth.ToString("F2");
    }

    partial void OnDepthInputChanged(string value)
    {
        OnPropertyChanged(nameof(DepthMeters));
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnGroundwaterStatusChanged(string value)
    {
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnLatitudeChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnLongitudeChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnGroundKotChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnBottomKotChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnStartDateChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnFinishDateChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
}
