using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class Vs30LayerViewModel : ObservableObject
{
    public int LayerNumber { get; set; }
    public string LayerName => $"{LayerNumber}. Tabaka";

    [ObservableProperty]
    private string _startDepthInput = "0.00";

    [ObservableProperty]
    private string _endDepthInput = "1.10"; // Max Derinlik (m)

    [ObservableProperty]
    private string _velocityInput = "157.00"; // Vs (m/s)

    public double StartDepth => double.TryParse(StartDepthInput.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0;
    public double EndDepth => double.TryParse(EndDepthInput.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0;
    public double Velocity => double.TryParse(VelocityInput.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0;

    public event EventHandler? LayerChanged;

    public Vs30LayerViewModel(int layerNumber, double defaultStartDepth, double defaultEndDepth, double defaultVelocity)
    {
        LayerNumber = layerNumber;
        _startDepthInput = defaultStartDepth.ToString("F2");
        _endDepthInput = defaultEndDepth.ToString("F2");
        _velocityInput = defaultVelocity.ToString("F2");
    }

    partial void OnStartDepthInputChanged(string value)
    {
        OnPropertyChanged(nameof(StartDepth));
        LayerChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnEndDepthInputChanged(string value)
    {
        OnPropertyChanged(nameof(EndDepth));
        LayerChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnVelocityInputChanged(string value)
    {
        OnPropertyChanged(nameof(Velocity));
        LayerChanged?.Invoke(this, EventArgs.Empty);
    }
}
