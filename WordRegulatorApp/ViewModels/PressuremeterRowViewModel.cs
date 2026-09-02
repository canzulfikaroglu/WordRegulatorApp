using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

/// <summary>
/// Presiyometre deneyi tek satır verisi (Derinlik | PL | PL* | Em).
/// </summary>
public partial class PressuremeterRowViewModel : ObservableObject
{
    public event EventHandler? RowChanged;

    [ObservableProperty]
    private string _depth = string.Empty;           // Derinlik (m)

    [ObservableProperty]
    private string _limitPressure = string.Empty;   // Limit Basınç PL

    [ObservableProperty]
    private string _netLimitPressure = string.Empty; // Net Limit Basınç PL*

    [ObservableProperty]
    private string _elasticityModulus = string.Empty; // Elastisite Modülü Em

    public int RowIndex { get; set; }

    public PressuremeterRowViewModel(int rowIndex)
    {
        RowIndex = rowIndex;
    }

    partial void OnDepthChanged(string value) => RowChanged?.Invoke(this, EventArgs.Empty);
    partial void OnLimitPressureChanged(string value) => RowChanged?.Invoke(this, EventArgs.Empty);
    partial void OnNetLimitPressureChanged(string value) => RowChanged?.Invoke(this, EventArgs.Empty);
    partial void OnElasticityModulusChanged(string value) => RowChanged?.Invoke(this, EventArgs.Empty);
}
