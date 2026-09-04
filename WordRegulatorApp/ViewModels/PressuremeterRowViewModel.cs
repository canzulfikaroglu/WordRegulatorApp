using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class PressuremeterRowViewModel : ObservableObject
{
    public event EventHandler? RowChanged;

    [ObservableProperty]
    private string _depth = string.Empty;           

    [ObservableProperty]
    private string _limitPressure = string.Empty;   

    [ObservableProperty]
    private string _netLimitPressure = string.Empty; 

    [ObservableProperty]
    private string _elasticityModulus = string.Empty; 

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
