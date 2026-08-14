using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Enums;

namespace WordRegulatorApp.ViewModels;

public partial class VariableItemViewModel : ObservableObject
{
    private readonly VariableDefinition _definition;

    public string Key => _definition.Key;
    public string DisplayName => _definition.DisplayName;
    public string Hint => _definition.Hint;
    public VariableCategory Category => _definition.Category;
    public VariableType Type => _definition.Type;
    public bool IsDependent => Type == VariableType.Dependent;
    public bool IsEnabled => !IsDependent; // Bağımlı değişkenler kapalı (disabled)
    public bool IsReadOnly => IsDependent;
    public int Order => _definition.Order;

    [ObservableProperty]
    private string _value = string.Empty;

    [ObservableProperty]
    private bool _isManualOverride;

    public event EventHandler<string>? ValueChanged;

    public VariableItemViewModel(VariableDefinition definition)
    {
        _definition = definition;
        _value = definition.DefaultValue;
    }

    partial void OnValueChanged(string value)
    {
        ValueChanged?.Invoke(this, value);
    }

    [RelayCommand]
    private void ResetToDefault()
    {
        IsManualOverride = false;
        Value = _definition.DefaultValue;
    }
}
