using WordRegulator.Domain.Enums;

namespace WordRegulator.Domain.Entities;

public class VariableDefinition
{
    public string Key { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string Hint { get; set; } = string.Empty;
    public VariableCategory Category { get; set; }
    public VariableType Type { get; set; }
    public string DefaultValue { get; set; } = string.Empty;
    public string? Formula { get; set; } // Bağımlı değişkenler için formül veya şablon metin (örn: "{SHEET} PAFTA {İSLAND} ADA {PARCEL} PARSEL")
    public int Order { get; set; }

    public VariableDefinition() { }

    public VariableDefinition(string key, string displayName, string hint, VariableCategory category, VariableType type = VariableType.Independent, string defaultValue = "", string? formula = null, int order = 0)
    {
        Key = key;
        DisplayName = displayName;
        Hint = hint;
        Category = category;
        Type = type;
        DefaultValue = defaultValue;
        Formula = formula;
        Order = order;
    }
}
