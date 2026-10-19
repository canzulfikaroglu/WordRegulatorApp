namespace WordRegulator.Domain.Entities;

public class ResidentialSuitabilityDefinition
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;

    public ResidentialSuitabilityDefinition(string name, string description)
    {
        Name = name;
        Description = description;
    }

    public override string ToString() => Name;
}
