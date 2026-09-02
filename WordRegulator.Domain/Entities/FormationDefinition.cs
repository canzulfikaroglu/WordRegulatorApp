namespace WordRegulator.Domain.Entities;


public class FormationDefinition
{
    public string Name { get; set; } = string.Empty;        // Dropdown'da görünen isim (örn: Pendik Formasyonu (Dp))
    public string Description { get; set; } = string.Empty; // {{FORMATİONTEXT}} değişkenine yerleşen açıklama

    public FormationDefinition(string name, string description)
    {
        Name = name;
        Description = description;
    }

    public override string ToString() => Name;
}
