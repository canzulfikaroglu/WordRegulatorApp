namespace WordRegulator.Domain.Entities;

public class Vs30SeriesItem
{
    public int SeriesNumber { get; set; }
    public string SeriesName { get; set; } = string.Empty; // Örn: SS-1, SS-2
    public double KotStart { get; set; }
    public double KotEnd { get; set; }
    public double KotAverage { get; set; }
    public double Df { get; set; }
    public double H1 { get; set; }
    public double AValue { get; set; }
    public string Vs30Result { get; set; } = string.Empty;
    public string Formula { get; set; } = string.Empty;
    public string FormulaFull { get; set; } = string.Empty;
    public List<Vs30LayerItem> Layers { get; set; } = new();
}
