namespace WordRegulator.Domain.Entities;

public class Vs30SeriesItem
{
    public int SeriesNumber { get; set; }
    public string SeriesName { get; set; } = string.Empty; 
    public double KotStart { get; set; }
    public double KotEnd { get; set; }
    public double KotAverage { get; set; }
    public double Df { get; set; }
    public double H1 { get; set; }
    public double AValue { get; set; }
    public string T0 { get; set; } = string.Empty;
    public string Vs30Result { get; set; } = string.Empty;
    public string SurfaceVs30 { get; set; } = string.Empty; 
    public string Formula { get; set; } = string.Empty;
    public string FormulaFull { get; set; } = string.Empty;
    public List<Vs30LayerItem> Layers { get; set; } = new();

    
    public string XStart  { get; set; } = string.Empty;
    public string YStart  { get; set; } = string.Empty;
    public string XFinish { get; set; } = string.Empty;
    public string YFinish { get; set; } = string.Empty;

    
    public string PoissonLayer1         { get; set; } = string.Empty;
    public string PoissonLayer2         { get; set; } = string.Empty;
    public string ShearModulusLayer1    { get; set; } = string.Empty;
    public string ShearModulusLayer2    { get; set; } = string.Empty;
    public string BulkModulusLayer1     { get; set; } = string.Empty;
    public string BulkModulusLayer2     { get; set; } = string.Empty;
    public string ElasticityModulusLayer1 { get; set; } = string.Empty;
    public string ElasticityModulusLayer2 { get; set; } = string.Empty;
}
