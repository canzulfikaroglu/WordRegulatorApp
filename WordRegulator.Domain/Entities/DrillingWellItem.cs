namespace WordRegulator.Domain.Entities;

public class DrillingWellItem
{
    public string WellName { get; set; } = string.Empty; 
    public string Latitude { get; set; } = string.Empty;  
    public string Longitude { get; set; } = string.Empty; 
    public string GroundKot { get; set; } = string.Empty; 
    public string BottomKot { get; set; } = string.Empty; 
    public double DepthMeters { get; set; }               
    public string GroundwaterStatus { get; set; } = "Yeraltı suyuna rastlanmamıştır."; 
    public bool IsCollapsed { get; set; }                  
    public string StartDate { get; set; } = string.Empty; 
    public string FinishDate { get; set; } = string.Empty;
    public List<GroundwaterMeasurementItem> Measurements { get; set; } = new(); 

    
    public bool HasPressuremeterTest { get; set; }
    public List<PressuremeterRowItem> PressuremeterRows { get; set; } = new();
}
