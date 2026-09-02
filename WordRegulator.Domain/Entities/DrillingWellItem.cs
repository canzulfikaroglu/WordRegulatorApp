namespace WordRegulator.Domain.Entities;

public class DrillingWellItem
{
    public string WellName { get; set; } = string.Empty; // Örn: SK-1, SK-2
    public string Latitude { get; set; } = string.Empty;  // Enlem
    public string Longitude { get; set; } = string.Empty; // Boylam
    public string GroundKot { get; set; } = string.Empty; // Arazi Kotu
    public string BottomKot { get; set; } = string.Empty; // Sondaj Dip Kotu
    public double DepthMeters { get; set; }               // Derinlik (m)
    public string GroundwaterStatus { get; set; } = "Yeraltı suyuna rastlanmamıştır."; // Kuyu su durumu
    public bool IsCollapsed { get; set; }                  // Kuyu çökmesi var mı?
    public string StartDate { get; set; } = string.Empty; // Başlangıç Tarihi
    public string FinishDate { get; set; } = string.Empty;// Bitiş Tarihi
    public List<GroundwaterMeasurementItem> Measurements { get; set; } = new(); // Yeraltı suyu ölçüm listesi

    // Presiyometre Deneyi
    public bool HasPressuremeterTest { get; set; }
    public List<PressuremeterRowItem> PressuremeterRows { get; set; } = new();
}
