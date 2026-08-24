namespace WordRegulator.Domain.Entities;

public class GroundwaterMeasurementItem
{
    public string MeasurementDate { get; set; } = string.Empty; // Örn: 10.04.2026
    public string WaterDepth { get; set; } = "2,70";            // Örn: 2,70 veya "Kuyu çökmüş"
}
