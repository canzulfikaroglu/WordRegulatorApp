namespace WordRegulator.Domain.Entities;

public class Vs30LayerItem
{
    public int LayerNumber { get; set; }
    public double StartDepth { get; set; } // Örn: 1.10
    public double EndDepth { get; set; }   // Örn: 7.50 (Max Derinlik)
    public double Velocity { get; set; }   // Kayma Dalga Hızı (Vs) m/s
}
