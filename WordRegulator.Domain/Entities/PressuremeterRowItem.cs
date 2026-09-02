namespace WordRegulator.Domain.Entities;

public class PressuremeterRowItem
{
    public string Depth { get; set; } = string.Empty;              // Derinlik (m)
    public string LimitPressure { get; set; } = string.Empty;      // Limit Basınç PL (kg/cm²)
    public string NetLimitPressure { get; set; } = string.Empty;   // Net Limit Basınç PL* (kg/cm²)
    public string ElasticityModulus { get; set; } = string.Empty;  // Elastisite Modülü Em (kg/cm²)
}
