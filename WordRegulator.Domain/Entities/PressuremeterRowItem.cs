namespace WordRegulator.Domain.Entities;

public class PressuremeterRowItem
{
    public string Depth { get; set; } = string.Empty;              
    public string LimitPressure { get; set; } = string.Empty;      
    public string NetLimitPressure { get; set; } = string.Empty;   
    public string ElasticityModulus { get; set; } = string.Empty;  
}
