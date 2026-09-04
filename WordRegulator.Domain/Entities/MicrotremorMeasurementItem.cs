namespace WordRegulator.Domain.Entities;

public class MicrotremorMeasurementItem
{
        public int MeasurementNo { get; set; }

    
        public string T0 { get; set; } = string.Empty;
        public string Ta { get; set; } = string.Empty;
        public string Tb { get; set; } = string.Empty;
        public string Frequency { get; set; } = string.Empty;

    
        public string Date { get; set; } = string.Empty;
        public string HorizontalCoordinate { get; set; } = string.Empty;
        public string VerticalCoordinate { get; set; } = string.Empty;
}
