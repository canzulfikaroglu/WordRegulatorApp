using System.Globalization;

namespace WordRegulator.Domain.Helpers;

public static class SeismicEvaluationHelper
{
    public static string EvaluatePoisson(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var clean = input.Trim().Replace(" ", "").Replace(',', '.');
        if (!double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            return string.Empty;

        double rounded = Math.Round(val, 2);
        if (rounded <= 0.25) return "Gözeneksiz";
        if (rounded <= 0.35) return "Orta Gözenekli";
        return "Gözenekli";
    }

    public static string EvaluateShearModulus(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var clean = input.Trim().Replace(" ", "").Replace(',', '.');
        if (!double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            return string.Empty;

        if (val < 600) return "Gevşek";
        if (val < 3000) return "Orta Sağlam";
        if (val < 10000) return "Sağlam";
        return "Çok Sağlam";
    }

    public static string EvaluateBulkModulus(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var clean = input.Trim().Replace(" ", "").Replace(',', '.');
        if (!double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            return string.Empty;

        if (val < 400) return "Çok Az";
        if (val < 10000) return "Az";
        if (val < 40000) return "Orta";
        if (val < 100000) return "Yüksek";
        return "Çok Yüksek";
    }

    public static string EvaluateElasticityModulus(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        var clean = input.Trim().Replace(" ", "").Replace(',', '.');
        if (!double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
            return string.Empty;

        if (val < 2000) return "Gevşek";
        if (val < 10000) return "Orta Sağlam";
        if (val < 30000) return "Sağlam";
        return "Çok Sağlam";
    }
}
