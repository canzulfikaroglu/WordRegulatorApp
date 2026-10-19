using System.Globalization;

namespace WordRegulator.Domain.Helpers;

public static class SeismicEvaluationHelper
{
    /// <summary>
    /// Poisson Oranı Değerlendirmesi:
    /// 0,0 - 0,25  => Gözeneksiz
    /// 0,26 - 0,35 => Orta Gözenekli
    /// 0,36 - 0,50 => Gözenekli
    /// </summary>
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

    /// <summary>
    /// Kayma Modülü (kg/cm²) Değerlendirmesi:
    /// 0 - 600      => Gevşek
    /// 600 - 3000   => Orta Sağlam
    /// 3000 - 10000 => Sağlam
    /// >= 10000     => Çok Sağlam
    /// </summary>
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

    /// <summary>
    /// Bulk Modülü (kg/cm²) Sıkışma Sınıfları:
    /// &lt; 400            => Çok Az
    /// 400 - 10000        => Az
    /// 10000 - 40000      => Orta
    /// 40000 - 100000     => Yüksek
    /// >= 100000          => Çok Yüksek
    /// </summary>
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

    /// <summary>
    /// Elastisite Modülü (kg/cm²) Değerlendirmesi:
    /// 0 - 2000     => Gevşek
    /// 2000 - 10000 => Orta Sağlam
    /// 10000 - 30000 => Sağlam
    /// >= 30000     => Çok Sağlam
    /// </summary>
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
