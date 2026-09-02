using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace WordRegulatorApp.ViewModels;

public partial class Vs30SeriesViewModel : ObservableObject
{
    private static readonly CultureInfo InvCulture = CultureInfo.InvariantCulture;
    private static readonly CultureInfo TrCulture = new CultureInfo("tr-TR");

    public int SeriesNumber { get; set; }
    public string SeriesName => $"{SeriesNumber}. Serim (SS-{SeriesNumber})";

    public List<int> LayerCountOptions { get; } = new() { 4, 5 };

    [ObservableProperty]
    private int _selectedLayerCount = 5;

    /// <summary>Serim kot başlangıç değeri (örn: 36,28)</summary>
    [ObservableProperty]
    private string _kotStartInput = "0.00";

    /// <summary>Serim kot bitiş değeri (örn: 36,50)</summary>
    [ObservableProperty]
    private string _kotEndInput = "0.00";

    /// <summary>Serim T0 değeri (örn: 0,18)</summary>
    [ObservableProperty]
    private string _t0Input = "";

    /// <summary>Yüzeyde ölçülen VS30 değeri (labaratuvardan elle girilir, m/sn)</summary>
    [ObservableProperty]
    private string _surfaceVs30Input = "";

    /// <summary>Hesaplanan Df: (kotOrtalama) - baseTopLevel. Temel Üst Kotu değişkeninden alınır.</summary>
    [ObservableProperty]
    private string _dfDisplay = "0,00 m";

    /// <summary>Hesaplanan VS30 sonucu</summary>
    [ObservableProperty]
    private string _vs30Result = "— m/s";

    /// <summary>Hesaplanan A katsayısı: A = 68 * (Vs30)^(-0.6)</summary>
    [ObservableProperty]
    private string _aDisplay = "—";

    public double AValue { get; private set; }

    public double KotStart => ParseDouble(KotStartInput);
    public double KotEnd => ParseDouble(KotEndInput);

    /// <summary>Ortalama kot = (KotStart + KotEnd) / 2</summary>
    public double KotAverage => (KotStart + KotEnd) / 2.0;

    public ObservableCollection<Vs30LayerViewModel> Layers { get; } = new();

    public event EventHandler? SeriesChanged;

    private bool _isSyncingLayers;

    public Vs30SeriesViewModel(int seriesNumber, double defaultKotStart = 36.28, double defaultKotEnd = 36.50, string defaultT0 = "")
    {
        SeriesNumber = seriesNumber;
        _kotStartInput = defaultKotStart.ToString("F2");
        _kotEndInput = defaultKotEnd.ToString("F2");
        _t0Input = defaultT0;
        SetupLayers();
    }

    public void SetupLayers()
    {
        Tuple<double, double, double>[] defaults;
        if (SeriesNumber == 1)
        {
            defaults = new Tuple<double, double, double>[]
            {
                new(0.00, 2.00, 120.0),
                new(2.00, 7.50, 491.0),
                new(7.50, 13.50, 570.0),
                new(13.50, 21.00, 767.0),
                new(21.00, 30.00, 821.0)
            };
        }
        else if (SeriesNumber == 2)
        {
            defaults = new Tuple<double, double, double>[]
            {
                new(0.00, 2.20, 118.0),
                new(2.20, 7.50, 452.0),
                new(7.50, 13.50, 624.0),
                new(13.50, 21.00, 782.0),
                new(21.00, 30.00, 816.0)
            };
        }
        else
        {
            defaults = new Tuple<double, double, double>[]
            {
                new(0.00, 2.00, 120.0),
                new(2.00, 7.50, 450.0),
                new(7.50, 13.50, 600.0),
                new(13.50, 21.00, 750.0),
                new(21.00, 30.00, 820.0)
            };
        }

        var oldValues = Layers.Select(l => new Tuple<double, double, double>(l.StartDepth, l.EndDepth, l.Velocity)).ToList();

        foreach (var old in Layers)
        {
            old.LayerChanged -= OnLayerChanged;
            old.DepthPropertyChanged -= OnLayerDepthPropertyChanged;
        }
        Layers.Clear();

        for (int i = 1; i <= SelectedLayerCount; i++)
        {
            Tuple<double, double, double> def;
            if (i - 1 < oldValues.Count)
            {
                def = oldValues[i - 1];
            }
            else if (i - 1 < defaults.Length)
            {
                def = defaults[i - 1];
            }
            else
            {
                def = new Tuple<double, double, double>(0.0, i * 6.0, 300.0 + (i * 100));
            }

            var layer = new Vs30LayerViewModel(i, def.Item1, def.Item2, def.Item3);
            layer.LayerChanged += OnLayerChanged;
            layer.DepthPropertyChanged += OnLayerDepthPropertyChanged;
            Layers.Add(layer);
        }

        SyncLayerChains();
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void SyncLayerChains()
    {
        if (_isSyncingLayers) return;
        _isSyncingLayers = true;
        try
        {
            for (int i = 0; i < Layers.Count - 1; i++)
            {
                if (Layers[i + 1].StartDepthInput != Layers[i].EndDepthInput)
                {
                    Layers[i + 1].StartDepthInput = Layers[i].EndDepthInput;
                }
            }
        }
        finally
        {
            _isSyncingLayers = false;
        }
    }

    private void OnLayerChanged(object? sender, EventArgs e)
    {
        if (_isSyncingLayers) return;
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void OnLayerDepthPropertyChanged(object? sender, string propName)
    {
        if (_isSyncingLayers) return;
        _isSyncingLayers = true;
        try
        {
            if (sender is Vs30LayerViewModel changedLayer)
            {
                int idx = Layers.IndexOf(changedLayer);
                if (idx >= 0)
                {
                    if (propName == "EndDepth" && idx < Layers.Count - 1)
                    {
                        // Üst değer değişti -> Sonraki tabakanın alt değerini otomatik güncelle
                        if (Layers[idx + 1].StartDepthInput != changedLayer.EndDepthInput)
                        {
                            Layers[idx + 1].StartDepthInput = changedLayer.EndDepthInput;
                        }
                    }
                    else if (propName == "StartDepth" && idx > 0)
                    {
                        // Alt değer değişti -> Önceki tabakanın üst değerini otomatik güncelle
                        if (Layers[idx - 1].EndDepthInput != changedLayer.StartDepthInput)
                        {
                            Layers[idx - 1].EndDepthInput = changedLayer.StartDepthInput;
                        }
                    }
                }
            }
        }
        finally
        {
            _isSyncingLayers = false;
        }

        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnSelectedLayerCountChanged(int value)
    {
        SetupLayers();
    }

    partial void OnKotStartInputChanged(string value)
    {
        OnPropertyChanged(nameof(KotStart));
        OnPropertyChanged(nameof(KotAverage));
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnKotEndInputChanged(string value)
    {
        OnPropertyChanged(nameof(KotEnd));
        OnPropertyChanged(nameof(KotAverage));
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnT0InputChanged(string value)
    {
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnSurfaceVs30InputChanged(string value)
    {
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Df ve VS30 sonucunu günceller. A = 68 * (Vs30)^(-0.6) katsayısını hesaplar.
    /// </summary>
    public void UpdateResults(double df, string vs30Result)
    {
        DfDisplay = df.ToString("F2", TrCulture) + " m";
        Vs30Result = vs30Result;

        // "1101 m/s" veya "1101" metninden sayısal Vs30 değerini ayrıştır
        var digits = new string(vs30Result.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
        if (double.TryParse(digits, NumberStyles.Any, InvCulture, out var vs30Val) && vs30Val > 0)
        {
            // A = 68 * (Vs30)^(-0.6)
            AValue = 68.0 * Math.Pow(vs30Val, -0.6);
            ADisplay = AValue.ToString("F2", TrCulture);
        }
        else
        {
            AValue = 0;
            ADisplay = "—";
        }
    }

    /// <summary>
    /// Görev 5: df değerine göre 1. tabakanın dahil edilip edilmeyeceğini belirler.
    /// Kural: df >= 1. tabakanın EndDepth değeri → 1. tabaka hariç tutulur (IsExcluded = true).
    /// Diğer tabakalar her zaman dahildir.
    /// </summary>
    public void UpdateLayerExclusion(double df)
    {
        for (int i = 0; i < Layers.Count; i++)
        {
            if (i == 0)
            {
                // 1. tabaka: df, tabakanın üst değerinden (EndDepth) büyük veya eşitse hariç tut
                Layers[i].IsExcluded = df >= Layers[i].EndDepth;
            }
            else
            {
                // Diğer tabakalar her zaman dahil
                Layers[i].IsExcluded = false;
            }
        }
    }

    private static double ParseDouble(string input)
        => double.TryParse(input.Replace(',', '.'), NumberStyles.Any, InvCulture, out var v) ? v : 0.0;
}
