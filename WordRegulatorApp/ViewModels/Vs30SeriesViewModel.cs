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
    public List<int> LayerCountOptions2 { get; } = new() { 4, 5, 6, 7, 8, 9, 10 };

    [ObservableProperty]
    private int _selectedLayerCount = 5;

    [ObservableProperty]
    private int _selectedLayerCount2 = 5;

        [ObservableProperty]
    private string _kotStartInput = "0.00";

        [ObservableProperty]
    private string _kotEndInput = "0.00";

        [ObservableProperty]
    private string _t0Input = "";

        [ObservableProperty]
    private string _surfaceVs30Input = "";

    
    [ObservableProperty] private string _xStart = "";
    [ObservableProperty] private string _yStart = "";
    [ObservableProperty] private string _xFinish = "";
    [ObservableProperty] private string _yFinish = "";

    
    [ObservableProperty] private string _poissonLayer1 = "";
    [ObservableProperty] private string _shearModulusLayer1 = "";
    [ObservableProperty] private string _bulkModulusLayer1 = "";
    [ObservableProperty] private string _elasticityModulusLayer1 = "";

    
    [ObservableProperty] private string _poissonLayer2 = "";
    [ObservableProperty] private string _shearModulusLayer2 = "";
    [ObservableProperty] private string _bulkModulusLayer2 = "";
    [ObservableProperty] private string _elasticityModulusLayer2 = "";

        [ObservableProperty]
    private string _dfDisplay = "0,00 m";

        [ObservableProperty]
    private string _vs30Result = "— m/s";

        [ObservableProperty]
    private string _aDisplay = "—";

    public double AValue { get; private set; }

    public double KotStart => ParseDouble(KotStartInput);
    public double KotEnd => ParseDouble(KotEndInput);

        public double KotAverage => (KotStart + KotEnd) / 2.0;

    public ObservableCollection<Vs30LayerViewModel> Layers { get; } = new();
    public ObservableCollection<Vs30LayerViewModel> Layers2 { get; } = new();

    public event EventHandler? SeriesChanged;

    private bool _isSyncingLayers;
    private bool _isSyncingLayers2;

    public Vs30SeriesViewModel(int seriesNumber, double defaultKotStart = 36.28, double defaultKotEnd = 36.50, string defaultT0 = "")
    {
        SeriesNumber = seriesNumber;
        _kotStartInput = defaultKotStart.ToString("F2");
        _kotEndInput = defaultKotEnd.ToString("F2");
        _t0Input = defaultT0;
        SetupLayers();
        SetupLayers2();
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

    public void SetupLayers2()
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

        var oldValues2 = Layers2.Select(l => new Tuple<double, double, double>(l.StartDepth, l.EndDepth, l.Velocity)).ToList();

        foreach (var old in Layers2)
        {
            old.LayerChanged -= OnLayer2Changed;
            old.DepthPropertyChanged -= OnLayer2DepthPropertyChanged;
        }
        Layers2.Clear();

        for (int i = 1; i <= SelectedLayerCount2; i++)
        {
            Tuple<double, double, double> def;
            if (i - 1 < oldValues2.Count)
                def = oldValues2[i - 1];
            else if (i - 1 < defaults.Length)
                def = defaults[i - 1];
            else
                def = new Tuple<double, double, double>(0.0, i * 6.0, 300.0 + (i * 100));

            var layer = new Vs30LayerViewModel(i, def.Item1, def.Item2, def.Item3);
            layer.LayerChanged += OnLayer2Changed;
            layer.DepthPropertyChanged += OnLayer2DepthPropertyChanged;
            Layers2.Add(layer);
        }

        SyncLayerChains2();
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

    private void OnLayer2Changed(object? sender, EventArgs e)
    {
        if (_isSyncingLayers2) return;
        SeriesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void SyncLayerChains2()
    {
        if (_isSyncingLayers2) return;
        _isSyncingLayers2 = true;
        try
        {
            for (int i = 0; i < Layers2.Count - 1; i++)
            {
                if (Layers2[i + 1].StartDepthInput != Layers2[i].EndDepthInput)
                    Layers2[i + 1].StartDepthInput = Layers2[i].EndDepthInput;
            }
        }
        finally
        {
            _isSyncingLayers2 = false;
        }
    }

    private void OnLayer2DepthPropertyChanged(object? sender, string propName)
    {
        if (_isSyncingLayers2) return;
        _isSyncingLayers2 = true;
        try
        {
            if (sender is Vs30LayerViewModel changedLayer)
            {
                int idx = Layers2.IndexOf(changedLayer);
                if (idx >= 0)
                {
                    if (propName == "EndDepth" && idx < Layers2.Count - 1)
                    {
                        if (Layers2[idx + 1].StartDepthInput != changedLayer.EndDepthInput)
                            Layers2[idx + 1].StartDepthInput = changedLayer.EndDepthInput;
                    }
                    else if (propName == "StartDepth" && idx > 0)
                    {
                        if (Layers2[idx - 1].EndDepthInput != changedLayer.StartDepthInput)
                            Layers2[idx - 1].EndDepthInput = changedLayer.StartDepthInput;
                    }
                }
            }
        }
        finally
        {
            _isSyncingLayers2 = false;
        }
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
                        
                        if (Layers[idx + 1].StartDepthInput != changedLayer.EndDepthInput)
                        {
                            Layers[idx + 1].StartDepthInput = changedLayer.EndDepthInput;
                        }
                    }
                    else if (propName == "StartDepth" && idx > 0)
                    {
                        
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

    partial void OnSelectedLayerCount2Changed(int value)
    {
        SetupLayers2();
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

    partial void OnXStartChanged(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnYStartChanged(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnXFinishChanged(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnYFinishChanged(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnPoissonLayer1Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnPoissonLayer2Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnShearModulusLayer1Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnShearModulusLayer2Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnBulkModulusLayer1Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnBulkModulusLayer2Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnElasticityModulusLayer1Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);
    partial void OnElasticityModulusLayer2Changed(string value) => SeriesChanged?.Invoke(this, EventArgs.Empty);

        public void UpdateResults(double df, string vs30Result)
    {
        DfDisplay = df.ToString("F2", TrCulture) + " m";
        Vs30Result = vs30Result;

        
        var digits = new string(vs30Result.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
        if (double.TryParse(digits, NumberStyles.Any, InvCulture, out var vs30Val) && vs30Val > 0)
        {
            
            AValue = 68.0 * Math.Pow(vs30Val, -0.6);
            ADisplay = AValue.ToString("F2", TrCulture);
        }
        else
        {
            AValue = 0;
            ADisplay = "—";
        }
    }

        public void UpdateLayerExclusion(double df)
    {
        for (int i = 0; i < Layers.Count; i++)
        {
            if (i == 0)
            {
                
                Layers[i].IsExcluded = df >= Layers[i].EndDepth;
            }
            else
            {
                
                Layers[i].IsExcluded = false;
            }
        }
    }

    private static double ParseDouble(string input)
        => double.TryParse(input.Replace(',', '.'), NumberStyles.Any, InvCulture, out var v) ? v : 0.0;
}
