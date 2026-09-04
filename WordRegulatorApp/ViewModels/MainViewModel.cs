using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using WordRegulator.Application.Interfaces;
using WordRegulator.Domain.Entities;
using WordRegulator.Domain.Enums;
using WordRegulator.Domain.Interfaces;

namespace WordRegulatorApp.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private static readonly CultureInfo TrCulture = new CultureInfo("tr-TR");

    private readonly IVariableService _variableService;
    private readonly IWordDocumentProcessor _documentProcessor;
    private readonly HashSet<string> _userOverriddenKeys = new(StringComparer.OrdinalIgnoreCase);
    private bool _isRecalculating;

    [ObservableProperty]
    private string _templateDirectoryPath = string.Empty;

    [ObservableProperty]
    private string _outputDirectoryPath = string.Empty;

    [ObservableProperty]
    private string _outputFolderTitle = "Proje_Çıktısı";

    [ObservableProperty]
    private int _detectedTemplateCount;

    [ObservableProperty]
    private string _detectedTemplateNames = "Şablon klasörü seçilmedi.";

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private bool _isProcessing;

    [ObservableProperty]
    private string _statusMessage = "Hazır. Şablonlar otomatik yüklendi, çıktı klasörünü belirleyip başlayabilirsiniz.";

    [ObservableProperty]
    private string _statusMessageType = "Info";

    [ObservableProperty]
    private bool _isInternetConnected = true;

    [ObservableProperty]
    private bool _isOffline;

    [ObservableProperty]
    private int _totalVariableCount;

    [ObservableProperty]
    private int _independentCount;

    [ObservableProperty]
    private int _dependentCount;

    
    public List<int> Vs30LayerCountOptions { get; } = new() { 4, 5 };

    [ObservableProperty]
    private int _selectedVs30LayerCount = 5;

    [ObservableProperty]
    private string _vs30SeriesCountInput = "1";

    [ObservableProperty]
    private string _baseTopLevelInput = "-3,24(35,57)";

    
    
    [ObservableProperty]
    private string _dfInput = "0.82";

    public double Df => double.TryParse(DfInput.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0.0;

    public ObservableCollection<Vs30LayerViewModel> Vs30Layers { get; } = new();
    public ObservableCollection<Vs30SeriesViewModel> Vs30Series { get; } = new();

    
    [ObservableProperty]
    private string _drillingWellCountInput = "3";

    public ObservableCollection<DrillingWellViewModel> DrillingWells { get; } = new();

    
    [ObservableProperty]
    private string _microtremorCountInput = "1";

        public int MicrotremorCount => MicrotremorMeasurements.Count;

    public ObservableCollection<MicrotremorMeasurementViewModel> MicrotremorMeasurements { get; } = new();

    
    public IReadOnlyList<WordRegulator.Domain.Entities.FormationDefinition> FormationList { get; }
        = WordRegulator.Application.Services.FormationRepository.All;

    [ObservableProperty]
    private WordRegulator.Domain.Entities.FormationDefinition? _selectedFormation;

    partial void OnSelectedFormationChanged(WordRegulator.Domain.Entities.FormationDefinition? value)
    {
        if (value == null) return;

        
        var typeVar = Variables.FirstOrDefault(v => string.Equals(v.Key, "FORMAŞİONTYPE", StringComparison.OrdinalIgnoreCase)
                                                 || string.Equals(v.Key, "FORMATİONTYPE", StringComparison.OrdinalIgnoreCase));
        if (typeVar != null) { typeVar.Value = value.Name; _userOverriddenKeys.Add(typeVar.Key); }

        
        var textVar = Variables.FirstOrDefault(v => string.Equals(v.Key, "FORMAŞİONTEXT", StringComparison.OrdinalIgnoreCase)
                                                 || string.Equals(v.Key, "FORMATİONTEXT", StringComparison.OrdinalIgnoreCase));
        if (textVar != null) { textVar.Value = value.Description; _userOverriddenKeys.Add(textVar.Key); }
    }

    public ObservableCollection<VariableItemViewModel> Variables { get; } = new();
    public ICollectionView FilteredVariablesView { get; }

    private readonly IExcelDocumentProcessor? _excelProcessor;
    private readonly ITemplateStorageService? _templateStorageService;
    private readonly IInternetConnectivityService? _connectivityService;

    public MainViewModel(
        IVariableService variableService, 
        IWordDocumentProcessor documentProcessor, 
        IExcelDocumentProcessor? excelProcessor = null,
        ITemplateStorageService? templateStorageService = null,
        IInternetConnectivityService? connectivityService = null)
    {
        _variableService = variableService;
        _documentProcessor = documentProcessor;
        _excelProcessor = excelProcessor;
        _templateStorageService = templateStorageService;
        _connectivityService = connectivityService;

        if (_connectivityService != null)
        {
            _isInternetConnected = _connectivityService.IsConnected;
            _isOffline = !_isInternetConnected;
            _connectivityService.ConnectivityChanged += (s, connected) =>
            {
                System.Windows.Application.Current?.Dispatcher?.Invoke(() =>
                {
                    IsInternetConnected = connected;
                    IsOffline = !connected;
                    if (!connected)
                    {
                        SetStatus("DİKKAT: İnternet bağlantısı kesildi! Rapor oluşturmak için internet gereklidir.", "Error");
                    }
                    else
                    {
                        SetStatus("İnternet bağlantısı sağlandı.", "Info");
                    }
                });
            };
        }

        if (_templateStorageService != null)
        {
            _templateDirectoryPath = _templateStorageService.TemplatesDirectoryPath;
        }
        else
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _templateDirectoryPath = Path.Combine(appData, "WordRegulator", "Templates");
        }

        _outputDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordRegulatorOutput");

        FilteredVariablesView = CollectionViewSource.GetDefaultView(Variables);
        FilteredVariablesView.Filter = FilterVariablesPredicate;

        LoadVariables();
        SetupVs30Layers();
        SetupVs30Series();
        SetupDrillingWells();
        SetupMicrotremorMeasurements();
        ScanTemplateDirectory();
    }

    private void SetupVs30Layers()
    {
        Vs30Layers.Clear();

        var defaults = new Tuple<double, double, double>[]
        {
            new(0.00, 1.10, 157.0),
            new(1.10, 7.50, 1008.0),
            new(7.50, 13.50, 1152.0),
            new(13.50, 21.00, 1236.0),
            new(21.00, 30.00, 1254.0)
        };

        for (int i = 1; i <= SelectedVs30LayerCount; i++)
        {
            var def = (i <= defaults.Length) ? defaults[i - 1] : new Tuple<double, double, double>(0.0, i * 6.0, 300.0 + (i * 100));
            var layer = new Vs30LayerViewModel(i, def.Item1, def.Item2, def.Item3);
            layer.LayerChanged += OnVs30LayerChanged;
            Vs30Layers.Add(layer);
        }

        RecalculateVs30();
    }

    partial void OnSelectedVs30LayerCountChanged(int value)
    {
        SetupVs30Layers();
        RecalculateAllSeries();
    }

    partial void OnVs30SeriesCountInputChanged(string value)
    {
        SetupVs30Series();
    }

    partial void OnDfInputChanged(string value)
    {
        OnPropertyChanged(nameof(Df));
        RecalculateVs30();
    }

    private void OnVs30LayerChanged(object? sender, EventArgs e)
    {
        RecalculateVs30();
        RecalculateAllSeries();
    }

    private void OnVs30SeriesChanged(object? sender, EventArgs e)
    {
        RecalculateAllSeries();
    }

        private void RecalculateVs30()
    {
        RecalculateAllSeries();
    }

    private void SetupVs30Series()
    {
        if (!int.TryParse(Vs30SeriesCountInput, out int count) || count < 1)
        {
            count = 1;
        }
        if (count > 50) count = 50;

        while (Vs30Series.Count < count)
        {
            int n = Vs30Series.Count + 1;
            double kotStart = 36.28 + (n - 1) * 5.0;
            double kotEnd = kotStart + 0.22;
            var s = new Vs30SeriesViewModel(n, kotStart, kotEnd);
            s.SeriesChanged += OnVs30SeriesChanged;
            Vs30Series.Add(s);
        }

        while (Vs30Series.Count > count)
        {
            var last = Vs30Series.Last();
            last.SeriesChanged -= OnVs30SeriesChanged;
            Vs30Series.RemoveAt(Vs30Series.Count - 1);
        }

        RecalculateAllSeries();
    }

    partial void OnBaseTopLevelInputChanged(string value)
    {
        var btVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "FROMTHEBASETOPLEVEL", StringComparison.OrdinalIgnoreCase));
        if (btVar != null && btVar.Value != value)
        {
            btVar.Value = value;
        }
        RecalculateAllSeries();
    }

        private double GetBaseTopLevel(out string rawValue)
    {
        rawValue = BaseTopLevelInput;
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            var btVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "FROMTHEBASETOPLEVEL", StringComparison.OrdinalIgnoreCase));
            if (btVar != null) rawValue = btVar.Value;
        }

        
        string val = rawValue ?? string.Empty;
        int parenOpen = val.IndexOf('(');
        int parenClose = val.IndexOf(')');
        if (parenOpen >= 0 && parenClose > parenOpen)
        {
            val = val.Substring(parenOpen + 1, parenClose - parenOpen - 1);
        }
        val = val.Replace(',', '.');
        return double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0.0;
    }

    [ObservableProperty]
    private string _allVs30Values = "—";

    private void RecalculateAllSeries()
    {
        if (Vs30Series.Count == 0) return;

        double baseTopLevel = GetBaseTopLevel(out _);

        var vs30NumbersList = new List<string>();

        foreach (var series in Vs30Series)
        {
            double kotAvg = series.KotAverage;
            double df = kotAvg - baseTopLevel;
            if (df < 0) df = 0;

            var seriesDomainLayers = series.Layers.Select(x => new Vs30LayerItem
            {
                LayerNumber = x.LayerNumber,
                StartDepth = x.StartDepth,
                EndDepth = x.EndDepth,
                Velocity = x.Velocity
            }).ToList();

            string vs30Str = _variableService.CalculateVs30(seriesDomainLayers, df);
            series.UpdateResults(df, vs30Str);

            
            series.UpdateLayerExclusion(df);

            
            var digits = new string(vs30Str.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
            if (double.TryParse(digits, NumberStyles.Any, CultureInfo.InvariantCulture, out var vs30Val) && vs30Val > 0)
            {
                vs30NumbersList.Add(Math.Round(vs30Val).ToString(CultureInfo.GetCultureInfo("tr-TR")));
            }
        }

        
        AllVs30Values = vs30NumbersList.Count > 0
            ? string.Join(" / ", vs30NumbersList)
            : "—";
    }

    partial void OnDrillingWellCountInputChanged(string value)
    {
        SetupDrillingWells();
    }

    partial void OnMicrotremorCountInputChanged(string value)
    {
        SetupMicrotremorMeasurements();
    }

    private void SetupMicrotremorMeasurements()
    {
        if (!int.TryParse(MicrotremorCountInput, out int count) || count < 1)
            count = 1;
        if (count > 50) count = 50;

        while (MicrotremorMeasurements.Count < count)
        {
            int nextIdx = MicrotremorMeasurements.Count + 1;
            MicrotremorMeasurements.Add(new MicrotremorMeasurementViewModel(nextIdx));
        }

        while (MicrotremorMeasurements.Count > count)
        {
            MicrotremorMeasurements.RemoveAt(MicrotremorMeasurements.Count - 1);
        }

        OnPropertyChanged(nameof(MicrotremorCount));

        
        var countVar = Variables.FirstOrDefault(x =>
            string.Equals(x.Key, "MICROTREMOR_COUNT", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(x.Key, "MİKROTREMOR_SAYISI", StringComparison.OrdinalIgnoreCase));
        if (countVar != null)
            countVar.Value = count.ToString();
    }

    private void SetupDrillingWells()
    {
        if (!int.TryParse(DrillingWellCountInput, out int count) || count < 1)
        {
            count = 1;
        }
        if (count > 30) count = 30;

        while (DrillingWells.Count < count)
        {
            int nextIdx = DrillingWells.Count + 1;
            var well = new DrillingWellViewModel($"SK-{nextIdx}", defaultDepth: 15.0);
            well.WellChanged += OnDrillingWellChanged;
            DrillingWells.Add(well);
        }

        while (DrillingWells.Count > count)
        {
            var last = DrillingWells.Last();
            last.WellChanged -= OnDrillingWellChanged;
            DrillingWells.RemoveAt(DrillingWells.Count - 1);
        }

        RecalculateDrillingWellsSummary();

        var countVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "DRİLLİNGWELLCOUNT", StringComparison.OrdinalIgnoreCase));
        if (countVar != null)
        {
            countVar.Value = count.ToString();
        }
    }

    private void OnDrillingWellChanged(object? sender, EventArgs e)
    {
        RecalculateDrillingWellsSummary();
    }

    private void RecalculateDrillingWellsSummary()
    {
        var domainWells = DrillingWells.Select(w => new DrillingWellItem
        {
            WellName = w.WellName,
            Latitude = w.Latitude,
            Longitude = w.Longitude,
            GroundKot = w.GroundKot,
            BottomKot = w.BottomKot,
            DepthMeters = w.DepthMeters,
            GroundwaterStatus = w.GroundwaterStatus,
            IsCollapsed = w.IsCollapsed,
            StartDate = w.StartDate,
            FinishDate = w.FinishDate,
            Measurements = w.Measurements.Select(m => new GroundwaterMeasurementItem
            {
                MeasurementDate = m.MeasurementDate,
                WaterDepth = m.WaterDepth,
                Day = m.Day
            }).ToList(),
            HasPressuremeterTest = w.HasPressuremeterTest,
            PressuremeterRows = w.PressuremeterRows.Select(r => new WordRegulator.Domain.Entities.PressuremeterRowItem
            {
                Depth = r.Depth,
                LimitPressure = r.LimitPressure,
                NetLimitPressure = r.NetLimitPressure,
                ElasticityModulus = r.ElasticityModulus
            }).ToList()
        });

        var (totalDrilling, groundwaterLevel) = _variableService.SummarizeDrillingWells(domainWells);

        var totalVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "TOTALDRİLLİNG", StringComparison.OrdinalIgnoreCase));
        if (totalVar != null)
        {
            totalVar.Value = totalDrilling;
        }

        var gwVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "GROUNDWATERLEVEL", StringComparison.OrdinalIgnoreCase));
        if (gwVar != null)
        {
            gwVar.Value = groundwaterLevel;
        }

        var firstWell = DrillingWells.FirstOrDefault();
        if (firstWell != null)
        {
            
            
            if (!string.IsNullOrWhiteSpace(firstWell.StartDate))
            {
                var dateVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "DRİLLİNGDATE", StringComparison.OrdinalIgnoreCase));
                if (dateVar != null)
                {
                    if (string.Equals(firstWell.StartDate, firstWell.FinishDate, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(firstWell.FinishDate))
                    {
                        dateVar.Value = firstWell.StartDate;
                    }
                    else
                    {
                        dateVar.Value = $"{firstWell.StartDate} - {firstWell.FinishDate}";
                    }
                }
            }
        }
    }

    public void ScanTemplateDirectory()
    {
        try
        {
            if (string.IsNullOrWhiteSpace(TemplateDirectoryPath) || !Directory.Exists(TemplateDirectoryPath))
            {
                DetectedTemplateCount = 0;
                DetectedTemplateNames = "Klasör seçilmedi. Lütfen Klasör Seç butonunu kullanın.";
                return;
            }

            var docxFiles = Directory.GetFiles(TemplateDirectoryPath, "*.docx", SearchOption.TopDirectoryOnly)
                                     .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                     .Select(Path.GetFileName)
                                     .ToList();

            var xlsxFiles = Directory.GetFiles(TemplateDirectoryPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                                     .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                     .Select(Path.GetFileName)
                                     .ToList();

            var allFiles = docxFiles.Concat(xlsxFiles).ToList();
            DetectedTemplateCount = allFiles.Count;

            if (allFiles.Count > 0)
            {
                var parts = new List<string>();
                if (docxFiles.Count > 0) parts.Add($"{docxFiles.Count} Word");
                if (xlsxFiles.Count > 0) parts.Add($"{xlsxFiles.Count} Excel");
                DetectedTemplateNames = $"Bulunan Şablonlar ({allFiles.Count}): " + string.Join(", ", allFiles);
                SetStatus($"%AppData% şablon klasöründe {string.Join(" ve ", parts)} dosyası hazır.", "Info");
            }
            else
            {
                DetectedTemplateNames = "⚠️ Şablon bulunamadı! 'Klasörü Aç' butonuna basarak şablon dosyalarınızı ekleyin.";
                SetStatus("Şablon bulunamadı! 'Klasörü Aç' butonuna basarak şablon dosyalarınızı bu klasöre yerleştirin.", "Warning");
            }
        }
        catch (Exception ex)
        {
            DetectedTemplateCount = 0;
            DetectedTemplateNames = $"Klasör taranırken hata: {ex.Message}";
        }
    }

    private void LoadVariables()
    {
        Variables.Clear();
        var definitions = _variableService.GetDefaultDefinitions();

        foreach (var def in definitions)
        {
            var item = new VariableItemViewModel(def);
            item.ValueChanged += OnVariableValueChanged;
            Variables.Add(item);
        }

        TotalVariableCount = Variables.Count;
        IndependentCount = Variables.Count(x => !x.IsDependent);
        DependentCount = Variables.Count(x => x.IsDependent);

        var btVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "FROMTHEBASETOPLEVEL", StringComparison.OrdinalIgnoreCase));
        if (btVar != null && !string.IsNullOrWhiteSpace(btVar.Value))
        {
            BaseTopLevelInput = btVar.Value;
        }

        RecalculateDependencies();
    }

    private void OnVariableValueChanged(object? sender, string newValue)
    {
        if (_isRecalculating || sender is not VariableItemViewModel item) return;

        if (string.Equals(item.Key, "FROMTHEBASETOPLEVEL", StringComparison.OrdinalIgnoreCase))
        {
            if (BaseTopLevelInput != newValue)
            {
                BaseTopLevelInput = newValue;
            }
        }

        
        if (string.Equals(item.Key, "SHEETİSLANDPARCEL", StringComparison.OrdinalIgnoreCase))
        {
            _userOverriddenKeys.Add(item.Key);
            item.IsManualOverride = true;
        }

        if (item.IsDependent)
        {
            _userOverriddenKeys.Add(item.Key);
            item.IsManualOverride = true;
        }

        RecalculateDependencies();
    }

    private void RecalculateDependencies()
    {
        _isRecalculating = true;

        try
        {
            var currentValues = Variables.ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
            var updated = _variableService.RecalculateDependentVariables(currentValues, _userOverriddenKeys);

            foreach (var item in Variables)
            {
                bool isCalculatedField = item.IsDependent || string.Equals(item.Key, "SHEETİSLANDPARCEL", StringComparison.OrdinalIgnoreCase);

                if (isCalculatedField && !_userOverriddenKeys.Contains(item.Key))
                {
                    if (updated.TryGetValue(item.Key, out var newVal) && item.Value != newVal)
                    {
                        item.Value = newVal;
                    }
                }
            }
        }
        finally
        {
            _isRecalculating = false;
        }
    }

    partial void OnSearchTextChanged(string value)
    {
        FilteredVariablesView.Refresh();
    }

    private bool FilterVariablesPredicate(object obj)
    {
        if (obj is not VariableItemViewModel item) return false;

        
        if (string.Equals(item.Key, "FORMATİONTYPE", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Key, "FORMAŞİONTYPE", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Key, "FORMATİONTEXT", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Key, "FORMAŞİONTEXT", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(SearchText))
        {
            var search = SearchText.Trim();
            bool matchKey = ContainsTurkishIgnoreCase(item.Key, search);
            bool matchName = ContainsTurkishIgnoreCase(item.DisplayName, search);
            bool matchHint = ContainsTurkishIgnoreCase(item.Hint, search);

            return matchKey || matchName || matchHint;
        }

        return true;
    }

    private static bool ContainsTurkishIgnoreCase(string source, string search)
    {
        if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(search)) return false;

        string sNorm = source.Replace('İ', 'i').Replace('I', 'ı').ToLower(TrCulture);
        string qNorm = search.Replace('İ', 'i').Replace('I', 'ı').ToLower(TrCulture);

        return sNorm.Contains(qNorm);
    }

    [RelayCommand]
    private void BrowseTemplateDirectory()
    {
        var initialDir = !string.IsNullOrWhiteSpace(TemplateDirectoryPath) && Directory.Exists(TemplateDirectoryPath)
            ? TemplateDirectoryPath
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        var dialog = new OpenFolderDialog
        {
            Title = "Şablon Klasörünü Seçin",
            InitialDirectory = initialDir
        };

        if (dialog.ShowDialog() == true)
        {
            TemplateDirectoryPath = dialog.FolderName;
            ScanTemplateDirectory();
        }
    }

    [RelayCommand]
    private void OpenTemplateFolder()
    {
        try
        {
            if (_templateStorageService != null)
            {
                _templateStorageService.OpenTemplatesFolder();
            }
            else if (!string.IsNullOrWhiteSpace(TemplateDirectoryPath) && Directory.Exists(TemplateDirectoryPath))
            {
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{TemplateDirectoryPath}\"") { UseShellExecute = true });
            }
            else
            {
                SetStatus("Şablon klasörü bulunamadı!", "Error");
            }
        }
        catch (Exception ex)
        {
            SetStatus($"Klasör açılamadı: {ex.Message}", "Error");
        }
    }

    [RelayCommand]
    private void RefreshTemplates()
    {
        ScanTemplateDirectory();
        if (DetectedTemplateCount > 0)
            SetStatus($"Şablonlar güncellendi. {DetectedTemplateCount} adet şablon hazır.", "Success");
        else
            SetStatus("Şablon bulunamadı. Lütfen 'Klasörü Aç' ile dosyalarınızı yerleştirin.", "Warning");
    }

    [RelayCommand]
    private void BrowseOutputDirectory()
    {
        var initialDir = !string.IsNullOrWhiteSpace(OutputDirectoryPath) && Directory.Exists(OutputDirectoryPath)
            ? OutputDirectoryPath
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        var openFolderDialog = new OpenFolderDialog
        {
            Title = "Çıktı Ana Klasörünü Seçin",
            InitialDirectory = initialDir
        };

        if (openFolderDialog.ShowDialog() == true)
        {
            OutputDirectoryPath = openFolderDialog.FolderName;
            SetStatus($"Çıktı klasörü güncellendi: {OutputDirectoryPath}", "Info");
        }
    }

    [RelayCommand]
    private async Task GenerateDocumentAsync()
    {
        if (_connectivityService != null && !await _connectivityService.IsInternetAvailableAsync(2500))
        {
            IsInternetConnected = false;
            IsOffline = true;
            SetStatus("HATA: İnternet bağlantısı bulunamadı! Rapor oluşturmak için aktif bir internet bağlantısı gereklidir.", "Error");
            System.Windows.MessageBox.Show(
                "İnternet bağlantısı bulunamadı!\nWord ve Excel raporlarını oluşturabilmek için aktif bir internet bağlantısı gereklidir.",
                "İnternet Bağlantısı Gerekli",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(TemplateDirectoryPath) || !Directory.Exists(TemplateDirectoryPath))
        {
            SetStatus("HATA: Şablon klasörü bulunamadı! Lütfen 'Klasörü Aç' butonuna basarak şablonlarınızı (%AppData%\\WordRegulator\\Templates) içine ekleyin.", "Error");
            return;
        }

        ScanTemplateDirectory();

        if (DetectedTemplateCount == 0)
        {
            SetStatus("HATA: Şablon klasöründe (*.docx veya *.xlsx) dosyası bulunamadı! Lütfen 'Klasörü Aç' butonu ile şablon dosyalarınızı yerleştirin.", "Error");
            return;
        }

        if (string.IsNullOrWhiteSpace(OutputDirectoryPath))
        {
            SetStatus("HATA: Çıktı klasörü belirlenmedi!", "Error");
            return;
        }

        IsProcessing = true;
        SetStatus($"Toplam {DetectedTemplateCount} adet Word dosyası taranıyor ve {{değişkenler}} değiştiriliyor...", "Info");

        try
        {
            string folderTitle = string.IsNullOrWhiteSpace(OutputFolderTitle) ? "Rapor_Çıktısı" : OutputFolderTitle.Trim();
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                folderTitle = folderTitle.Replace(invalidChar, '_');
            }

            string targetSubfolder = Path.Combine(OutputDirectoryPath, folderTitle);

            var replacements = Variables.ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);

            var domainWells = DrillingWells.Select(w => new DrillingWellItem
            {
                WellName = w.WellName,
                Latitude = w.Latitude,
                Longitude = w.Longitude,
                GroundKot = w.GroundKot,
                BottomKot = w.BottomKot,
                DepthMeters = w.DepthMeters,
                GroundwaterStatus = w.GroundwaterStatus,
                IsCollapsed = w.IsCollapsed,
                StartDate = w.StartDate,
                FinishDate = w.FinishDate,
                Measurements = w.Measurements.Select(m => new GroundwaterMeasurementItem
                {
                    MeasurementDate = m.MeasurementDate,
                    WaterDepth = m.WaterDepth,
                    Day = m.Day
                }).ToList(),
                HasPressuremeterTest = w.HasPressuremeterTest,
                PressuremeterRows = w.PressuremeterRows.Select(r => new WordRegulator.Domain.Entities.PressuremeterRowItem
                {
                    Depth = r.Depth,
                    LimitPressure = r.LimitPressure,
                    NetLimitPressure = r.NetLimitPressure,
                    ElasticityModulus = r.ElasticityModulus
                }).ToList()
            }).ToList();

            
            double baseTopLevelForWells = GetBaseTopLevel(out _);
            var wellReplacements = _variableService.GenerateDrillingWellReplacements(domainWells, baseTopLevelForWells);
            foreach (var kvp in wellReplacements)
            {
                replacements[kvp.Key] = kvp.Value;
            }

            
            double baseTopLevelForReplace = GetBaseTopLevel(out _);

            if (Vs30Series.Count > 0)
            {
                var firstSeries = Vs30Series[0];
                var firstSeriesDomainLayers = firstSeries.Layers.Select(x => new Vs30LayerItem
                {
                    LayerNumber = x.LayerNumber,
                    StartDepth = x.StartDepth,
                    EndDepth = x.EndDepth,
                    Velocity = x.Velocity
                }).ToList();

                double dfForReplace = firstSeries.KotAverage - baseTopLevelForReplace;
                if (dfForReplace < 0) dfForReplace = 0;

                var vs30Replacements = _variableService.GenerateVs30Replacements(firstSeriesDomainLayers, dfForReplace);
                foreach (var kvp in vs30Replacements)
                {
                    replacements[kvp.Key] = kvp.Value;
                }
            }

            var domainVs30Series = new List<Vs30SeriesItem>();

            
            for (int si = 0; si < Vs30Series.Count; si++)
            {
                var series = Vs30Series[si];
                int sNum = si + 1;
                double seriesDf = series.KotAverage - baseTopLevelForReplace;
                if (seriesDf < 0) seriesDf = 0;

                var seriesDomainLayers = series.Layers.Select(x => new Vs30LayerItem
                {
                    LayerNumber = x.LayerNumber,
                    StartDepth = x.StartDepth,
                    EndDepth = x.EndDepth,
                    Velocity = x.Velocity
                }).ToList();

                var seriesReplacements = _variableService.GenerateVs30Replacements(seriesDomainLayers, seriesDf);
                foreach (var kvp in seriesReplacements)
                {
                    
                    string serimKey = kvp.Key.Replace("VS30_", $"VS30_S{sNum}_");
                    replacements[serimKey] = kvp.Value;
                }

                
                replacements[$"VS30_S{sNum}_KOTSTART"] = series.KotStartInput;
                replacements[$"VS30_S{sNum}_KOTEND"] = series.KotEndInput;
                replacements[$"VS30_S{sNum}_KOTAVG"] = series.KotAverage.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
                replacements[$"VS30_S{sNum}_DF"] = seriesDf.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
                replacements[$"VS30_S{sNum}_RESULT"] = series.Vs30Result;

                
                double aVal = series.AValue;
                if (aVal <= 0)
                {
                    var digits = new string(series.Vs30Result.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
                    if (double.TryParse(digits, NumberStyles.Any, CultureInfo.InvariantCulture, out var vs30Val) && vs30Val > 0)
                    {
                        aVal = 68.0 * Math.Pow(vs30Val, -0.6);
                    }
                }
                string aStr = aVal.ToString("F2", System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));
                replacements[$"A{sNum}"] = aStr;
                replacements[$"VS30_S{sNum}_A"] = aStr;
                replacements[$"A_{sNum}"] = aStr;

                
                replacements[$"S{sNum}T0"] = series.T0Input;
                replacements[$"VS30_S{sNum}_T0"] = series.T0Input;
                replacements[$"T0_{sNum}"] = series.T0Input;

                
                replacements[$"SS_{sNum}_NAME"] = $"SS-{sNum}";
                replacements[$"SERIM_{sNum}_NAME"] = $"{sNum}. Serim";
                replacements[$"SERIM_{sNum}_NO"] = sNum.ToString();

                
                replacements[$"SS_{sNum}_XSTART"]  = series.XStart;
                replacements[$"SS_{sNum}_YSTART"]  = series.YStart;
                replacements[$"SS_{sNum}_XFINISH"] = series.XFinish;
                replacements[$"SS_{sNum}_YFINISH"] = series.YFinish;
                
                replacements[$"SEİSMİCCOORDİNATESS-{sNum}XSTART"]  = series.XStart;
                replacements[$"SEİSMİCCOORDİNATESS-{sNum}YSTART"]  = series.YStart;
                replacements[$"SEİSMİCCOORDİNATESS-{sNum}XFİNİSH"] = series.XFinish;
                replacements[$"SEİSMİCCOORDİNATESS-{sNum}YFİNİSH"] = series.YFinish;

                
                replacements[$"SS_{sNum}_POISSON1"] = series.PoissonLayer1;
                replacements[$"SS_{sNum}_POISSON2"] = series.PoissonLayer2;
                replacements[$"POİSSONRATİO1S{sNum}"] = series.PoissonLayer1;
                replacements[$"POİSSONRATİO2S{sNum}"] = series.PoissonLayer2;

                
                replacements[$"SS_{sNum}_SHEAR1"] = series.ShearModulusLayer1;
                replacements[$"SS_{sNum}_SHEAR2"] = series.ShearModulusLayer2;
                replacements[$"SHEARMODULUS1S{sNum}"] = series.ShearModulusLayer1;
                replacements[$"SHEARMODULUS2S{sNum}"] = series.ShearModulusLayer2;

                
                replacements[$"SS_{sNum}_BULK1"] = series.BulkModulusLayer1;
                replacements[$"SS_{sNum}_BULK2"] = series.BulkModulusLayer2;
                replacements[$"BULKMODULE1S{sNum}"] = series.BulkModulusLayer1;
                replacements[$"BULKMODULE2S{sNum}"] = series.BulkModulusLayer2;

                
                replacements[$"SS_{sNum}_ELASTICITY1"] = series.ElasticityModulusLayer1;
                replacements[$"SS_{sNum}_ELASTICITY2"] = series.ElasticityModulusLayer2;
                replacements[$"ELASTİCİTYMODULUS1S{sNum}"] = series.ElasticityModulusLayer1;
                replacements[$"ELASTİCİTYMODULUS2S{sNum}"] = series.ElasticityModulusLayer2;

                string formula = seriesReplacements.TryGetValue("VS30_FORMULA", out var f) ? f : "";
                string formulaFull = seriesReplacements.TryGetValue("VS30_FORMULA_FULL", out var ff) ? ff : "";
                double h1Val = double.TryParse(seriesReplacements.TryGetValue("VS30_H1", out var h1Str) ? h1Str.Replace(',', '.') : "0", System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsedH1) ? parsedH1 : 0.0;

                domainVs30Series.Add(new Vs30SeriesItem
                {
                    SeriesNumber = sNum,
                    SeriesName = $"SS-{sNum}",
                    KotStart = series.KotStart,
                    KotEnd = series.KotEnd,
                    KotAverage = series.KotAverage,
                    Df = seriesDf,
                    H1 = h1Val,
                    AValue = aVal,
                    T0 = series.T0Input,
                    Vs30Result = series.Vs30Result,
                    SurfaceVs30 = series.SurfaceVs30Input,
                    Formula = formula,
                    FormulaFull = formulaFull,
                    Layers = seriesDomainLayers,
                    
                    XStart  = series.XStart,
                    YStart  = series.YStart,
                    XFinish = series.XFinish,
                    YFinish = series.YFinish,
                    PoissonLayer1          = series.PoissonLayer1,
                    PoissonLayer2          = series.PoissonLayer2,
                    ShearModulusLayer1     = series.ShearModulusLayer1,
                    ShearModulusLayer2     = series.ShearModulusLayer2,
                    BulkModulusLayer1      = series.BulkModulusLayer1,
                    BulkModulusLayer2      = series.BulkModulusLayer2,
                    ElasticityModulusLayer1 = series.ElasticityModulusLayer1,
                    ElasticityModulusLayer2 = series.ElasticityModulusLayer2,
                });
            }

            
            if (Vs30Series.Count > 0)
            {
                replacements["A"] = replacements.TryGetValue("A1", out var a1) ? a1 : "1,00";
                replacements["T0"] = Vs30Series[0].T0Input;
            }

            
            
            var t0Parts = new List<string>();
            var t0DashParts = new List<string>();
            foreach (var s in Vs30Series)
            {
                string t0Val = string.IsNullOrWhiteSpace(s.T0Input) ? "—" : s.T0Input.Trim();
                t0Parts.Add($"{s.SeriesNumber}. Serim T0 = {t0Val} sn");
                t0DashParts.Add(t0Val);
            }
            string t0AllText = string.Join("\n", t0Parts);
            replacements["T0_ALL_SERIES"] = t0AllText;    
            replacements["ALL_T0"] = t0AllText;            
            replacements["T0_OZET"] = t0AllText;           

            
            
            string t0DashText = string.Join(" - ", t0DashParts);
            replacements["T0_DASH"] = t0DashText;          
            replacements["ALL_T0_DASH"] = t0DashText;      
            replacements["T0_TUM"] = t0DashText;           
            replacements["T0_LISTE"] = t0DashText;         

            
            var vs30NumbersList = new List<string>();
            var vs30NumericValues = new List<double>();
            foreach (var s in Vs30Series)
            {
                var digits = new string(s.Vs30Result.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray()).Replace(',', '.');
                if (double.TryParse(digits, NumberStyles.Any, CultureInfo.InvariantCulture, out var vs30Val) && vs30Val > 0)
                {
                    vs30NumbersList.Add(Math.Round(vs30Val).ToString(CultureInfo.GetCultureInfo("tr-TR")));
                    vs30NumericValues.Add(vs30Val);
                }
            }

            string combinedVs30Nums = string.Join(" / ", vs30NumbersList);
            string combinedVs30Msn = vs30NumbersList.Count > 0 ? $"{combinedVs30Nums} m/sn" : "—";

            replacements["VS30_ALL"] = combinedVs30Msn;
            replacements["VS30_SUMMARY"] = combinedVs30Msn;
            replacements["ALL_VS30"] = combinedVs30Msn;
            replacements["ALL_VS30_RESULTS"] = combinedVs30Msn;
            replacements["TUK_VS30"] = combinedVs30Msn;
            replacements["TÜK_VS30"] = combinedVs30Msn;

            
            replacements["VS30_ALL_VALUES"] = combinedVs30Nums;
            replacements["VS30_ALL_NUM"] = combinedVs30Nums;
            replacements["ALL_VS30_NUM"] = combinedVs30Nums;
            replacements["TUK_VS30_NUM"] = combinedVs30Nums;
            replacements["TÜK_VS30_NUM"] = combinedVs30Nums;

            
            string maxMinVs30Str;
            if (vs30NumericValues.Count >= 2)
            {
                double maxVal = vs30NumericValues.Max();
                double minVal = vs30NumericValues.Min();
                maxMinVs30Str = $"{Math.Round(maxVal).ToString(CultureInfo.GetCultureInfo("tr-TR"))} - {Math.Round(minVal).ToString(CultureInfo.GetCultureInfo("tr-TR"))}";
            }
            else if (vs30NumericValues.Count == 1)
            {
                maxMinVs30Str = Math.Round(vs30NumericValues[0]).ToString(CultureInfo.GetCultureInfo("tr-TR"));
            }
            else
            {
                maxMinVs30Str = "—";
            }

            replacements["MAXMİNVS30"] = maxMinVs30Str;
            replacements["MAXMINVS30"] = maxMinVs30Str;
            replacements["MAX_MIN_VS30"] = maxMinVs30Str;
            replacements["VS30_MAX_MIN"] = maxMinVs30Str;

            
            var surfaceVs30Parts = new List<string>();
            foreach (var s in Vs30Series)
            {
                int sn = s.SeriesNumber;
                string surfVal = string.IsNullOrWhiteSpace(s.SurfaceVs30Input) ? "—" : s.SurfaceVs30Input.Trim();
                
                string surfWithUnit = surfVal == "—" ? "—"
                    : (surfVal.EndsWith("m/sn", StringComparison.OrdinalIgnoreCase) || surfVal.EndsWith("m/s", StringComparison.OrdinalIgnoreCase)
                       ? surfVal : $"{surfVal} m/sn");

                replacements[$"VS30_S{sn}_YUZEY"] = surfWithUnit;   
                replacements[$"SURFACE_VS30_{sn}"] = surfWithUnit;   
                replacements[$"VS30_YUZEY_{sn}"] = surfWithUnit;     

                if (surfVal != "—") surfaceVs30Parts.Add(surfVal);
            }

            
            string combinedSurface = surfaceVs30Parts.Count > 0
                ? string.Join(" / ", surfaceVs30Parts) + " m/sn"
                : "—";
            replacements["VS30_ALL_SURFACE"] = combinedSurface;   
            replacements["ALL_SURFACE_VS30"] = combinedSurface;   
            replacements["VS30_YUZEY_ALL"] = combinedSurface;     

            
            void AddAlias(string aliasKey, string originalKey)
            {
                if (replacements.TryGetValue(originalKey, out var val) && !replacements.ContainsKey(aliasKey))
                {
                    replacements[aliasKey] = val;
                }
            }

            
            AddAlias("DRILLERPERSON", "DRILLER");
            AddAlias("DRİLLERPERSON", "DRILLER");
            AddAlias("DRILLER_PERSON", "DRILLER");
            AddAlias("SONDOR", "DRILLER");
            AddAlias("SONDÖR", "DRILLER");

            
            AddAlias("TOTALDRILLING", "TOTALDRİLLİNG");
            AddAlias("TOTAL_DRILLING", "TOTALDRİLLİNG");
            AddAlias("TOPLAMSONDAJ", "TOTALDRİLLİNG");
            AddAlias("TOPLAM_SONDAJ", "TOTALDRİLLİNG");

            
            AddAlias("YERALTISUYU", "GROUNDWATER");
            AddAlias("YERALTI_SUYU", "GROUNDWATER");
            AddAlias("GROUND_WATER", "GROUNDWATER");
            AddAlias("GROUNDWATERLEVEL", "GROUNDWATER");
            AddAlias("YERALTISUYUSEVIYESI", "GROUNDWATER");

            
            AddAlias("AREA_OF_STUDY", "AREAOFSTUDY");
            AddAlias("INCELEMEALANI", "AREAOFSTUDY");
            AddAlias("İNCELEMEALANI", "AREAOFSTUDY");
            AddAlias("INCELEME_ALANI", "AREAOFSTUDY");

            

            
            AddAlias("LATITUDE", "LATİTUDE");
            AddAlias("LONGITUDE", "LONGİTUDE");
            AddAlias("ENLEM", "LATİTUDE");
            AddAlias("BOYLAM", "LONGİTUDE");

            
            AddAlias("PROVINCE", "PROVİNCE");
            AddAlias("DISTRICT", "DİSTRİCT");
            AddAlias("NEIGHBOURHOOD", "NEİGHBOURHOOD");

            
            AddAlias("SHEETISLANDPARCEL", "SHEETİSLANDPARCEL");
            AddAlias("PAFTAADAPARSEL", "SHEETİSLANDPARCEL");

            
            int mtCount = MicrotremorMeasurements.Count;
            replacements["MICROTREMOR_COUNT"]  = mtCount.ToString();  
            replacements["MİKROTREMOR_SAYISI"] = mtCount.ToString();  
            replacements["MT_COUNT"]           = mtCount.ToString();  

            
            var microtremorDomainItems = new List<MicrotremorMeasurementItem>();
            foreach (var mt in MicrotremorMeasurements)
            {
                int n = mt.MeasurementNo;
                string t0mt  = mt.T0;
                string tamt  = mt.Ta;
                string tbmt  = mt.Tb;
                string freqmt = mt.Frequency;
                string datemt = mt.Date;
                string xmt   = mt.HorizontalCoordinate;
                string ymt   = mt.VerticalCoordinate;

                
                replacements[$"MT_{n}_NO"]    = $"MT-{n}";
                replacements[$"MT_{n}_ID"]    = $"MT-{n}";
                replacements[$"MT_{n}_NUM"]   = n.ToString();
                replacements[$"MT_{n}_T0"]    = t0mt;
                replacements[$"MT_{n}_TA"]    = tamt;
                replacements[$"MT_{n}_TB"]    = tbmt;
                replacements[$"MT_{n}_FREQ"]  = freqmt;
                replacements[$"MT_{n}_FREKANS"] = freqmt;
                replacements[$"MT_{n}_DATE"]  = datemt;
                replacements[$"MT_{n}_TARIH"] = datemt;
                replacements[$"MT_{n}_X"]     = xmt;
                replacements[$"MT_{n}_YATAY"] = xmt;
                replacements[$"MT_{n}_Y"]     = ymt;
                replacements[$"MT_{n}_DUZEY"] = ymt;

                microtremorDomainItems.Add(new MicrotremorMeasurementItem
                {
                    MeasurementNo         = n,
                    T0                    = t0mt,
                    Ta                    = tamt,
                    Tb                    = tbmt,
                    Frequency             = freqmt,
                    Date                  = datemt,
                    HorizontalCoordinate  = xmt,
                    VerticalCoordinate    = ymt
                });
            }

            var batchResult = await _documentProcessor.ProcessTemplateDirectoryAsync(TemplateDirectoryPath, targetSubfolder, replacements, domainWells, domainVs30Series);

            
            if (_excelProcessor != null)
            {
                var excelBatchResult = await _excelProcessor.ProcessExcelTemplatesForWellsAsync(
                    TemplateDirectoryPath, targetSubfolder, replacements, domainWells);

                
                batchResult.ProcessedFileCount += excelBatchResult.ProcessedFileCount;
                batchResult.TotalReplacementsCount += excelBatchResult.TotalReplacementsCount;
                batchResult.OutputFiles.AddRange(excelBatchResult.OutputFiles);
                batchResult.Errors.AddRange(excelBatchResult.Errors);
            }

            if (batchResult.Success)
            {
                SetStatus($"BAŞARILI! Toplam {batchResult.ProcessedFileCount} dosya işlendi, {batchResult.TotalReplacementsCount} adet {{değişken}} değiştirildi. Çıktı Klasörü: {targetSubfolder}", "Success");
            }
            else
            {
                var errorText = string.Join("; ", batchResult.Errors);
                SetStatus($"HATA: {errorText}", "Error");
            }
        }
        catch (Exception ex)
        {
            SetStatus($"Beklenmeyen hata: {ex.Message}", "Error");
        }
        finally
        {
            IsProcessing = false;
        }
    }

    [RelayCommand]
    private void ResetAll()
    {
        _userOverriddenKeys.Clear();
        foreach (var item in Variables)
        {
            item.ResetToDefaultCommand.Execute(null);
        }
        RecalculateDependencies();
        RecalculateVs30();
        RecalculateDrillingWellsSummary();
        SetStatus("Tüm değişkenler varsayılan değerlerine sıfırlandı.", "Info");
    }

    private void SetStatus(string message, string type)
    {
        StatusMessage = message;
        StatusMessageType = type;
    }

        private static double ParseDouble(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return 0.0;
        return double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0;
    }
}
