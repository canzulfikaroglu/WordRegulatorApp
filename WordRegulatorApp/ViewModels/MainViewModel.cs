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
using WordRegulator.Domain.Helpers;
using WordRegulator.Domain.Interfaces;

namespace WordRegulatorApp.ViewModels;

public partial class FormationSelectionViewModel : ObservableObject
{
    public FormationDefinition Formation { get; }
    public string Name => Formation.Name;

    [ObservableProperty]
    private bool _isSelected;

    public FormationSelectionViewModel(FormationDefinition formation)
    {
        Formation = formation;
    }
}

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

    public ObservableCollection<FormationSelectionViewModel> Formations { get; } = new();

    public void UpdateFormationVariables()
    {
        var selected = Formations.Where(f => f.IsSelected).Select(f => f.Formation).ToList();

        // FORMATİONTYPE: only names, comma-separated (for inline use in other parts of report)
        var typeVar = Variables.FirstOrDefault(v => string.Equals(v.Key, "FORMAŞİONTYPE", StringComparison.OrdinalIgnoreCase)
                                                 || string.Equals(v.Key, "FORMATİONTYPE", StringComparison.OrdinalIgnoreCase));
        if (typeVar != null)
        {
            typeVar.Value = string.Join(", ", selected.Select(f => f.Name));
            _userOverriddenKeys.Add(typeVar.Key);
        }

        // FORMATİONTEXT:
        // When 1 formation: only description
        // When >1 formations: alternating Name / Description lines per formation
        var textVar = Variables.FirstOrDefault(v => string.Equals(v.Key, "FORMAŞİONTEXT", StringComparison.OrdinalIgnoreCase)
                                                 || string.Equals(v.Key, "FORMATİONTEXT", StringComparison.OrdinalIgnoreCase));
        if (textVar != null)
        {
            if (selected.Count == 1)
            {
                textVar.Value = selected[0].Description.Trim().Replace("\r\n", "\n").Replace("\n", " ");
            }
            else
            {
                var lines = new System.Text.StringBuilder();
                for (int i = 0; i < selected.Count; i++)
                {
                    if (i > 0) lines.Append("\n\n");
                    lines.Append(selected[i].Name);
                    lines.Append('\n');
                    lines.Append(selected[i].Description.Trim().Replace("\r\n", "\n").Replace("\n", " "));
                }
                textVar.Value = lines.ToString();
            }
            _userOverriddenKeys.Add(textVar.Key);
        }
    }

    public ObservableCollection<ResidentialSuitabilityDefinition> ResidentialSuitabilities { get; } = new();

    [ObservableProperty]
    private ResidentialSuitabilityDefinition? _selectedResidentialSuitability;

    partial void OnSelectedResidentialSuitabilityChanged(ResidentialSuitabilityDefinition? value)
    {
        UpdateResidentialSuitabilityVariables();
    }

    public void UpdateResidentialSuitabilityVariables()
    {
        var resVar = Variables.FirstOrDefault(v => string.Equals(v.Key, "RESİDENTİALSUİTABİLİTY", StringComparison.OrdinalIgnoreCase)
                                               || string.Equals(v.Key, "RESIDENTIALSUITABILITY", StringComparison.OrdinalIgnoreCase));
        if (resVar != null)
        {
            string val = SelectedResidentialSuitability != null
                ? (!string.IsNullOrWhiteSpace(SelectedResidentialSuitability.Description)
                    ? SelectedResidentialSuitability.Description.Trim().Replace("\r\n", "\n")
                    : SelectedResidentialSuitability.Name)
                : string.Empty;
            resVar.Value = val;
            _userOverriddenKeys.Add(resVar.Key);
        }
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

        foreach (var f in WordRegulator.Application.Services.FormationRepository.All)
        {
            var vm = new FormationSelectionViewModel(f);
            vm.PropertyChanged += (s, e) => { if (e.PropertyName == nameof(FormationSelectionViewModel.IsSelected)) UpdateFormationVariables(); };
            Formations.Add(vm);
        }

        foreach (var rs in WordRegulator.Application.Services.ResidentialSuitabilityRepository.All)
        {
            ResidentialSuitabilities.Add(rs);
        }
        if (ResidentialSuitabilities.Count > 0)
        {
            SelectedResidentialSuitability = ResidentialSuitabilities[0];
        }

        LoadVariables();
        UpdateResidentialSuitabilityVariables();
        SetupVs30Layers();
        SetupVs30Series();
        SetupDrillingWells();
        SetupMicrotremorMeasurements();
        ScanTemplateDirectory();
        
        LoadState();
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
            string.Equals(item.Key, "FORMAŞİONTEXT", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Key, "RESİDENTİALSUİTABİLİTY", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.Key, "RESIDENTIALSUITABILITY", StringComparison.OrdinalIgnoreCase))
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

                // Otomatik Değerlendirmeler
                string pEval1 = SeismicEvaluationHelper.EvaluatePoisson(series.PoissonLayer1);
                string pEval2 = SeismicEvaluationHelper.EvaluatePoisson(series.PoissonLayer2);
                string shEval1 = SeismicEvaluationHelper.EvaluateShearModulus(series.ShearModulusLayer1);
                string shEval2 = SeismicEvaluationHelper.EvaluateShearModulus(series.ShearModulusLayer2);
                string bEval1 = SeismicEvaluationHelper.EvaluateBulkModulus(series.BulkModulusLayer1);
                string bEval2 = SeismicEvaluationHelper.EvaluateBulkModulus(series.BulkModulusLayer2);
                string eEval1 = SeismicEvaluationHelper.EvaluateElasticityModulus(series.ElasticityModulusLayer1);
                string eEval2 = SeismicEvaluationHelper.EvaluateElasticityModulus(series.ElasticityModulusLayer2);

                replacements[$"SS_{sNum}_POISSON1_EVAL"] = pEval1;
                replacements[$"SS_{sNum}_POISSON2_EVAL"] = pEval2;
                replacements[$"SS_{sNum}_POISSON1_DESC"] = pEval1;
                replacements[$"SS_{sNum}_POISSON2_DESC"] = pEval2;
                replacements[$"SS_{sNum}_POISSON1_DEG"]  = pEval1;
                replacements[$"SS_{sNum}_POISSON2_DEG"]  = pEval2;
                replacements[$"SS_{sNum}_POISSON1_DEGERLENDIRME"] = pEval1;
                replacements[$"SS_{sNum}_POISSON2_DEGERLENDIRME"] = pEval2;
                replacements[$"POİSSONRATİO1S{sNum}_EVAL"] = pEval1;
                replacements[$"POİSSONRATİO2S{sNum}_EVAL"] = pEval2;

                replacements[$"SS_{sNum}_SHEAR1_EVAL"] = shEval1;
                replacements[$"SS_{sNum}_SHEAR2_EVAL"] = shEval2;
                replacements[$"SS_{sNum}_SHEAR1_DESC"] = shEval1;
                replacements[$"SS_{sNum}_SHEAR2_DESC"] = shEval2;
                replacements[$"SS_{sNum}_SHEAR1_DEG"]  = shEval1;
                replacements[$"SS_{sNum}_SHEAR2_DEG"]  = shEval2;
                replacements[$"SS_{sNum}_SHEAR1_DEGERLENDIRME"] = shEval1;
                replacements[$"SS_{sNum}_SHEAR2_DEGERLENDIRME"] = shEval2;
                replacements[$"SHEARMODULUS1S{sNum}_EVAL"] = shEval1;
                replacements[$"SHEARMODULUS2S{sNum}_EVAL"] = shEval2;

                replacements[$"SS_{sNum}_BULK1_EVAL"] = bEval1;
                replacements[$"SS_{sNum}_BULK2_EVAL"] = bEval2;
                replacements[$"SS_{sNum}_BULK1_DESC"] = bEval1;
                replacements[$"SS_{sNum}_BULK2_DESC"] = bEval2;
                replacements[$"SS_{sNum}_BULK1_DEG"]  = bEval1;
                replacements[$"SS_{sNum}_BULK2_DEG"]  = bEval2;
                replacements[$"SS_{sNum}_BULK1_SINIF"] = bEval1;
                replacements[$"SS_{sNum}_BULK2_SINIF"] = bEval2;
                replacements[$"SS_{sNum}_BULK1_CLASS"] = bEval1;
                replacements[$"SS_{sNum}_BULK2_CLASS"] = bEval2;
                replacements[$"SS_{sNum}_BULK1_DEGERLENDIRME"] = bEval1;
                replacements[$"SS_{sNum}_BULK2_DEGERLENDIRME"] = bEval2;
                replacements[$"BULKMODULE1S{sNum}_EVAL"] = bEval1;
                replacements[$"BULKMODULE2S{sNum}_EVAL"] = bEval2;

                replacements[$"SS_{sNum}_ELASTICITY1_EVAL"] = eEval1;
                replacements[$"SS_{sNum}_ELASTICITY2_EVAL"] = eEval2;
                replacements[$"SS_{sNum}_ELASTICITY1_DESC"] = eEval1;
                replacements[$"SS_{sNum}_ELASTICITY2_DESC"] = eEval2;
                replacements[$"SS_{sNum}_ELASTICITY1_DEG"]  = eEval1;
                replacements[$"SS_{sNum}_ELASTICITY2_DEG"]  = eEval2;
                replacements[$"SS_{sNum}_ELASTICITY1_DEGERLENDIRME"] = eEval1;
                replacements[$"SS_{sNum}_ELASTICITY2_DEGERLENDIRME"] = eEval2;
                replacements[$"ELASTİCİTYMODULUS1S{sNum}_EVAL"] = eEval1;
                replacements[$"ELASTİCİTYMODULUS2S{sNum}_EVAL"] = eEval2;

                string formula = seriesReplacements.TryGetValue("VS30_FORMULA", out var f) ? f : "";
                string formulaFull = seriesReplacements.TryGetValue("VS30_FORMULA_FULL", out var ff) ? ff : "";
                double h1Val = double.TryParse(seriesReplacements.TryGetValue("VS30_H1", out var h1Str) ? h1Str.Replace(',', '.') : "0", System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsedH1) ? parsedH1 : 0.0;

                var seriesDomainLayers2 = series.Layers2.Select(x => new Vs30LayerItem
                {
                    LayerNumber = x.LayerNumber,
                    StartDepth = x.StartDepth,
                    EndDepth = x.EndDepth,
                    Velocity = x.Velocity
                }).ToList();

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
                    Layers2 = seriesDomainLayers2,
                    
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

            AddAlias("FORMATION_DETAILS", "FORMATİONTEXT");
            AddAlias("FORMASYON_DETAY", "FORMATİONTEXT");
            AddAlias("ALL_FORMATIONS", "FORMATİONTEXT");
            AddAlias("FORMATIONS_ALL", "FORMATİONTEXT");

            if (SelectedResidentialSuitability != null)
            {
                string resDesc = !string.IsNullOrWhiteSpace(SelectedResidentialSuitability.Description)
                    ? SelectedResidentialSuitability.Description.Trim().Replace("\r\n", "\n")
                    : SelectedResidentialSuitability.Name;
                string resName = SelectedResidentialSuitability.Name;

                replacements["RESİDENTİALSUİTABİLİTY"] = resDesc;
                replacements["RESIDENTIALSUITABILITY"] = resDesc;
                replacements["RESİDENTİALSUİTABİLİTY_TEXT"] = resDesc;
                replacements["RESIDENTIALSUITABILITY_TEXT"] = resDesc;
                replacements["RESİDENTİALSUİTABİLİTY_NAME"] = resName;
                replacements["RESIDENTIALSUITABILITY_NAME"] = resName;
                replacements["RESİDENTİALSUİTABİLİTY_TITLE"] = resName;
                replacements["RESIDENTIALSUITABILITY_TITLE"] = resName;
            }

            AddAlias("RESIDENTIALSUITABILITY", "RESİDENTİALSUİTABİLİTY");
            AddAlias("RESIDENTIAL_SUITABILITY", "RESİDENTİALSUİTABİLİTY");
            AddAlias("RESİDENTİAL_SUİTABİLİTY", "RESİDENTİALSUİTABİLİTY");
            AddAlias("YERLESIMEUYGUNLUK", "RESİDENTİALSUİTABİLİTY");
            AddAlias("YERLEŞİMEUYGUNLUK", "RESİDENTİALSUİTABİLİTY");
            AddAlias("YERLESIME_UYGUNLUK", "RESİDENTİALSUİTABİLİTY");
            AddAlias("YERLEŞİME_UYGUNLUK", "RESİDENTİALSUİTABİLİTY");

            
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

            var domainFormations = Formations.Where(f => f.IsSelected).Select(f => f.Formation).ToList();

            var batchResult = await _documentProcessor.ProcessTemplateDirectoryAsync(TemplateDirectoryPath, targetSubfolder, replacements, domainWells, domainVs30Series, domainFormations);

            
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
        if (ResidentialSuitabilities.Count > 0)
        {
            SelectedResidentialSuitability = ResidentialSuitabilities[0];
        }
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

    private string GetStateFilePath()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string folder = Path.Combine(appData, "WordRegulator");
        if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
        return Path.Combine(folder, "app_state.json");
    }

    public void SaveState()
    {
        try
        {
            var variablesDict = Variables.ToDictionary(v => v.Key, v => v.Value);

            // Serialize DrillingWells
            var wellsData = DrillingWells.Select(w => new
            {
                WellName = w.WellName,
                Latitude = w.Latitude,
                Longitude = w.Longitude,
                GroundKot = w.GroundKot,
                BottomKot = w.BottomKot,
                DepthInput = w.DepthInput,
                GroundwaterStatus = w.GroundwaterStatus,
                IsCollapsed = w.IsCollapsed,
                StartDate = w.StartDate,
                FinishDate = w.FinishDate,
                SelectedMeasurementCount = w.SelectedMeasurementCount,
                HasPressuremeterTest = w.HasPressuremeterTest,
                Measurements = w.Measurements.Select(m => new
                {
                    MeasurementDate = m.MeasurementDate,
                    WaterDepth = m.WaterDepth,
                    Day = m.Day
                }).ToList(),
                PressuremeterRows = w.PressuremeterRows.Select(r => new
                {
                    Depth = r.Depth,
                    LimitPressure = r.LimitPressure,
                    NetLimitPressure = r.NetLimitPressure,
                    ElasticityModulus = r.ElasticityModulus
                }).ToList()
            }).ToList();

            // Serialize Vs30 Layers
            var vs30LayersData = Vs30Layers.Select(l => new
            {
                StartDepth = l.StartDepthInput,
                EndDepth = l.EndDepthInput,
                Velocity = l.VelocityInput
            }).ToList();

            // Serialize Vs30 Series
            var vs30SeriesData = Vs30Series.Select(s => new
            {
                KotStart = s.KotStartInput,
                KotEnd = s.KotEndInput,
                T0 = s.T0Input,
                SurfaceVs30 = s.SurfaceVs30Input,
                XStart = s.XStart,
                YStart = s.YStart,
                XFinish = s.XFinish,
                YFinish = s.YFinish,
                PoissonLayer1 = s.PoissonLayer1,
                PoissonLayer2 = s.PoissonLayer2,
                ShearModulusLayer1 = s.ShearModulusLayer1,
                ShearModulusLayer2 = s.ShearModulusLayer2,
                BulkModulusLayer1 = s.BulkModulusLayer1,
                BulkModulusLayer2 = s.BulkModulusLayer2,
                ElasticityModulusLayer1 = s.ElasticityModulusLayer1,
                ElasticityModulusLayer2 = s.ElasticityModulusLayer2,
                Layers2 = s.Layers2.Select(l2 => new
                {
                    StartDepth = l2.StartDepthInput,
                    EndDepth = l2.EndDepthInput,
                    Velocity = l2.VelocityInput
                }).ToList()
            }).ToList();

            // Serialize Microtremor Measurements
            var microtremorData = MicrotremorMeasurements.Select(m => new
            {
                T0 = m.T0,
                Ta = m.Ta,
                Tb = m.Tb,
                Frequency = m.Frequency,
                Date = m.Date,
                HorizontalCoordinate = m.HorizontalCoordinate,
                VerticalCoordinate = m.VerticalCoordinate
            }).ToList();

            // Serialize selected formations
            var selectedFormations = Formations.Where(f => f.IsSelected).Select(f => f.Name).ToList();

            var state = new
            {
                Variables = variablesDict,
                BaseTopLevelInput = this.BaseTopLevelInput,
                DfInput = this.DfInput,
                DrillingWellCountInput = this.DrillingWellCountInput,
                Vs30SeriesCountInput = this.Vs30SeriesCountInput,
                MicrotremorCountInput = this.MicrotremorCountInput,
                SelectedVs30LayerCount = this.SelectedVs30LayerCount,
                OutputFolderTitle = this.OutputFolderTitle,
                OutputDirectoryPath = this.OutputDirectoryPath,
                DrillingWells = wellsData,
                Vs30Layers = vs30LayersData,
                Vs30Series = vs30SeriesData,
                MicrotremorMeasurements = microtremorData,
                SelectedFormations = selectedFormations,
                SelectedResidentialSuitability = this.SelectedResidentialSuitability?.Name
            };

            string json = System.Text.Json.JsonSerializer.Serialize(state,
                new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
            File.WriteAllText(GetStateFilePath(), json);
        }
        catch { }
    }

    public void LoadState()
    {
        try
        {
            string path = GetStateFilePath();
            if (!File.Exists(path)) return;

            string json = File.ReadAllText(path);
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            var state = doc.RootElement;

            // --- Scalar settings ---
            if (state.TryGetProperty("BaseTopLevelInput", out var b)) BaseTopLevelInput = b.GetString() ?? BaseTopLevelInput;
            if (state.TryGetProperty("DfInput", out var d)) DfInput = d.GetString() ?? DfInput;
            if (state.TryGetProperty("OutputFolderTitle", out var ot)) OutputFolderTitle = ot.GetString() ?? OutputFolderTitle;
            if (state.TryGetProperty("OutputDirectoryPath", out var odp))
            {
                var savedPath = odp.GetString();
                if (!string.IsNullOrWhiteSpace(savedPath)) OutputDirectoryPath = savedPath;
            }
            if (state.TryGetProperty("SelectedVs30LayerCount", out var sl)) SelectedVs30LayerCount = sl.GetInt32();

            // --- Counts (must come before restoring per-item data) ---
            if (state.TryGetProperty("DrillingWellCountInput", out var dc)) DrillingWellCountInput = dc.GetString() ?? DrillingWellCountInput;
            if (state.TryGetProperty("Vs30SeriesCountInput", out var vc)) Vs30SeriesCountInput = vc.GetString() ?? Vs30SeriesCountInput;
            if (state.TryGetProperty("MicrotremorCountInput", out var mc)) MicrotremorCountInput = mc.GetString() ?? MicrotremorCountInput;

            // --- Vs30 Layers ---
            if (state.TryGetProperty("Vs30Layers", out var layersEl) && layersEl.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                var layerArr = layersEl.EnumerateArray().ToList();
                for (int i = 0; i < Math.Min(layerArr.Count, Vs30Layers.Count); i++)
                {
                    var el = layerArr[i];
                    var layer = Vs30Layers[i];
                    if (el.TryGetProperty("StartDepth", out var sd)) layer.StartDepthInput = sd.GetString() ?? layer.StartDepthInput;
                    if (el.TryGetProperty("EndDepth", out var ed)) layer.EndDepthInput = ed.GetString() ?? layer.EndDepthInput;
                    if (el.TryGetProperty("Velocity", out var vel)) layer.VelocityInput = vel.GetString() ?? layer.VelocityInput;
                }
            }

            // --- Vs30 Series ---
            if (state.TryGetProperty("Vs30Series", out var seriesEl) && seriesEl.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                var seriesArr = seriesEl.EnumerateArray().ToList();
                for (int i = 0; i < Math.Min(seriesArr.Count, Vs30Series.Count); i++)
                {
                    var el = seriesArr[i];
                    var s = Vs30Series[i];
                    if (el.TryGetProperty("KotStart", out var ks)) s.KotStartInput = ks.GetString() ?? s.KotStartInput;
                    if (el.TryGetProperty("KotEnd", out var ke)) s.KotEndInput = ke.GetString() ?? s.KotEndInput;
                    if (el.TryGetProperty("T0", out var t0)) s.T0Input = t0.GetString() ?? s.T0Input;
                    if (el.TryGetProperty("SurfaceVs30", out var sv)) s.SurfaceVs30Input = sv.GetString() ?? s.SurfaceVs30Input;
                    if (el.TryGetProperty("XStart", out var xs)) s.XStart = xs.GetString() ?? s.XStart;
                    if (el.TryGetProperty("YStart", out var ys)) s.YStart = ys.GetString() ?? s.YStart;
                    if (el.TryGetProperty("XFinish", out var xf)) s.XFinish = xf.GetString() ?? s.XFinish;
                    if (el.TryGetProperty("YFinish", out var yf)) s.YFinish = yf.GetString() ?? s.YFinish;
                    if (el.TryGetProperty("PoissonLayer1", out var pl1)) s.PoissonLayer1 = pl1.GetString() ?? s.PoissonLayer1;
                    if (el.TryGetProperty("PoissonLayer2", out var pl2)) s.PoissonLayer2 = pl2.GetString() ?? s.PoissonLayer2;
                    if (el.TryGetProperty("ShearModulusLayer1", out var sh1)) s.ShearModulusLayer1 = sh1.GetString() ?? s.ShearModulusLayer1;
                    if (el.TryGetProperty("ShearModulusLayer2", out var sh2)) s.ShearModulusLayer2 = sh2.GetString() ?? s.ShearModulusLayer2;
                    if (el.TryGetProperty("BulkModulusLayer1", out var bk1)) s.BulkModulusLayer1 = bk1.GetString() ?? s.BulkModulusLayer1;
                    if (el.TryGetProperty("BulkModulusLayer2", out var bk2)) s.BulkModulusLayer2 = bk2.GetString() ?? s.BulkModulusLayer2;
                    if (el.TryGetProperty("ElasticityModulusLayer1", out var em1)) s.ElasticityModulusLayer1 = em1.GetString() ?? s.ElasticityModulusLayer1;
                    if (el.TryGetProperty("ElasticityModulusLayer2", out var em2)) s.ElasticityModulusLayer2 = em2.GetString() ?? s.ElasticityModulusLayer2;
                    // Layers2
                    if (el.TryGetProperty("Layers2", out var l2El) && l2El.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        var l2Arr = l2El.EnumerateArray().ToList();
                        for (int j = 0; j < Math.Min(l2Arr.Count, s.Layers2.Count); j++)
                        {
                            var l2 = s.Layers2[j];
                            if (l2Arr[j].TryGetProperty("StartDepth", out var l2sd)) l2.StartDepthInput = l2sd.GetString() ?? l2.StartDepthInput;
                            if (l2Arr[j].TryGetProperty("EndDepth", out var l2ed)) l2.EndDepthInput = l2ed.GetString() ?? l2.EndDepthInput;
                            if (l2Arr[j].TryGetProperty("Velocity", out var l2v)) l2.VelocityInput = l2v.GetString() ?? l2.VelocityInput;
                        }
                    }
                }
            }

            // --- Drilling Wells ---
            if (state.TryGetProperty("DrillingWells", out var wellsEl) && wellsEl.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                var wellArr = wellsEl.EnumerateArray().ToList();
                for (int i = 0; i < Math.Min(wellArr.Count, DrillingWells.Count); i++)
                {
                    var el = wellArr[i];
                    var w = DrillingWells[i];
                    if (el.TryGetProperty("WellName", out var wn)) w.WellName = wn.GetString() ?? w.WellName;
                    if (el.TryGetProperty("Latitude", out var lat)) w.Latitude = lat.GetString() ?? w.Latitude;
                    if (el.TryGetProperty("Longitude", out var lon)) w.Longitude = lon.GetString() ?? w.Longitude;
                    if (el.TryGetProperty("GroundKot", out var gk)) w.GroundKot = gk.GetString() ?? w.GroundKot;
                    if (el.TryGetProperty("BottomKot", out var bk)) w.BottomKot = bk.GetString() ?? w.BottomKot;
                    if (el.TryGetProperty("DepthInput", out var dep)) w.DepthInput = dep.GetString() ?? w.DepthInput;
                    if (el.TryGetProperty("GroundwaterStatus", out var gws)) w.GroundwaterStatus = gws.GetString() ?? w.GroundwaterStatus;
                    if (el.TryGetProperty("IsCollapsed", out var ic)) w.IsCollapsed = ic.GetBoolean();
                    if (el.TryGetProperty("StartDate", out var sd2)) w.StartDate = sd2.GetString() ?? w.StartDate;
                    if (el.TryGetProperty("FinishDate", out var fd)) w.FinishDate = fd.GetString() ?? w.FinishDate;
                    if (el.TryGetProperty("SelectedMeasurementCount", out var smc)) w.SelectedMeasurementCount = smc.GetInt32();
                    if (el.TryGetProperty("HasPressuremeterTest", out var hpt)) w.HasPressuremeterTest = hpt.GetBoolean();
                    // Measurements
                    if (el.TryGetProperty("Measurements", out var measEl) && measEl.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        var measArr = measEl.EnumerateArray().ToList();
                        for (int j = 0; j < Math.Min(measArr.Count, w.Measurements.Count); j++)
                        {
                            var m = w.Measurements[j];
                            if (measArr[j].TryGetProperty("MeasurementDate", out var md)) m.MeasurementDate = md.GetString() ?? m.MeasurementDate;
                            if (measArr[j].TryGetProperty("WaterDepth", out var wd)) m.WaterDepth = wd.GetString() ?? m.WaterDepth;
                            if (measArr[j].TryGetProperty("Day", out var dy)) m.Day = dy.GetString() ?? m.Day;
                        }
                    }
                    // Pressurometer Rows
                    if (el.TryGetProperty("PressuremeterRows", out var pmEl) && pmEl.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        var pmArr = pmEl.EnumerateArray().ToList();
                        // Add missing rows
                        while (w.PressuremeterRows.Count < pmArr.Count)
                            w.AddPressuremeterRowCommand.Execute(null);
                        for (int j = 0; j < Math.Min(pmArr.Count, w.PressuremeterRows.Count); j++)
                        {
                            var r = w.PressuremeterRows[j];
                            if (pmArr[j].TryGetProperty("Depth", out var pmd)) r.Depth = pmd.GetString() ?? r.Depth;
                            if (pmArr[j].TryGetProperty("LimitPressure", out var pml)) r.LimitPressure = pml.GetString() ?? r.LimitPressure;
                            if (pmArr[j].TryGetProperty("NetLimitPressure", out var pmn)) r.NetLimitPressure = pmn.GetString() ?? r.NetLimitPressure;
                            if (pmArr[j].TryGetProperty("ElasticityModulus", out var pme)) r.ElasticityModulus = pme.GetString() ?? r.ElasticityModulus;
                        }
                    }
                }
            }

            // --- Microtremor Measurements ---
            if (state.TryGetProperty("MicrotremorMeasurements", out var mtEl) && mtEl.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                var mtArr = mtEl.EnumerateArray().ToList();
                for (int i = 0; i < Math.Min(mtArr.Count, MicrotremorMeasurements.Count); i++)
                {
                    var el = mtArr[i];
                    var m = MicrotremorMeasurements[i];
                    if (el.TryGetProperty("T0", out var t0)) m.T0 = t0.GetString() ?? m.T0;
                    if (el.TryGetProperty("Ta", out var ta)) m.Ta = ta.GetString() ?? m.Ta;
                    if (el.TryGetProperty("Tb", out var tb)) m.Tb = tb.GetString() ?? m.Tb;
                    if (el.TryGetProperty("Frequency", out var freq)) m.Frequency = freq.GetString() ?? m.Frequency;
                    if (el.TryGetProperty("Date", out var dt)) m.Date = dt.GetString() ?? m.Date;
                    if (el.TryGetProperty("HorizontalCoordinate", out var hc)) m.HorizontalCoordinate = hc.GetString() ?? m.HorizontalCoordinate;
                    if (el.TryGetProperty("VerticalCoordinate", out var vc2)) m.VerticalCoordinate = vc2.GetString() ?? m.VerticalCoordinate;
                }
            }

            // --- Formation Selections ---
            if (state.TryGetProperty("SelectedFormations", out var sfEl) && sfEl.ValueKind == System.Text.Json.JsonValueKind.Array)
            {
                var selectedNames = sfEl.EnumerateArray()
                    .Select(e => e.GetString() ?? string.Empty)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var f in Formations)
                {
                    f.IsSelected = selectedNames.Contains(f.Name);
                }
            }

            // --- Residential Suitability Selection ---
            if (state.TryGetProperty("SelectedResidentialSuitability", out var rsEl))
            {
                var rsName = rsEl.GetString();
                if (!string.IsNullOrEmpty(rsName))
                {
                    var found = ResidentialSuitabilities.FirstOrDefault(r => string.Equals(r.Name, rsName, StringComparison.OrdinalIgnoreCase));
                    if (found != null)
                    {
                        SelectedResidentialSuitability = found;
                    }
                }
            }

            // --- Variables (overrides) ---
            if (state.TryGetProperty("Variables", out var dictEl))
            {
                foreach (var prop in dictEl.EnumerateObject())
                {
                    var v = Variables.FirstOrDefault(x => x.Key == prop.Name);
                    if (v != null)
                    {
                        v.Value = prop.Value.GetString() ?? "";
                        _userOverriddenKeys.Add(prop.Name);
                    }
                }
            }

            RecalculateDependencies();
            RecalculateAllSeries();
            RecalculateDrillingWellsSummary();
        }
        catch { }
    }
}
