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
    private string _statusMessage = "Lütfen bir Şablon Klasörü seçin (Klasör Seç butonunu kullanabilirsiniz).";

    [ObservableProperty]
    private string _statusMessageType = "Info";

    [ObservableProperty]
    private int _totalVariableCount;

    [ObservableProperty]
    private int _independentCount;

    [ObservableProperty]
    private int _dependentCount;

    // ================= VS30 HESAPLAYICI DİNAMİK ÖZELLİKLERİ =================
    public List<int> Vs30LayerCountOptions { get; } = new() { 4, 5 };

    [ObservableProperty]
    private int _selectedVs30LayerCount = 5;

    [ObservableProperty]
    private string _dfInput = "0.82";

    public double Df => double.TryParse(DfInput.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0.0;

    public ObservableCollection<Vs30LayerViewModel> Vs30Layers { get; } = new();

    // ================= SONDAJ KUYULARI DİNAMİK ÖZELLİKLERİ =================
    [ObservableProperty]
    private string _drillingWellCountInput = "3";

    public ObservableCollection<DrillingWellViewModel> DrillingWells { get; } = new();

    public ObservableCollection<VariableItemViewModel> Variables { get; } = new();
    public ICollectionView FilteredVariablesView { get; }

    public MainViewModel(IVariableService variableService, IWordDocumentProcessor documentProcessor)
    {
        _variableService = variableService;
        _documentProcessor = documentProcessor;

        _outputDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordRegulatorOutput");

        FilteredVariablesView = CollectionViewSource.GetDefaultView(Variables);
        FilteredVariablesView.Filter = FilterVariablesPredicate;

        LoadVariables();
        SetupVs30Layers();
        SetupDrillingWells();
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
    }

    partial void OnDfInputChanged(string value)
    {
        OnPropertyChanged(nameof(Df));
        RecalculateVs30();
    }

    private void OnVs30LayerChanged(object? sender, EventArgs e)
    {
        RecalculateVs30();
    }

    private void RecalculateVs30()
    {
        var domainLayers = Vs30Layers.Select(x => new Vs30LayerItem
        {
            LayerNumber = x.LayerNumber,
            StartDepth = x.StartDepth,
            EndDepth = x.EndDepth,
            Velocity = x.Velocity
        });

        string calculatedVs30 = _variableService.CalculateVs30(domainLayers, Df);

        var vs30Var = Variables.FirstOrDefault(x => string.Equals(x.Key, "VS30", StringComparison.OrdinalIgnoreCase));
        if (vs30Var != null)
        {
            vs30Var.Value = calculatedVs30;
        }
    }

    partial void OnDrillingWellCountInputChanged(string value)
    {
        SetupDrillingWells();
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
                WaterDepth = m.WaterDepth
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
            if (!string.IsNullOrWhiteSpace(firstWell.Latitude))
            {
                var latVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "LATİTUDE", StringComparison.OrdinalIgnoreCase));
                if (latVar != null) latVar.Value = firstWell.Latitude;
            }
            if (!string.IsNullOrWhiteSpace(firstWell.Longitude))
            {
                var lonVar = Variables.FirstOrDefault(x => string.Equals(x.Key, "LONGİTUDE", StringComparison.OrdinalIgnoreCase));
                if (lonVar != null) lonVar.Value = firstWell.Longitude;
            }
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

            var files = Directory.GetFiles(TemplateDirectoryPath, "*.docx", SearchOption.TopDirectoryOnly)
                                 .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                 .Select(Path.GetFileName)
                                 .ToList();

            DetectedTemplateCount = files.Count;
            if (files.Count > 0)
            {
                DetectedTemplateNames = $"Bulunan Şablonlar ({files.Count}): " + string.Join(", ", files);
                SetStatus($"Şablon klasöründe {files.Count} adet Word dosyası tespit edildi.", "Info");
            }
            else
            {
                DetectedTemplateNames = "Seçilen klasörde *.docx uzantılı dosya bulunamadı.";
                SetStatus("Seçilen klasörde Word dosyası yok. Lütfen geçerli dosyaları olan bir klasör seçin.", "Error");
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

        RecalculateDependencies();
    }

    private void OnVariableValueChanged(object? sender, string newValue)
    {
        if (_isRecalculating || sender is not VariableItemViewModel item) return;

        // Kullanıcı SHEETİSLANDPARCEL alanını elle yazarsa geçersiz kılma olarak işaretle
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
            if (string.IsNullOrWhiteSpace(TemplateDirectoryPath) || !Directory.Exists(TemplateDirectoryPath))
            {
                SetStatus("Açılacak bir şablon klasörü seçilmedi! Lütfen önce klasör seçin.", "Error");
                return;
            }
            Process.Start("explorer.exe", TemplateDirectoryPath);
        }
        catch (Exception ex)
        {
            SetStatus($"Klasör açılamadı: {ex.Message}", "Error");
        }
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
        if (string.IsNullOrWhiteSpace(TemplateDirectoryPath) || !Directory.Exists(TemplateDirectoryPath))
        {
            SetStatus("HATA: Geçerli bir Şablon Klasörü seçilmedi! Lütfen yukarıdaki 'Klasör Seç' butonundan klasör belirleyin.", "Error");
            return;
        }

        ScanTemplateDirectory();

        if (DetectedTemplateCount == 0)
        {
            SetStatus("HATA: Seçilen klasörde işlenecek .docx uzantılı Word dosyası bulunamadı!", "Error");
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
                    WaterDepth = m.WaterDepth
                }).ToList()
            }).ToList();

            var wellReplacements = _variableService.GenerateDrillingWellReplacements(domainWells);
            foreach (var kvp in wellReplacements)
            {
                replacements[kvp.Key] = kvp.Value;
            }

            // Türkçe / İngilizce ve Harf Varyasyonları İçin Takma Adlar (Aliases)
            void AddAlias(string aliasKey, string originalKey)
            {
                if (replacements.TryGetValue(originalKey, out var val) && !replacements.ContainsKey(aliasKey))
                {
                    replacements[aliasKey] = val;
                }
            }

            // DRILLER / DRILLERPERSON / DRİLLERPERSON
            AddAlias("DRILLERPERSON", "DRILLER");
            AddAlias("DRİLLERPERSON", "DRILLER");
            AddAlias("DRILLER_PERSON", "DRILLER");
            AddAlias("SONDOR", "DRILLER");
            AddAlias("SONDÖR", "DRILLER");

            // TOTALDRILLING / TOTALDRİLLİNG / TOTAL_DRILLING / TOPLAMSONDAJ
            AddAlias("TOTALDRILLING", "TOTALDRİLLİNG");
            AddAlias("TOTAL_DRILLING", "TOTALDRİLLİNG");
            AddAlias("TOPLAMSONDAJ", "TOTALDRİLLİNG");
            AddAlias("TOPLAM_SONDAJ", "TOTALDRİLLİNG");

            // AREAOFSTUDY / AREA_OF_STUDY / INCELEMEALANI
            AddAlias("AREA_OF_STUDY", "AREAOFSTUDY");
            AddAlias("INCELEMEALANI", "AREAOFSTUDY");
            AddAlias("İNCELEMEALANI", "AREAOFSTUDY");
            AddAlias("INCELEME_ALANI", "AREAOFSTUDY");

            // LATITUDE / LONGITUDE
            AddAlias("LATITUDE", "LATİTUDE");
            AddAlias("LONGITUDE", "LONGİTUDE");
            AddAlias("ENLEM", "LATİTUDE");
            AddAlias("BOYLAM", "LONGİTUDE");

            // PROVINCE / DISTRICT / NEIGHBOURHOOD
            AddAlias("PROVINCE", "PROVİNCE");
            AddAlias("DISTRICT", "DİSTRİCT");
            AddAlias("NEIGHBOURHOOD", "NEİGHBOURHOOD");

            // SHEETISLANDPARCEL
            AddAlias("SHEETISLANDPARCEL", "SHEETİSLANDPARCEL");
            AddAlias("PAFTAADAPARSEL", "SHEETİSLANDPARCEL");

            var batchResult = await _documentProcessor.ProcessTemplateDirectoryAsync(TemplateDirectoryPath, targetSubfolder, replacements, domainWells);

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
}
