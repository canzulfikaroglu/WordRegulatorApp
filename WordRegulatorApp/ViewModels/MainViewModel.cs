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
    private string _statusMessageType = "Info"; // Info, Success, Error

    [ObservableProperty]
    private int _totalVariableCount;

    [ObservableProperty]
    private int _independentCount;

    [ObservableProperty]
    private int _dependentCount;

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
        ScanTemplateDirectory();
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
                if (item.IsDependent && !_userOverriddenKeys.Contains(item.Key))
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
            // Kullanıcının girdiği klasör başlığını temizle (Geçersiz karakterleri engelle)
            string folderTitle = string.IsNullOrWhiteSpace(OutputFolderTitle) ? "Rapor_Çıktısı" : OutputFolderTitle.Trim();
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                folderTitle = folderTitle.Replace(invalidChar, '_');
            }

            // Çıktıların toplanacağı özel alt klasör yolu
            string targetSubfolder = Path.Combine(OutputDirectoryPath, folderTitle);

            var replacements = Variables.ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);

            var batchResult = await _documentProcessor.ProcessTemplateDirectoryAsync(TemplateDirectoryPath, targetSubfolder, replacements);

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
        SetStatus("Tüm değişkenler varsayılan değerlerine sıfırlandı.", "Info");
    }

    private void SetStatus(string message, string type)
    {
        StatusMessage = message;
        StatusMessageType = type;
    }
}
