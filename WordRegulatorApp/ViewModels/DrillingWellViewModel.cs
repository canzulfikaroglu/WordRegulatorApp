using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace WordRegulatorApp.ViewModels;

public partial class DrillingWellViewModel : ObservableObject
{
    public string WellName { get; set; } = string.Empty;

    [ObservableProperty]
    private string _latitude = string.Empty;

    [ObservableProperty]
    private string _longitude = string.Empty;

    [ObservableProperty]
    private string _groundKot = string.Empty;

    [ObservableProperty]
    private string _bottomKot = string.Empty;

    [ObservableProperty]
    private string _depthInput = "15.00";

    [ObservableProperty]
    private string _groundwaterStatus = "Sontaj Kuyusu Açıklaması ";

    [ObservableProperty]
    private bool _isCollapsed;

    [ObservableProperty]
    private string _startDate = DateTime.Now.ToString("dd.MM.yyyy");

    [ObservableProperty]
    private string _finishDate = DateTime.Now.ToString("dd.MM.yyyy");

    [ObservableProperty]
    private int _selectedMeasurementCount = 1;

    
    [ObservableProperty]
    private bool _hasPressuremeterTest;

    public ObservableCollection<PressuremeterRowViewModel> PressuremeterRows { get; } = new();

    public int[] MeasurementCountOptions { get; } = new[] { 1, 2, 3, 4, 5 };

    public ObservableCollection<GroundwaterMeasurementViewModel> Measurements { get; } = new();

    public double DepthMeters => double.TryParse(DepthInput.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0.0;

    public event EventHandler? WellChanged;

    public DrillingWellViewModel(string wellName, double defaultDepth = 15.0)
    {
        WellName = wellName;
        _depthInput = defaultDepth.ToString("F2");
        SyncMeasurements();
    }

    partial void OnSelectedMeasurementCountChanged(int value)
    {
        SyncMeasurements();
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnIsCollapsedChanged(bool value)
    {
        if (value)
        {
            _selectedMeasurementCount = 1;
            OnPropertyChanged(nameof(SelectedMeasurementCount));
        }
        SyncMeasurements();
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnHasPressuremeterTestChanged(bool value)
    {
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnStartDateChanged(string value)
    {
        RecalculateMeasurementDates();
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RecalculateMeasurementDates()
    {
        if (!DateTime.TryParseExact(StartDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var baseDt))
        {
            if (!DateTime.TryParse(StartDate, out baseDt))
            {
                baseDt = DateTime.Now;
            }
        }

        DateTime current = baseDt.AddDays(1);
        int stepDays = 2;

        for (int i = 0; i < Measurements.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(Measurements[i].Day))
            {
                Measurements[i].Day = GetDefaultDayString(i + 1);
            }

            if (i == 0)
            {
                Measurements[i].MeasurementDate = current.ToString("dd.MM.yyyy");
            }
            else
            {
                current = current.AddDays(stepDays);
                Measurements[i].MeasurementDate = current.ToString("dd.MM.yyyy");
                stepDays += 2;
            }
        }
    }

    public void SyncMeasurements()
    {
        int targetCount = IsCollapsed ? 1 : SelectedMeasurementCount;
        if (targetCount < 1) targetCount = 1;

        if (!DateTime.TryParseExact(StartDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var baseDt))
        {
            if (!DateTime.TryParse(StartDate, out baseDt))
            {
                baseDt = DateTime.Now;
            }
        }

        while (Measurements.Count < targetCount)
        {
            int idx = Measurements.Count + 1;
            string dateStr;
            string dayStr = GetDefaultDayString(idx);
            if (idx == 1)
            {
                dateStr = baseDt.AddDays(1).ToString("dd.MM.yyyy");
            }
            else if (idx == 2)
            {
                var dt1 = baseDt.AddDays(1);
                dateStr = dt1.AddDays(2).ToString("dd.MM.yyyy");
            }
            else
            {
                var dt2 = baseDt.AddDays(1).AddDays(2);
                dateStr = dt2.AddDays(4 * (idx - 2)).ToString("dd.MM.yyyy");
            }

            string depth = IsCollapsed ? "Kuyu çökmüş" : "2,70";
            var m = new GroundwaterMeasurementViewModel(idx, dateStr, depth, dayStr, !IsCollapsed);
            m.MeasurementChanged += OnMeasurementItemChanged;
            Measurements.Add(m);
        }

        while (Measurements.Count > targetCount)
        {
            var last = Measurements.Last();
            last.MeasurementChanged -= OnMeasurementItemChanged;
            Measurements.RemoveAt(Measurements.Count - 1);
        }

        if (Measurements.Count > 0)
        {
            if (IsCollapsed)
            {
                Measurements[0].WaterDepth = "Kuyu çökmüş";
                Measurements[0].IsEditable = false;
            }
            else
            {
                Measurements[0].IsEditable = true;
                if (Measurements[0].WaterDepth == "Kuyu çökmüş")
                {
                    Measurements[0].WaterDepth = "2,70";
                }
            }
        }
    }

    private void OnMeasurementItemChanged(object? sender, EventArgs e)
    {
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    

    [RelayCommand]
    private void AddPressuremeterRow()
    {
        int newIdx = PressuremeterRows.Count + 1;
        var row = new PressuremeterRowViewModel(newIdx);
        row.RowChanged += OnPressuremeterRowChanged;
        PressuremeterRows.Add(row);
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    [RelayCommand]
    private void RemovePressuremeterRow()
    {
        if (PressuremeterRows.Count == 0) return;
        var last = PressuremeterRows.Last();
        last.RowChanged -= OnPressuremeterRowChanged;
        PressuremeterRows.RemoveAt(PressuremeterRows.Count - 1);
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    private void OnPressuremeterRowChanged(object? sender, EventArgs e)
    {
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnDepthInputChanged(string value)
    {
        OnPropertyChanged(nameof(DepthMeters));
        WellChanged?.Invoke(this, EventArgs.Empty);
    }

    partial void OnGroundwaterStatusChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnLatitudeChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnLongitudeChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnGroundKotChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnBottomKotChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);
    partial void OnFinishDateChanged(string value) => WellChanged?.Invoke(this, EventArgs.Empty);

    public static string GetDefaultDayString(int idx)
    {
        if (idx == 1) return "1. gün";
        if (idx == 2) return "3. gün";
        if (idx == 3) return "7. gün";
        int days = 7 + (idx - 3) * 4;
        return $"{days}. gün";
    }
}
