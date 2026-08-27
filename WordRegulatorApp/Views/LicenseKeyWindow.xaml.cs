using System.Windows;
using System.Windows.Controls;
using WordRegulator.Domain.Interfaces;

namespace WordRegulatorApp.Views;

public partial class LicenseKeyWindow : Window
{
    private readonly ILicenseService _licenseService;
    public bool LicenseActivated { get; private set; }

    public LicenseKeyWindow(ILicenseService licenseService)
    {
        InitializeComponent();
        _licenseService = licenseService;
    }

    private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => DragMove();

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        LicenseActivated = false;
        Close();
    }

    private void LicenseKeyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        // Otomatik tire ekleme: WORD-XXXX-XXXX-XXXX
        var raw = LicenseKeyTextBox.Text.Replace("-", "").ToUpper();
        if (raw.Length > 16) raw = raw[..16];

        var formatted = FormatKey(raw);

        LicenseKeyTextBox.TextChanged -= LicenseKeyTextBox_TextChanged;
        LicenseKeyTextBox.Text = formatted;
        LicenseKeyTextBox.CaretIndex = formatted.Length;
        LicenseKeyTextBox.TextChanged += LicenseKeyTextBox_TextChanged;

        // Buton aktif: 19 karakter (WORD-XXXX-XXXX-XXXX) dolduğunda
        ActivateButton.IsEnabled = formatted.Length == 19;
        ErrorBorder.Visibility = Visibility.Collapsed;
    }

    private static string FormatKey(string raw)
    {
        if (raw.Length == 0) return string.Empty;
        var parts = new List<string>();
        for (int i = 0; i < raw.Length; i += 4)
            parts.Add(raw[i..Math.Min(i + 4, raw.Length)]);
        return string.Join("-", parts);
    }

    private async void ActivateButton_Click(object sender, RoutedEventArgs e)
    {
        var key = LicenseKeyTextBox.Text.Trim();

        // UI: yükleniyor
        ActivateButton.IsEnabled = false;
        LoadingOverlay.Visibility = Visibility.Visible;
        ErrorBorder.Visibility = Visibility.Collapsed;

        var result = await _licenseService.CheckLicenseAsync(key);

        LoadingOverlay.Visibility = Visibility.Collapsed;

        if (result.IsValid)
        {
            _licenseService.StoreLicenseKey(key);
            LicenseActivated = true;
            Close();
        }
        else
        {
            ErrorText.Text = result.Message;
            ErrorBorder.Visibility = Visibility.Visible;
            ActivateButton.IsEnabled = true;
        }
    }
}
