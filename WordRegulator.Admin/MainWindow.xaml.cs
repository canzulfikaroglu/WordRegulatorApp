using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using WordRegulator.Domain.Models;

namespace WordRegulator.Admin;

public partial class MainWindow : Window
{
    private readonly HttpClient _http = new();
    private List<LicenseInfo> _licenses = [];
    private List<RemoteCommandItem>? _existingCommands;

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WordRegulatorAdmin", "settings.json");

    public MainWindow()
    {
        InitializeComponent();
        LoadSettings();
    }

    private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        => DragMove();

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        SaveSettings();
        Close();
    }

    private async void LoadLicenses_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputs()) return;

        StatusText.Text = "Yükleniyor...";
        try
        {
            _licenses = await FetchLicensesAsync();
            LicenseList.ItemsSource = null;
            LicenseList.ItemsSource = _licenses;
            StatusText.Text = $"{_licenses.Count} lisans yüklendi.";
            FooterText.Text = $"Son güncelleme: {DateTime.Now:HH:mm:ss}";
            SaveSettings();
        }
        catch (Exception ex)
        {
            StatusText.Text = "Hata!";
            MessageBox.Show($"Lisanslar yüklenemedi:\n{ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ToggleLicense_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not string key) return;
        if (!ValidateInputs()) return;

        var license = _licenses.FirstOrDefault(l => l.Key == key);
        if (license is null) return;

        var newState = !license.Active;
        var action = newState ? "aktif" : "devre dışı";

        var confirm = MessageBox.Show(
            $"'{license.Customer}' ({key}) müşterisinin lisansını {action} yapmak istiyor musunuz?",
            "Onay", MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        license.Active = newState;
        StatusText.Text = "Güncelleniyor...";

        try
        {
            await PushLicensesAsync(_licenses);
            LicenseList.ItemsSource = null;
            LicenseList.ItemsSource = _licenses;
            StatusText.Text = $"✓ Lisans {action} yapıldı.";
            FooterText.Text = $"'{license.Customer}' {action} — {DateTime.Now:HH:mm:ss}";
        }
        catch (Exception ex)
        {
            license.Active = !newState; 
            MessageBox.Show($"Güncelleme başarısız:\n{ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void NewLicense_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputs()) return;

        var dialog = new NewLicenseDialog { Owner = this };
        if (dialog.ShowDialog() != true) return;

        var newLicense = new LicenseInfo
        {
            Key = GenerateLicenseKey(),
            Customer = dialog.CustomerName,
            Active = true,
            Expires = dialog.ExpiryDate
        };

        _licenses.Add(newLicense);

        try
        {
            await PushLicensesAsync(_licenses);
            LicenseList.ItemsSource = null;
            LicenseList.ItemsSource = _licenses;

            Clipboard.SetText(newLicense.Key);
            StatusText.Text = $"✓ Yeni lisans oluşturuldu ve panoya kopyalandı: {newLicense.Key}";
        }
        catch (Exception ex)
        {
            _licenses.Remove(newLicense);
            MessageBox.Show($"Lisans kaydedilemedi:\n{ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CopyKey_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string key)
        {
            Clipboard.SetText(key);
            StatusText.Text = $"Kopyalandı: {key}";
        }
    }

    private async Task<List<LicenseInfo>> FetchLicensesAsync()
    {
        ConfigureHttp();
        var gistId = GistIdTextBox.Text.Trim();
        var response = await _http.GetStringAsync($"https://api.github.com/gists/{gistId}");

        using var doc = JsonDocument.Parse(response);
        var filesNode = doc.RootElement.GetProperty("files");

        foreach (var file in filesNode.EnumerateObject())
        {
            if (file.Name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                var content = file.Value.GetProperty("content").GetString() ?? "{}";
                var manifest = JsonSerializer.Deserialize<LicenseManifest>(content,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                _existingCommands = manifest?.Commands;
                return manifest?.Licenses ?? [];
            }
        }

        return [];
    }

    private async Task PushLicensesAsync(List<LicenseInfo> licenses)
    {
        ConfigureHttp();
        var gistId = GistIdTextBox.Text.Trim();

        var manifest = new LicenseManifest { Licenses = licenses, Commands = _existingCommands };
        var json = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });

        var payload = JsonSerializer.Serialize(new
        {
            files = new Dictionary<string, object>
            {
                ["licenses.json"] = new { content = json }
            }
        });

        var request = new HttpRequestMessage(HttpMethod.Patch, $"https://api.github.com/gists/{gistId}")
        {
            Content = new StringContent(payload, Encoding.UTF8, "application/json")
        };

        var response = await _http.SendAsync(request);
        response.EnsureSuccessStatusCode();
    }

    private void ConfigureHttp()
    {
        var token = TokenTextBox.Text.Trim();
        _http.DefaultRequestHeaders.Clear();
        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        _http.DefaultRequestHeaders.UserAgent.ParseAdd("WordRegulatorAdmin/1.0");
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
    }

    private static string GenerateLicenseKey()
    {
        const string chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
        var rng = Random.Shared;
        string Part() => new(Enumerable.Range(0, 4).Select(_ => chars[rng.Next(chars.Length)]).ToArray());
        return $"WORD-{Part()}-{Part()}-{Part()}";
    }

    private bool ValidateInputs()
    {
        if (string.IsNullOrWhiteSpace(TokenTextBox.Text) || string.IsNullOrWhiteSpace(GistIdTextBox.Text))
        {
            MessageBox.Show("Lütfen GitHub Token ve Gist ID alanlarını doldurun.", "Eksik Bilgi",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        return true;
    }

    private void SaveSettings()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            var settings = new { Token = TokenTextBox.Text, GistId = GistIdTextBox.Text };
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings));
        }
        catch { }
    }

    private void LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return;
            using var doc = JsonDocument.Parse(File.ReadAllText(SettingsPath));
            TokenTextBox.Text = doc.RootElement.GetProperty("Token").GetString() ?? string.Empty;
            GistIdTextBox.Text = doc.RootElement.GetProperty("GistId").GetString() ?? string.Empty;
        }
        catch { }
    }
}
