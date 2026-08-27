using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using WordRegulator.Application.Interfaces;
using WordRegulator.Application.Services;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Infrastructure.Services;
using WordRegulatorApp.ViewModels;
using WordRegulatorApp.Views;

namespace WordRegulatorApp;

public partial class App : Application
{
    public static IServiceProvider ServiceProvider { get; private set; } = null!;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();
        ConfigureServices(services);
        ServiceProvider = services.BuildServiceProvider();

        // ── Lisans Kontrolü ─────────────────────────────────────────────────────
        var licenseService = ServiceProvider.GetRequiredService<ILicenseService>();
        var storedKey = licenseService.GetStoredLicenseKey();

        if (string.IsNullOrWhiteSpace(storedKey))
        {
            // İlk kurulum: Lisans anahtarı giriş ekranını göster
            var licenseWindow = ServiceProvider.GetRequiredService<LicenseKeyWindow>();
            licenseWindow.ShowDialog();

            if (!licenseWindow.LicenseActivated)
            {
                Shutdown(1);
                return;
            }
        }
        else
        {
            // Günlük kontrol: Kayıtlı anahtarı doğrula
            var result = await licenseService.CheckLicenseAsync(storedKey);

            if (!result.IsValid)
            {
                MessageBox.Show(
                    result.Message,
                    "Lisans Geçersiz",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                Shutdown(1);
                return;
            }
        }
        // ────────────────────────────────────────────────────────────────────────

        var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        // Core & Infrastructure Services
        services.AddSingleton<IVariableService, VariableService>();
        services.AddSingleton<IWordDocumentProcessor, OpenXmlWordProcessor>();

        // License Service
        services.AddSingleton<ILicenseService, LicenseService>();

        // ViewModels
        services.AddTransient<MainViewModel>();

        // Views
        services.AddTransient<MainWindow>();
        services.AddTransient<LicenseKeyWindow>();
    }
}
