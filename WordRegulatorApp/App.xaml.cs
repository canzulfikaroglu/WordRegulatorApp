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

        
        var licenseService = ServiceProvider.GetRequiredService<ILicenseService>();
        var storedKey = licenseService.GetStoredLicenseKey();

        if (string.IsNullOrWhiteSpace(storedKey))
        {
            
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
        
        var templateService = ServiceProvider.GetRequiredService<ITemplateStorageService>();
        templateService.EnsureDirectoryExists();

        var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        
        services.AddSingleton<IVariableService, VariableService>();
        services.AddSingleton<IWordDocumentProcessor, OpenXmlWordProcessor>();
        services.AddSingleton<IExcelDocumentProcessor, OpenXmlExcelProcessor>();
        services.AddSingleton<ITemplateStorageService, WordRegulatorApp.Services.TemplateStorageService>();

        
        services.AddSingleton<ILicenseService, LicenseService>();

        
        services.AddTransient<MainViewModel>();

        
        services.AddTransient<MainWindow>();
        services.AddTransient<LicenseKeyWindow>();
    }
}
