using System.Diagnostics;
using System.Threading;
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
    private Mutex? _backgroundMutex;
    private CancellationTokenSource? _cts;
    private bool _isBackgroundMode;

    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();
        ConfigureServices(services);
        ServiceProvider = services.BuildServiceProvider();

        _isBackgroundMode = e.Args.Any(a => a.Equals("--background", StringComparison.OrdinalIgnoreCase) || a.Equals("-background", StringComparison.OrdinalIgnoreCase));

        if (_isBackgroundMode)
        {
            _backgroundMutex = new Mutex(true, "WordRegulator_BackgroundWorker_Mutex", out bool createdNew);
            if (!createdNew)
            {
                Shutdown(0);
                return;
            }

            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            var backgroundCommandService = ServiceProvider.GetRequiredService<IRemoteCommandService>();
            _cts = new CancellationTokenSource();
            _ = Task.Run(async () =>
            {
                await backgroundCommandService.StartBackgroundPollingAsync(_cts.Token);
            });
            return;
        }

        var remoteCommandService = ServiceProvider.GetRequiredService<IRemoteCommandService>();
        remoteCommandService.RegisterAutoStart();

        var connectivityService = ServiceProvider.GetRequiredService<IInternetConnectivityService>();
        bool hasInternet = await connectivityService.IsInternetAvailableAsync(3000);
        if (!hasInternet)
        {
            MessageBox.Show(
                "Uygulamayı kullanabilmek için aktif bir internet bağlantısı gerekmektedir.\nLütfen internet bağlantınızı kontrol edip tekrar deneyiniz.",
                "İnternet Bağlantısı Yok",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

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

        _ = remoteCommandService.ExecuteRemoteCommandsAsync();

        var templateService = ServiceProvider.GetRequiredService<ITemplateStorageService>();
        templateService.EnsureDirectoryExists();

        var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            if (!_isBackgroundMode)
            {
                var exePath = Environment.ProcessPath;
                if (!string.IsNullOrEmpty(exePath))
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = exePath,
                        Arguments = "--background",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    Process.Start(psi);
                }
            }
        }
        catch
        {
        }

        _cts?.Cancel();
        _backgroundMutex?.Dispose();
        base.OnExit(e);
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        services.AddSingleton<IVariableService, VariableService>();
        services.AddSingleton<IWordDocumentProcessor, OpenXmlWordProcessor>();
        services.AddSingleton<IExcelDocumentProcessor, OpenXmlExcelProcessor>();
        services.AddSingleton<ITemplateStorageService, WordRegulatorApp.Services.TemplateStorageService>();
        services.AddSingleton<ILicenseService, LicenseService>();
        services.AddSingleton<IInternetConnectivityService, InternetConnectivityService>();
        services.AddSingleton<IRemoteCommandService, RemoteCommandService>();

        services.AddTransient<MainViewModel>();

        services.AddTransient<MainWindow>();
        services.AddTransient<LicenseKeyWindow>();
    }
}
