using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using WordRegulator.Application.Interfaces;
using WordRegulator.Application.Services;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Infrastructure.Services;
using WordRegulatorApp.ViewModels;

namespace WordRegulatorApp;

public partial class App : Application
{
    public static IServiceProvider ServiceProvider { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();
        ConfigureServices(services);

        ServiceProvider = services.BuildServiceProvider();

        var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        // Core & Infrastructure Services
        services.AddSingleton<IVariableService, VariableService>();
        services.AddSingleton<IWordDocumentProcessor, OpenXmlWordProcessor>();

        // ViewModels
        services.AddTransient<MainViewModel>();

        // Views
        services.AddTransient<MainWindow>();
    }
}
