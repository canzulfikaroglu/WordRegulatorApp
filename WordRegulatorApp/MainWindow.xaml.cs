using System.ComponentModel;
using System.Windows;
using WordRegulatorApp.ViewModels;

namespace WordRegulatorApp;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = _viewModel;
        this.Closing += MainWindow_Closing;
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        _viewModel.SaveState();
    }
}
