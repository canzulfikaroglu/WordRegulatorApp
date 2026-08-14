using System.Windows;
using WordRegulatorApp.ViewModels;

namespace WordRegulatorApp;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}