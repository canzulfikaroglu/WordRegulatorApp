using System.Windows;

namespace WordRegulator.Admin;

public partial class NewLicenseDialog : Window
{
    public string CustomerName { get; private set; } = string.Empty;
    public DateTime ExpiryDate { get; private set; } = DateTime.Today.AddYears(1);

    public NewLicenseDialog()
    {
        InitializeComponent();
        ExpiryPicker.SelectedDate = DateTime.Today.AddYears(1);
    }

    private void Create_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(CustomerNameBox.Text))
        {
            MessageBox.Show("Müşteri adı boş olamaz.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        CustomerName = CustomerNameBox.Text.Trim();
        ExpiryDate = ExpiryPicker.SelectedDate ?? DateTime.Today.AddYears(1);
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
        => DialogResult = false;
}
