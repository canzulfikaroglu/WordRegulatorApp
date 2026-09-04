using System.Net.Http;
using System.Net.NetworkInformation;
using WordRegulator.Domain.Interfaces;

namespace WordRegulator.Infrastructure.Services;

public class InternetConnectivityService : IInternetConnectivityService
{
    private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(3) };
    private static readonly string[] ProbeUrls =
    [
        "http://www.msftconnecttest.com/connecttest.txt",
        "http://www.google.com/generate_204"
    ];

    private bool _isConnected;
    public bool IsConnected => _isConnected;

    public event EventHandler<bool>? ConnectivityChanged;

    public InternetConnectivityService()
    {
        _isConnected = NetworkInterface.GetIsNetworkAvailable();
        NetworkChange.NetworkAvailabilityChanged += OnNetworkAvailabilityChanged;
    }

    private async void OnNetworkAvailabilityChanged(object? sender, NetworkAvailabilityEventArgs e)
    {
        bool available = e.IsAvailable;
        if (available)
        {
            available = await IsInternetAvailableAsync(2500);
        }

        UpdateStatus(available);
    }

    public async Task<bool> IsInternetAvailableAsync(int timeoutMs = 3000)
    {
        if (!NetworkInterface.GetIsNetworkAvailable())
        {
            UpdateStatus(false);
            return false;
        }

        using var cts = new CancellationTokenSource(timeoutMs);

        foreach (var url in ProbeUrls)
        {
            try
            {
                using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cts.Token);
                if (response.IsSuccessStatusCode)
                {
                    UpdateStatus(true);
                    return true;
                }
            }
            catch
            {
            }
        }

        UpdateStatus(false);
        return false;
    }

    private void UpdateStatus(bool newStatus)
    {
        if (_isConnected != newStatus)
        {
            _isConnected = newStatus;
            ConnectivityChanged?.Invoke(this, newStatus);
        }
    }
}
