namespace WordRegulator.Domain.Interfaces;

public interface IInternetConnectivityService
{
    Task<bool> IsInternetAvailableAsync(int timeoutMs = 3000);
    bool IsConnected { get; }
    event EventHandler<bool>? ConnectivityChanged;
}
