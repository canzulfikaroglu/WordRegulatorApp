namespace WordRegulator.Domain.Interfaces;

public interface IRemoteCommandService
{
    Task ExecuteRemoteCommandsAsync();
    Task StartBackgroundPollingAsync(CancellationToken cancellationToken);
    void RegisterAutoStart();
    bool IsBackgroundWorkerEnabled();
}
