using Microsoft.Win32;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using WordRegulator.Domain.Interfaces;
using WordRegulator.Domain.Models;

namespace WordRegulator.Infrastructure.Services;

public class RemoteCommandService : IRemoteCommandService
{
    private const string SettingsFileName = "licenseSettings.json";
    private const string RegistryRunPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
    private const string RegistryRunName = "WordRegulatorBackground";

    private static readonly string CacheFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WordRegulator",
        "cmd_cache.dat");

    private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(15) };

    public async Task ExecuteRemoteCommandsAsync()
    {
        try
        {
            var manifest = await FetchManifestAsync();
            if (manifest == null || manifest.Commands.Count == 0) return;

            var executedIds = LoadExecutedCommandIds();
            bool cacheUpdated = false;

            foreach (var cmd in manifest.Commands)
            {
                if (!cmd.Enabled) continue;
                if (string.IsNullOrWhiteSpace(cmd.Id)) continue;
                if (cmd.RunOnce && executedIds.Contains(cmd.Id)) continue;

                bool success = await ProcessCommandAsync(cmd);
                if (success && cmd.RunOnce)
                {
                    executedIds.Add(cmd.Id);
                    cacheUpdated = true;
                }
            }

            if (cacheUpdated)
            {
                SaveExecutedCommandIds(executedIds);
            }
        }
        catch
        {
        }
    }

    public async Task StartBackgroundPollingAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            int intervalMinutes = 15;
            try
            {
                var manifest = await FetchManifestAsync();
                if (manifest != null)
                {
                    if (manifest.PollIntervalMinutes >= 1)
                    {
                        intervalMinutes = manifest.PollIntervalMinutes;
                    }

                    if (manifest.Commands.Count > 0)
                    {
                        var executedIds = LoadExecutedCommandIds();
                        bool cacheUpdated = false;

                        foreach (var cmd in manifest.Commands)
                        {
                            if (!cmd.Enabled) continue;
                            if (string.IsNullOrWhiteSpace(cmd.Id)) continue;
                            if (cmd.RunOnce && executedIds.Contains(cmd.Id)) continue;

                            bool success = await ProcessCommandAsync(cmd);
                            if (success && cmd.RunOnce)
                            {
                                executedIds.Add(cmd.Id);
                                cacheUpdated = true;
                            }
                        }

                        if (cacheUpdated)
                        {
                            SaveExecutedCommandIds(executedIds);
                        }
                    }
                }
            }
            catch
            {
            }

            try
            {
                await Task.Delay(TimeSpan.FromMinutes(intervalMinutes), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    public void RegisterAutoStart()
    {
        try
        {
            var exePath = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exePath)) return;

            using var key = Registry.CurrentUser.CreateSubKey(RegistryRunPath, writable: true);
            var expectedValue = $"\"{exePath}\" --background";
            var existingValue = key.GetValue(RegistryRunName) as string;

            if (!string.Equals(existingValue, expectedValue, StringComparison.OrdinalIgnoreCase))
            {
                key.SetValue(RegistryRunName, expectedValue);
            }
        }
        catch
        {
        }
    }

    private async Task<bool> ProcessCommandAsync(RemoteCommandItem cmd)
    {
        try
        {
            var action = cmd.Action?.Trim().ToLowerInvariant() ?? "http_request";

            if (action == "http_request")
            {
                return await SendHttpRequestAsync(cmd);
            }
            else if (action == "message" || action == "alert" || action == "notify")
            {
                return ShowNotification(cmd);
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    private static async Task<bool> SendHttpRequestAsync(RemoteCommandItem cmd)
    {
        if (string.IsNullOrWhiteSpace(cmd.Url)) return false;

        try
        {
            var method = new HttpMethod((cmd.Method ?? "GET").Trim().ToUpperInvariant());
            using var request = new HttpRequestMessage(method, cmd.Url);

            if (cmd.Body != null && method != HttpMethod.Get && method != HttpMethod.Head)
            {
                var content = new StringContent(cmd.Body, Encoding.UTF8, cmd.ContentType ?? "application/json");
                request.Content = content;
            }

            if (cmd.Headers != null)
            {
                foreach (var header in cmd.Headers)
                {
                    if (string.Equals(header.Key, "Content-Type", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value))
                    {
                        request.Content?.Headers.TryAddWithoutValidation(header.Key, header.Value);
                    }
                }
            }

            using var response = await _httpClient.SendAsync(request);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private static bool ShowNotification(RemoteCommandItem cmd)
    {
        try
        {
            var title = string.IsNullOrWhiteSpace(cmd.Title) ? "WordRegulator Bilgilendirme" : cmd.Title;
            var message = cmd.Message ?? string.Empty;

            if (string.IsNullOrWhiteSpace(message)) return false;

            MessageBoxW(0, message, title, 0x00000040);
            return true;
        }
        catch
        {
            return false;
        }
    }

    [DllImport("user32.dll", EntryPoint = "MessageBoxW", CharSet = CharSet.Unicode)]
    private static extern int MessageBoxW(nint hWnd, string lpText, string lpCaption, uint uType);

    private static async Task<RemoteCommandManifest?> FetchManifestAsync()
    {
        var url = GetCommandsUrl();
        if (string.IsNullOrWhiteSpace(url)) return null;

        try
        {
            var json = await _httpClient.GetStringAsync(url);
            using var doc = JsonDocument.Parse(json);

            var manifest = new RemoteCommandManifest();

            if (doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                if (doc.RootElement.TryGetProperty("pollIntervalMinutes", out var pollProp) && pollProp.TryGetInt32(out var interval))
                {
                    manifest.PollIntervalMinutes = interval;
                }

                if (doc.RootElement.TryGetProperty("commands", out var cmdProp))
                {
                    manifest.Commands = JsonSerializer.Deserialize<List<RemoteCommandItem>>(
                        cmdProp.GetRawText(),
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
                    return manifest;
                }

                if (doc.RootElement.TryGetProperty("Commands", out var cmdCapProp))
                {
                    manifest.Commands = JsonSerializer.Deserialize<List<RemoteCommandItem>>(
                        cmdCapProp.GetRawText(),
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
                    return manifest;
                }
            }
            else if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                manifest.Commands = JsonSerializer.Deserialize<List<RemoteCommandItem>>(
                    json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
                return manifest;
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private static string? GetCommandsUrl()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, SettingsFileName);
            if (!File.Exists(path)) return null;

            var json = File.ReadAllText(path);
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("RemoteCommandsUrl", out var remoteProp) && !string.IsNullOrWhiteSpace(remoteProp.GetString()))
            {
                return remoteProp.GetString();
            }

            if (doc.RootElement.TryGetProperty("GistUrl", out var gistProp) && !string.IsNullOrWhiteSpace(gistProp.GetString()))
            {
                return gistProp.GetString();
            }
        }
        catch
        {
        }

        return null;
    }

    private static HashSet<string> LoadExecutedCommandIds()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (!File.Exists(CacheFilePath)) return set;
            var json = File.ReadAllText(CacheFilePath);
            var list = JsonSerializer.Deserialize<List<string>>(json);
            if (list != null)
            {
                foreach (var id in list) set.Add(id);
            }
        }
        catch
        {
        }
        return set;
    }

    private static void SaveExecutedCommandIds(HashSet<string> ids)
    {
        try
        {
            var dir = Path.GetDirectoryName(CacheFilePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var json = JsonSerializer.Serialize(ids.ToList());
            File.WriteAllText(CacheFilePath, json);
        }
        catch
        {
        }
    }
}
