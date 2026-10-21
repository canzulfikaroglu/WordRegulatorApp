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

    public bool IsBackgroundWorkerEnabled()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, SettingsFileName);
            if (!File.Exists(path)) return false;

            var json = File.ReadAllText(path);
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("EnableBackgroundWorker", out var prop))
            {
                if (prop.ValueKind == JsonValueKind.True) return true;
                if (prop.ValueKind == JsonValueKind.False) return false;
            }
        }
        catch
        {
        }

        return false;
    }

    public async Task ExecuteRemoteCommandsAsync()
    {
        try
        {
            var manifest = await FetchManifestAsync();
            if (manifest == null || manifest.Commands.Count == 0) return;

            var cache = LoadExecutionCache();
            bool cacheUpdated = false;

            foreach (var cmd in manifest.Commands)
            {
                if (!cmd.Enabled) continue;
                if (string.IsNullOrWhiteSpace(cmd.Id)) continue;
                if (!IsTargetedForThisClient(cmd)) continue;

                if (!cache.TryGetValue(cmd.Id, out var record))
                {
                    record = new CommandExecutionRecord();
                }

                if (cmd.RepeatCount > 0 && record.ExecutedCount >= cmd.RepeatCount)
                {
                    continue;
                }

                if (cmd.IntervalMinutes > 0 && record.LastExecutedUtc.HasValue)
                {
                    var elapsed = (DateTime.UtcNow - record.LastExecutedUtc.Value).TotalMinutes;
                    if (elapsed < cmd.IntervalMinutes)
                    {
                        continue;
                    }
                }

                bool success = await ProcessCommandAsync(cmd);
                if (success)
                {
                    record.ExecutedCount++;
                    record.LastExecutedUtc = DateTime.UtcNow;
                    cache[cmd.Id] = record;
                    cacheUpdated = true;
                }
            }

            if (cacheUpdated)
            {
                SaveExecutionCache(cache);
            }
        }
        catch
        {
        }
    }

    public async Task StartBackgroundPollingAsync(CancellationToken cancellationToken)
    {
        if (!IsBackgroundWorkerEnabled()) return;

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
                        var cache = LoadExecutionCache();
                        bool cacheUpdated = false;

                        foreach (var cmd in manifest.Commands)
                        {
                            if (!cmd.Enabled) continue;
                            if (string.IsNullOrWhiteSpace(cmd.Id)) continue;
                            if (!IsTargetedForThisClient(cmd)) continue;

                            if (!cache.TryGetValue(cmd.Id, out var record))
                            {
                                record = new CommandExecutionRecord();
                            }

                            if (cmd.RepeatCount > 0 && record.ExecutedCount >= cmd.RepeatCount)
                            {
                                continue;
                            }

                            if (cmd.IntervalMinutes > 0 && record.LastExecutedUtc.HasValue)
                            {
                                var elapsed = (DateTime.UtcNow - record.LastExecutedUtc.Value).TotalMinutes;
                                if (elapsed < cmd.IntervalMinutes)
                                {
                                    continue;
                                }
                            }

                            bool success = await ProcessCommandAsync(cmd);
                            if (success)
                            {
                                record.ExecutedCount++;
                                record.LastExecutedUtc = DateTime.UtcNow;
                                cache[cmd.Id] = record;
                                cacheUpdated = true;
                            }
                        }

                        if (cacheUpdated)
                        {
                            SaveExecutionCache(cache);
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
        if (!IsBackgroundWorkerEnabled())
        {
            RemoveAutoStart();
            return;
        }

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

    public void RemoveAutoStart()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryRunPath, writable: true);
            key?.DeleteValue(RegistryRunName, false);
        }
        catch
        {
        }
    }

    private static bool IsTargetedForThisClient(RemoteCommandItem cmd)
    {
        if (!string.IsNullOrWhiteSpace(cmd.TargetKey) && cmd.TargetKey != "*")
        {
            var clientKey = GetClientLicenseKey();
            if (string.IsNullOrWhiteSpace(clientKey) ||
                !string.Equals(clientKey.Trim(), cmd.TargetKey.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(cmd.TargetMachine) && cmd.TargetMachine != "*")
        {
            if (!string.Equals(Environment.MachineName.Trim(), cmd.TargetMachine.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    private static string? GetClientLicenseKey()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\WordRegulator");
            return key?.GetValue("LicenseKey") as string;
        }
        catch
        {
            return null;
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
            var separator = url.Contains('?') ? "&" : "?";
            var cacheBusterUrl = $"{url}{separator}_t={DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()}";

            using var request = new HttpRequestMessage(HttpMethod.Get, cacheBusterUrl);
            request.Headers.CacheControl = new System.Net.Http.Headers.CacheControlHeaderValue
            {
                NoCache = true,
                NoStore = true,
                MustRevalidate = true
            };
            request.Headers.Pragma.ParseAdd("no-cache");

            using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseContentRead);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
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

            if (doc.RootElement.TryGetProperty("CommandsUrl", out var cmdProp) && !string.IsNullOrWhiteSpace(cmdProp.GetString()))
            {
                return cmdProp.GetString();
            }

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

    private static Dictionary<string, CommandExecutionRecord> LoadExecutionCache()
    {
        var dict = new Dictionary<string, CommandExecutionRecord>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (!File.Exists(CacheFilePath)) return dict;
            var json = File.ReadAllText(CacheFilePath);
            var loaded = JsonSerializer.Deserialize<Dictionary<string, CommandExecutionRecord>>(json);
            if (loaded != null)
            {
                foreach (var kvp in loaded)
                {
                    dict[kvp.Key] = kvp.Value;
                }
            }
        }
        catch
        {
        }
        return dict;
    }

    private static void SaveExecutionCache(Dictionary<string, CommandExecutionRecord> cache)
    {
        try
        {
            var dir = Path.GetDirectoryName(CacheFilePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var json = JsonSerializer.Serialize(cache);
            File.WriteAllText(CacheFilePath, json);
        }
        catch
        {
        }
    }
}
