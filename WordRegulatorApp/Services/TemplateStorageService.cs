using System.Diagnostics;
using System.IO;
using WordRegulator.Domain.Interfaces;

namespace WordRegulatorApp.Services;

public class TemplateStorageService : ITemplateStorageService
{
    public string TemplatesDirectoryPath { get; }

    public bool HasTemplates
    {
        get
        {
            if (!Directory.Exists(TemplatesDirectoryPath)) return false;

            var docx = Directory.GetFiles(TemplatesDirectoryPath, "*.docx", SearchOption.TopDirectoryOnly)
                                .Where(f => !Path.GetFileName(f).StartsWith("~$"));
            var xlsx = Directory.GetFiles(TemplatesDirectoryPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                                .Where(f => !Path.GetFileName(f).StartsWith("~$"));

            return docx.Any() || xlsx.Any();
        }
    }

    public TemplateStorageService()
    {
        
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        TemplatesDirectoryPath = Path.Combine(appData, "WordRegulator", "Templates");
    }

    public string EnsureDirectoryExists()
    {
        if (!Directory.Exists(TemplatesDirectoryPath))
        {
            Directory.CreateDirectory(TemplatesDirectoryPath);
        }
        return TemplatesDirectoryPath;
    }

    public void OpenTemplatesFolder()
    {
        EnsureDirectoryExists();

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{TemplatesDirectoryPath}\"",
                UseShellExecute = true
            });
        }
        catch
        {
            
        }
    }
}
