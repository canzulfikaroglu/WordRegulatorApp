namespace WordRegulator.Domain.Interfaces;

public interface ITemplateStorageService
{
        string TemplatesDirectoryPath { get; }

        bool HasTemplates { get; }

        string EnsureDirectoryExists();

        void OpenTemplatesFolder();
}
