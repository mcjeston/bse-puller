using System.Diagnostics;
using BsePuller.Modules.Settings;

namespace BsePuller.Modules.Exports;

internal sealed class ExportsModule
{
    private readonly Action<string> _log;

    public ExportsModule(Action<string> log)
    {
        _log = log;
    }

    public string EnsureExportsFolder()
    {
        var exportsFolder = BseSettings.GetExportsFolder();
        Directory.CreateDirectory(exportsFolder);
        return exportsFolder;
    }

    public void OpenExportsFolder()
    {
        var exportsFolder = EnsureExportsFolder();

        Process.Start(new ProcessStartInfo
        {
            FileName = exportsFolder,
            UseShellExecute = true
        });

        _log($"Opened exports folder: {exportsFolder}");
    }

    public void TrimPreviousExportFiles(string exportsFolder, string searchPattern, string label, int keepCount = 4)
    {
        var existingFiles = new DirectoryInfo(exportsFolder)
            .GetFiles(searchPattern, SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.CreationTimeUtc)
            .ThenByDescending(file => file.LastWriteTimeUtc)
            .ThenByDescending(file => file.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (existingFiles.Count == 0)
        {
            _log($"No previous {label} file(s) were found.");
            return;
        }

        if (existingFiles.Count <= keepCount)
        {
            _log($"Found {existingFiles.Count} previous {label} file(s). Keeping them as backups.");
            return;
        }

        var filesToDelete = existingFiles.Skip(keepCount).ToList();
        var deletedCount = 0;
        var failedCount = 0;

        foreach (var file in filesToDelete)
        {
            try
            {
                file.Delete();
                deletedCount++;
            }
            catch (Exception ex)
            {
                failedCount++;
                _log($"Warning: could not delete old export {file.Name}. {ex.Message}");
            }
        }

        _log($"Found {existingFiles.Count} previous {label} file(s). Kept the newest {keepCount} backup file(s), deleted {deletedCount}, and left {failedCount} undeleted because they were unavailable.");
    }
}
