using System.IO;
using System.Text.Json;
using Gongwen.Core;

namespace Gongwen.Desktop;

internal static class AuditStore
{
    public static async Task WriteCompletedAsync(FormatTemplate template, FormatResult result, int editedParagraphCount = 0)
    {
        var taskId = Guid.NewGuid();
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GongwenFormatter", "audit");
        Directory.CreateDirectory(directory);
        string target = Path.Combine(directory, taskId.ToString("N") + ".json");
        string temporary = target + ".tmp";
        // Store traceable metadata only. Do not add filenames, paths, document text, or error dumps.
        var record = new
        {
            TaskId = taskId,
            OccurredAtUtc = DateTimeOffset.UtcNow,
            TemplateId = template.Id,
            TemplateParameters = template,
            result.SourceSha256,
            result.OutputSha256,
            result.FormattedParagraphCount,
            EditedParagraphCount = editedParagraphCount,
            result.EditProtectionRemoved,
            Result = "completed"
        };
        try
        {
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(record, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temporary, target, overwrite: false);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }
}
