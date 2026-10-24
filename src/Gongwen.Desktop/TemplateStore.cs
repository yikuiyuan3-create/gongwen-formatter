using System.IO;
using System.Text.Json;
using Gongwen.Core;

namespace Gongwen.Desktop;

internal static class TemplateStore
{
    private static string ConfigPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GongwenFormatter", "templates.json");
    public static List<FormatTemplate> Load(out bool failed)
    {
        failed = false;
        List<FormatTemplate> result = [.. BuiltInTemplates.All];
        try
        {
            if (!File.Exists(ConfigPath)) return result;
            if (new FileInfo(ConfigPath).Length > 128 * 1024) { failed = true; return result; }
            var saved = JsonSerializer.Deserialize<List<FormatTemplate?>>(File.ReadAllText(ConfigPath));
            if (saved is null) { failed = true; return result; }
            foreach (var item in saved)
            {
                if (item is null) { failed = true; continue; }
                var index = result.FindIndex(t => t.Id == item.Id);
                if (index >= 0 && TemplateValidation.IsValid(item)) result[index] = item with { Name = result[index].Name };
                else failed = true;
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or NotSupportedException)
        { failed = true; }
        return result;
    }
    public static void Save(IReadOnlyList<FormatTemplate> templates)
    {
        string path = ConfigPath;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            // This payload contains template parameters only; never document paths or text.
            File.WriteAllText(temporary, JsonSerializer.Serialize(templates, new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temporary, path, overwrite: true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}

internal static class TemplateValidation
{
    // Parameter limits match DocxFormatter.ValidateTemplate; keep them synchronized with the format engine.
    public static bool IsValid(FormatTemplate t) =>
        BuiltInTemplates.All.Any(p => p.Id == t.Id) &&
        ValidFont(t.TitleFont) && ValidFont(t.BodyFont) && ValidFont(t.Heading1Font) && ValidFont(t.Heading2Font) && ValidFont(t.LatinFont) &&
        InRange(t.TopMarginCm, 0.5, 8) && InRange(t.BottomMarginCm, 0.5, 8) &&
        InRange(t.LeftMarginCm, 0.5, 8) && InRange(t.RightMarginCm, 0.5, 8) &&
        t.LeftMarginCm + t.RightMarginCm < 16 && t.TopMarginCm + t.BottomMarginCm < 24.7 &&
        InRange(t.BodyFontPt, 5, 48) && InRange(t.TitleFontPt, 5, 72) &&
        t.LinesPerPage is >= 8 and <= 60 && t.CharactersPerLine is >= 10 and <= 60;
    private static bool InRange(double value, double min, double max) => double.IsFinite(value) && value >= min && value <= max;
    private static bool ValidFont(string font) => !string.IsNullOrWhiteSpace(font) && font.Length <= 100 && !font.Any(char.IsControl);
}
