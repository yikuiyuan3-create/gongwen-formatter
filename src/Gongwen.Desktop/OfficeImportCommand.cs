using System.Text;

namespace Gongwen.Desktop;

internal enum OfficeImportLaunchMode { Normal, Open, SmokeTest }
internal sealed record OfficeImportLaunch(OfficeImportLaunchMode Mode, string? DocumentPath = null, string? OutputPath = null, string? ResultPath = null);

/// <summary>Pure argument/path validation. Document existence and DOCX contents are checked by the normal document loader.</summary>
internal static class OfficeImportCommand
{
    internal const int MaxPathCharacters = 32_760;
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    public static bool TryParse(IReadOnlyList<string> arguments, out OfficeImportLaunch launch, out string error)
    {
        launch = new(OfficeImportLaunchMode.Normal);
        error = "";
        if (arguments.Count == 0) return true;
        if (arguments.Count == 2 && arguments[0] == "--open")
        {
            if (!TryValidateDocumentPath(arguments[1], out string path))
            { error = "插件传入的文档路径无效。请使用已保存的本地或共享文件夹中的 DOCX。"; return false; }
            launch = new(OfficeImportLaunchMode.Open, path);
            return true;
        }
        if (arguments.Count == 4 && arguments[0] == "--smoke-test" && arguments.Skip(1).All(IsSafeArgument))
        {
            // Keep the existing smoke-test contract, including relative fixture/output paths.
            launch = new(OfficeImportLaunchMode.SmokeTest, arguments[1], arguments[2], arguments[3]);
            return true;
        }
        error = "启动参数无效。支持直接启动，或使用 --open 后跟一个已保存的 DOCX 路径。";
        return false;
    }

    public static bool TryValidateDocumentPath(string? input, out string path)
    {
        path = "";
        if (!IsSafeArgument(input)) return false;
        string normalized = input!.Replace('/', '\\');
        if (normalized.IndexOfAny(['<', '>', '"', '|', '?', '*']) >= 0) return false;
        string prefix;
        string[] components;
        if (normalized.StartsWith(@"\\", StringComparison.Ordinal))
        {
            // Exclude Win32 device/extended namespaces, URI syntax, and incomplete UNC roots.
            var parts = normalized[2..].Split('\\');
            if (parts.Length < 3 || !ValidComponent(parts[0]) || !ValidComponent(parts[1]) ||
                parts[0] is "." or ".." || parts[1] is "." or ".." || normalized.Contains(':')) return false;
            prefix = @"\\" + parts[0] + "\\" + parts[1] + "\\";
            components = parts[2..];
        }
        else
        {
            if (normalized.Length < 4 || !char.IsAsciiLetter(normalized[0]) || normalized[1] != ':' || normalized[2] != '\\' ||
                normalized[2..].Contains(':')) return false;
            prefix = char.ToUpperInvariant(normalized[0]) + @":\";
            components = normalized[3..].Split('\\');
        }
        var result = new List<string>();
        foreach (string component in components)
        {
            if (component == ".") continue;
            if (component == "..")
            {
                if (result.Count == 0) return false;
                result.RemoveAt(result.Count - 1);
                continue;
            }
            if (!ValidComponent(component) || IsReservedDeviceName(component)) return false;
            result.Add(component);
        }
        if (result.Count == 0 || !result[^1].EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) return false;
        path = prefix + string.Join('\\', result);
        return path.Length <= MaxPathCharacters;
    }

    private static bool IsSafeArgument(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaxPathCharacters ||
            value.Any(c => char.IsControl(c) || c is '\u2028' or '\u2029')) return false;
        try { _ = StrictUtf8.GetByteCount(value); return true; }
        catch (EncoderFallbackException) { return false; }
    }
    private static bool ValidComponent(string value) => value.Length is > 0 and <= 255 && value is not "." and not ".." &&
        !value.EndsWith(' ') && !value.EndsWith('.') && !value.Contains(':');
    private static bool IsReservedDeviceName(string component)
    {
        string name = component.Split('.')[0].TrimEnd(' ').ToUpperInvariant();
        return name is "CON" or "PRN" or "AUX" or "NUL" ||
            (name.Length == 4 && (name.StartsWith("COM", StringComparison.Ordinal) || name.StartsWith("LPT", StringComparison.Ordinal)) && name[3] is >= '1' and <= '9' or '¹' or '²' or '³');
    }
}
