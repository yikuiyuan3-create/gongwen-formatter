using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Gongwen.OfficeAddin;

internal sealed class ActiveDocumentSnapshot
{
    internal ActiveDocumentSnapshot(string fullName, string directory, bool saved)
    {
        FullName = fullName;
        Directory = directory;
        Saved = saved;
    }

    internal string FullName { get; }
    internal string Directory { get; }
    internal bool Saved { get; }
}

internal interface IActiveDocumentHost
{
    ActiveDocumentSnapshot ReadActiveDocument();
}

internal sealed class AddinOperationException : Exception
{
    internal AddinOperationException(string message) : base(message) { }
}

internal sealed class LaunchPlan
{
    internal LaunchPlan(string executablePath, string arguments)
    {
        ExecutablePath = executablePath;
        Arguments = arguments;
    }

    internal string ExecutablePath { get; }
    internal string Arguments { get; }

    internal ProcessStartInfo ToProcessStartInfo() => new ProcessStartInfo
    {
        FileName = ExecutablePath,
        Arguments = Arguments,
        WorkingDirectory = Path.GetDirectoryName(ExecutablePath)!,
        UseShellExecute = false,
        CreateNoWindow = false
    };
}

internal static class LaunchPolicy
{
    internal const string ExecutableName = "GongwenFormatter.exe";

    internal static LaunchPlan ForCurrentDocument(IActiveDocumentHost host, string assemblyLocation)
    {
        if (host is null) throw new AddinOperationException("请先在 WPS 或 Word 中打开需要排版的文档。");
        ActiveDocumentSnapshot document;
        try { document = host.ReadActiveDocument(); }
        catch (AddinOperationException) { throw; }
        catch (Exception) { throw new AddinOperationException("无法读取当前文档，请先打开文档并按 Ctrl+S 保存后重试。"); }
        if (document is null || !document.Saved || string.IsNullOrWhiteSpace(document.Directory) || string.IsNullOrWhiteSpace(document.FullName))
            throw new AddinOperationException("请先按 Ctrl+S 保存当前文档；新文档请另存为 DOCX，再点击“排版当前文档”。");

        var sourcePath = ValidateSavedDocument(document);
        return new LaunchPlan(FindExecutable(assemblyLocation), "--open " + WindowsCommandLine.QuoteArgument(sourcePath));
    }

    internal static LaunchPlan ForTool(string assemblyLocation) => new LaunchPlan(FindExecutable(assemblyLocation), "");

    private static string ValidateSavedDocument(ActiveDocumentSnapshot document)
    {
        try
        {
            if (ContainsControl(document.FullName) || ContainsControl(document.Directory) ||
                !Path.IsPathRooted(document.FullName) || !Path.IsPathRooted(document.Directory))
                throw new AddinOperationException("当前文档没有可用的本地保存路径，请另存为 DOCX 后重试。");
            if (!string.Equals(Path.GetExtension(document.FullName), ".docx", StringComparison.OrdinalIgnoreCase))
                throw new AddinOperationException("请在 WPS 或 Word 中将当前文档另存为 DOCX，再点击“排版当前文档”。");
            var source = Path.GetFullPath(document.FullName);
            var sourceDirectory = Path.GetDirectoryName(source);
            var savedDirectory = Path.GetFullPath(document.Directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (!string.Equals(sourceDirectory?.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), savedDirectory, StringComparison.OrdinalIgnoreCase))
                throw new AddinOperationException("文档保存状态发生变化，请按 Ctrl+S 保存后重试。");
            // Access is read-only and shares the file with Office's existing write
            // handles. The formatter performs the full DOCX and stability checks.
            using (var stream = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                if (stream.Length == 0)
                    throw new AddinOperationException("当前文档尚未完整保存，请按 Ctrl+S 完成保存后重试。");
            return source;
        }
        catch (AddinOperationException) { throw; }
        catch (Exception error) when (error is IOException || error is UnauthorizedAccessException || error is ArgumentException || error is NotSupportedException)
        { throw new AddinOperationException("当前文档无法读取，请确认文件已经保存且具有读取权限。"); }
    }

    private static string FindExecutable(string assemblyLocation)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(assemblyLocation) || ContainsControl(assemblyLocation) || !Path.IsPathRooted(assemblyLocation))
                throw new AddinOperationException("未找到排版工具，请重新安装包含 Office/WPS 加载项的完整安装包。");
            var assemblyDirectory = Path.GetDirectoryName(Path.GetFullPath(assemblyLocation));
            if (assemblyDirectory is null)
                throw new AddinOperationException("未找到排版工具，请重新安装完整安装包。");
            var parent = Path.GetDirectoryName(assemblyDirectory);
            // Deployment: app-0.4.0/office-addin/Gongwen.OfficeAddin.dll.
            // Never search PATH, the document directory or arbitrary ancestors.
            if (parent is not null)
            {
                var deployed = Path.Combine(parent, ExecutableName);
                if (File.Exists(deployed)) return deployed;
            }
            var adjacent = Path.Combine(assemblyDirectory, ExecutableName);
            if (File.Exists(adjacent)) return adjacent;
            throw new AddinOperationException("未找到排版工具，请重新安装包含 Office/WPS 加载项的完整安装包。");
        }
        catch (AddinOperationException) { throw; }
        catch (Exception error) when (error is IOException || error is UnauthorizedAccessException || error is ArgumentException || error is NotSupportedException)
        { throw new AddinOperationException("无法定位排版工具，请重新安装完整安装包。"); }
    }

    private static bool ContainsControl(string value)
    {
        foreach (var character in value) if (char.IsControl(character)) return true;
        return false;
    }
}

internal static class WindowsCommandLine
{
    // Windows CRT / CommandLineToArgvW escaping. Always quote so document text
    // cannot become additional switches. Shell metacharacters remain plain data.
    internal static string QuoteArgument(string value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        if (value.IndexOf('\0') >= 0) throw new ArgumentException("参数不能包含空字符。", nameof(value));
        var quoted = new StringBuilder("\"");
        var slashes = 0;
        foreach (var character in value)
        {
            if (character == '\\') { slashes++; continue; }
            if (character == '"')
                quoted.Append('\\', slashes * 2 + 1).Append('"');
            else
                quoted.Append('\\', slashes).Append(character);
            slashes = 0;
        }
        quoted.Append('\\', slashes * 2).Append('"');
        return quoted.ToString();
    }
}
