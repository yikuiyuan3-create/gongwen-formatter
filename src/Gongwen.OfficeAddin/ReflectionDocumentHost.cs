using System;
using System.Globalization;
using System.Reflection;

namespace Gongwen.OfficeAddin;

internal sealed class ReflectionDocumentHost : IActiveDocumentHost
{
    private readonly object _application;

    internal ReflectionDocumentHost(object application) => _application = application;

    public ActiveDocumentSnapshot ReadActiveDocument()
    {
        try
        {
            var document = ReadProperty(_application, "ActiveDocument");
            if (document is null) throw new AddinOperationException("请先在 WPS 或 Word 中打开需要排版的文档。");
            var fullName = Convert.ToString(ReadProperty(document, "FullName"), CultureInfo.InvariantCulture) ?? "";
            var directory = Convert.ToString(ReadProperty(document, "Path"), CultureInfo.InvariantCulture) ?? "";
            // Read Saved last, after the path metadata. Unsaved host content is
            // never exported, implicitly saved or sent to the formatter.
            var savedValue = ReadProperty(document, "Saved");
            var saved = savedValue is not null && Convert.ToBoolean(savedValue, CultureInfo.InvariantCulture);
            return new ActiveDocumentSnapshot(fullName, directory, saved);
        }
        catch (AddinOperationException) { throw; }
        catch (Exception) { throw new AddinOperationException("无法读取当前文档，请先按 Ctrl+S 保存 DOCX 后重试；也可从排版工具手动导入。"); }
    }

    private static object? ReadProperty(object target, string name) => target.GetType().InvokeMember(name,
        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding,
        null, target, null, CultureInfo.InvariantCulture);
}
