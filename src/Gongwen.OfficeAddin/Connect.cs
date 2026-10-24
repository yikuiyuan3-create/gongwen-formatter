using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Gongwen.OfficeAddin;

[ComVisible(true)]
[Guid("3B873652-4D36-4CE7-B43C-B93B9F1B7A01")]
[ProgId("Gongwen.OfficeAddin.Connect")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IGongwenRibbonCallbacks))]
public sealed class Connect : IDTExtensibility2, IRibbonExtensibility, IGongwenRibbonCallbacks
{
    private object? _application;

    public void OnConnection(object application, ExtConnectMode connectMode, object addInInstance, ref Array? custom)
    {
        try { _application = application; }
        catch (Exception) { _application = null; }
    }

    public void OnDisconnection(ExtDisconnectMode removeMode, ref Array? custom) => _application = null;
    public void OnAddInsUpdate(ref Array? custom) { }
    public void OnStartupComplete(ref Array? custom) { }
    public void OnBeginShutdown(ref Array? custom) => _application = null;

    public string GetCustomUI(string ribbonId)
    {
        try { return RibbonDefinition.Xml; }
        catch (Exception) { return ""; }
    }

    public void FormatCurrentDocument(object control) => AtHostBoundary(() =>
    {
        if (_application is null)
            throw new AddinOperationException("请先在 WPS 或 Word 中打开需要排版的文档。");
        StartTool(LaunchPolicy.ForCurrentDocument(new ReflectionDocumentHost(_application), Assembly.GetExecutingAssembly().Location));
    });

    public void OpenTool(object control) => AtHostBoundary(() =>
        StartTool(LaunchPolicy.ForTool(Assembly.GetExecutingAssembly().Location)));

    public void ShowHelp(object control) => AtHostBoundary(() => ShowMessage(
        "1. 在 WPS 或 Word 中打开文档，按 Ctrl+S 保存；旧格式请另存为 DOCX。\n\n" +
        "2. 点击“排版当前文档”，工具将读取已保存的磁盘文件。修改后仍需先保存，再次导入。\n\n" +
        "3. 在工具中检查排版并另存新副本，原稿保持不变。\n\n" +
        "加载项适用于支持 COM 加载项与 Ribbon 的 Windows 桌面版 Word/WPS。"));

    private static void StartTool(LaunchPlan plan)
    {
        using (var process = Process.Start(plan.ToProcessStartInfo()))
            if (process is null) throw new AddinOperationException("无法启动排版工具，请重新安装完整安装包。");
    }

    private static void AtHostBoundary(Action action)
    {
        try { action(); }
        catch (AddinOperationException error) { ShowMessage(error.Message); }
        catch (Exception) { ShowMessage("无法读取当前文档或启动排版工具。请先保存文档，或从工具窗口手动导入。"); }
    }

    private static void ShowMessage(string message)
    {
        try { MessageBox.Show(message, "公文一键排版", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        catch (Exception) { /* Keep Office responsive if its UI is unavailable. */ }
    }
}
