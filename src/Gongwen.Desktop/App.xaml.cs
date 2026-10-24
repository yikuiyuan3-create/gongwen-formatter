using System.IO;
using System.Text.Json;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace Gongwen.Desktop;

public partial class App : Application
{
    private Mutex? _instanceMutex;
    private bool _ownsMutex;
    private OfficeImportPipeServer? _officeImportServer;

    protected override async void OnStartup(StartupEventArgs e)
    {
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        base.OnStartup(e);
        if (!OfficeImportCommand.TryParse(e.Args, out var launch, out string error))
        { ShowStartupMessage(error, MessageBoxImage.Warning); Shutdown(2); return; }
        _instanceMutex = new Mutex(true, @"Local\GongwenFormatter.SingleInstance", out bool firstInstance);
        _ownsMutex = firstInstance;
        if (!firstInstance)
        {
            if (launch.Mode == OfficeImportLaunchMode.Open)
            {
                try
                {
                    var delivery = await OfficeImportPipeClient.SendAsync(OfficeImportPipeName.ForCurrentUser(), launch.DocumentPath!);
                    if (delivery == OfficeImportDelivery.Accepted) { Shutdown(0); return; }
                    ShowStartupMessage(delivery == OfficeImportDelivery.Busy
                        ? "现有窗口正在处理文档或等待操作。请先完成当前操作，再点击公文排版按钮。"
                        : "现有窗口未接收文档。请切换到公文一键排版窗口检查状态后重试。", MessageBoxImage.Information);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                { ShowStartupMessage("未能把文档传给已打开的窗口。请稍后重试；旧版程序需先关闭并升级。", MessageBoxImage.Warning); }
                Shutdown(3); return;
            }
            ShowStartupMessage("公文一键排版已在运行，请切换到现有窗口。", MessageBoxImage.Information);
            Shutdown(2); return;
        }
        DispatcherUnhandledException += (_, eventArgs) =>
        {
            eventArgs.Handled = true;
            ShowStartupMessage("程序遇到异常，请关闭后重新打开。原稿不会被覆盖。", MessageBoxImage.Error);
        };
        var window = new MainWindow();
        MainWindow = window;
        ShutdownMode = ShutdownMode.OnMainWindowClose;
        bool initialContentHandled = false;
        window.ContentRendered += async (_, _) =>
        {
            if (initialContentHandled) return;
            initialContentHandled = true;
            if (launch.Mode == OfficeImportLaunchMode.SmokeTest)
            {
                int exitCode = 1;
                try
                {
                    var result = await window.RunSmokeTestAsync(launch.DocumentPath!, launch.OutputPath!);
                    string resultPath = Path.GetFullPath(launch.ResultPath!);
                    Directory.CreateDirectory(Path.GetDirectoryName(resultPath)!);
                    // Never persist source filenames, document contents, or exception dumps.
                    await File.WriteAllTextAsync(resultPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
                    exitCode = result.Passed ? 0 : 1;
                }
                catch (Exception ex) when (ex is not OutOfMemoryException) { exitCode = 1; }
                Shutdown(exitCode);
                return;
            }
            StartOfficeImportServer(window);
            if (launch.Mode == OfficeImportLaunchMode.Open)
                window.QueueOfficeImport(launch.DocumentPath!);
        };
        window.Show();
    }
    private void StartOfficeImportServer(MainWindow window)
    {
        try
        {
            _officeImportServer = new OfficeImportPipeServer(OfficeImportPipeName.ForCurrentUser(), (path, token) =>
                window.Dispatcher.InvokeAsync(() => window.QueueOfficeImport(path), DispatcherPriority.Normal, token).Task);
            _ = ObserveOfficeImportServerAsync(window, _officeImportServer);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException) { window.ReportOfficeImportUnavailable(); }
    }
    private async Task ObserveOfficeImportServerAsync(MainWindow window, OfficeImportPipeServer server)
    {
        try { await server.RunAsync().ConfigureAwait(false); }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            if (Dispatcher.HasShutdownStarted) return;
            try { await Dispatcher.InvokeAsync(window.ReportOfficeImportUnavailable); }
            catch (TaskCanceledException) { }
            catch (InvalidOperationException) { }
        }
    }
    private static void ShowStartupMessage(string message, MessageBoxImage image) =>
        MessageBox.Show(message, "公文一键排版", MessageBoxButton.OK, image);

    protected override void OnExit(ExitEventArgs e)
    {
        _officeImportServer?.Dispose();
        if (_ownsMutex)
        {
            try { _instanceMutex?.ReleaseMutex(); }
            catch (ApplicationException) { }
        }
        _instanceMutex?.Dispose();
        base.OnExit(e);
    }
}
