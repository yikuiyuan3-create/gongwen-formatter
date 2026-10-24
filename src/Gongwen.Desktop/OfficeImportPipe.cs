using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace Gongwen.Desktop;

internal static class OfficeImportPipeName
{
    public static string ForCurrentUser()
    {
        if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("此入口仅用于 Windows。");
        using var identity = WindowsIdentity.GetCurrent();
        string sid = identity.User?.Value ?? throw new InvalidOperationException("无法识别当前 Windows 用户。");
        using var process = Process.GetCurrentProcess();
        return ForIdentity(sid, process.SessionId);
    }
    internal static string ForIdentity(string identity, int sessionId)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(identity + "\n" + sessionId.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        return "GongwenFormatter.Open." + Convert.ToHexString(hash.AsSpan(0, 16));
    }
}

internal sealed class OfficeImportPipeServer(string pipeName, Func<string, CancellationToken, Task<OfficeImportDelivery>> deliver) : IDisposable
{
    private readonly CancellationTokenSource _stop = new();
    private NamedPipeServerStream? _active;
    private int _disposed;
    public async Task RunAsync()
    {
        // Keep the listener alive across clients. On Unix, disposing the final server
        // reference closes the shared listening socket and can reset already queued clients.
        using var pipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly, 4096, 4096);
        Interlocked.Exchange(ref _active, pipe);
        try
        {
            while (!_stop.IsCancellationRequested)
            {
                bool connected = false;
                try
                {
                    await pipe.WaitForConnectionAsync(_stop.Token).ConfigureAwait(false);
                    connected = true;
                    using var requestTimeout = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token);
                    requestTimeout.CancelAfter(TimeSpan.FromSeconds(3));
                    var token = requestTimeout.Token;
                    OfficeImportDelivery status;
                    try
                    {
                        string path = await OfficeImportProtocol.ReadPathAsync(pipe, token).ConfigureAwait(false);
                        status = await deliver(path, token).ConfigureAwait(false);
                    }
                    catch (InvalidDataException) { status = OfficeImportDelivery.Rejected; }
                    await OfficeImportProtocol.WriteDeliveryAsync(pipe, status, token).ConfigureAwait(false);
                    // SendAsync closes its connection after reading the ACK. Wait for that
                    // close before Disconnect so Windows cannot discard an unread ACK.
                    // The same request deadline bounds stalled or extra-data clients.
                    byte[] clientClosed = new byte[1];
                    _ = await pipe.ReadAsync(clientClosed, token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (_stop.IsCancellationRequested) { break; }
                catch (OperationCanceledException) when (connected) { }
                catch (IOException) when (_stop.IsCancellationRequested) { break; }
                catch (IOException) when (connected) { }
                catch (ObjectDisposedException) when (_stop.IsCancellationRequested) { break; }
                finally
                {
                    // A peer close may mark the stream Broken (IsConnected == false),
                    // but Disconnect is still required before the next accept.
                    if (connected && !_stop.IsCancellationRequested)
                    {
                        try { pipe.Disconnect(); }
                        catch (ObjectDisposedException) when (_stop.IsCancellationRequested) { }
                        catch (IOException) when (_stop.IsCancellationRequested) { }
                    }
                }
            }
        }
        finally { Interlocked.CompareExchange(ref _active, null, pipe); }
    }
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _stop.Cancel();
        Interlocked.Exchange(ref _active, null)?.Dispose();
        // The run loop can still observe its cancellation token while the app is exiting.
    }
}

internal static class OfficeImportPipeClient
{
    public static async Task<OfficeImportDelivery> SendAsync(string pipeName, string documentPath, CancellationToken cancellationToken = default)
    {
        if (!OfficeImportCommand.TryValidateDocumentPath(documentPath, out string path))
            throw new InvalidDataException("文档路径无效。");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(5));
        // Only connect to this machine, and verify the server belongs to the current user.
        using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        await pipe.ConnectAsync(timeout.Token).ConfigureAwait(false);
        await OfficeImportProtocol.WritePathAsync(pipe, path, timeout.Token).ConfigureAwait(false);
        return await OfficeImportProtocol.ReadDeliveryAsync(pipe, timeout.Token).ConfigureAwait(false);
    }
}
