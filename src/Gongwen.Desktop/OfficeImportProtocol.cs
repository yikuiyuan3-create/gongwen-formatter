using System.Buffers.Binary;
using System.IO;
using System.Text;

namespace Gongwen.Desktop;

internal enum OfficeImportDelivery : byte { Accepted = 1, Busy = 2, Rejected = 3 }

/// <summary>Bounded UTF-8 path frames; no executable commands, document text, or user identifiers are transmitted.</summary>
internal static class OfficeImportProtocol
{
    internal const int MaxPayloadBytes = 128 * 1024;
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    public static async Task WritePathAsync(Stream stream, string path, CancellationToken cancellationToken)
    {
        if (!OfficeImportCommand.TryValidateDocumentPath(path, out string normalized))
            throw new InvalidDataException("文档路径无效。");
        byte[] payload = StrictUtf8.GetBytes(normalized);
        if (payload.Length is < 1 or > MaxPayloadBytes) throw new InvalidDataException("插件消息超出范围。");
        byte[] header = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(header, payload.Length);
        await stream.WriteAsync(header, cancellationToken).ConfigureAwait(false);
        await stream.WriteAsync(payload, cancellationToken).ConfigureAwait(false);
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }
    public static async Task<string> ReadPathAsync(Stream stream, CancellationToken cancellationToken)
    {
        byte[] header = new byte[4];
        await stream.ReadExactlyAsync(header, cancellationToken).ConfigureAwait(false);
        int length = BinaryPrimitives.ReadInt32LittleEndian(header);
        if (length is < 1 or > MaxPayloadBytes) throw new InvalidDataException("插件消息超出范围。");
        byte[] payload = new byte[length];
        await stream.ReadExactlyAsync(payload, cancellationToken).ConfigureAwait(false);
        string input;
        try { input = StrictUtf8.GetString(payload); }
        catch (DecoderFallbackException) { throw new InvalidDataException("插件消息编码无效。"); }
        if (!OfficeImportCommand.TryValidateDocumentPath(input, out string path))
            throw new InvalidDataException("文档路径无效。");
        return path;
    }
    public static async Task WriteDeliveryAsync(Stream stream, OfficeImportDelivery delivery, CancellationToken cancellationToken)
    {
        await stream.WriteAsync(new byte[] { (byte)delivery }, cancellationToken).ConfigureAwait(false);
        await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }
    public static async Task<OfficeImportDelivery> ReadDeliveryAsync(Stream stream, CancellationToken cancellationToken)
    {
        byte[] response = new byte[1];
        await stream.ReadExactlyAsync(response, cancellationToken).ConfigureAwait(false);
        return response[0] is >= 1 and <= 3 ? (OfficeImportDelivery)response[0] : throw new InvalidDataException("插件响应无效。");
    }
}
