using System;
using System.Linq;
using System.Runtime.CompilerServices;
using Gongwen.OfficeAddin;
using Microsoft.Win32;

namespace Gongwen.AddinRegistrar;

internal static class Program
{
    [STAThread]
    private static int Main(string[] arguments)
    {
        try
        {
            if (!RegistrarCommandLine.TryParse(arguments, out var operation)) return 2;
            return RegistrationRunner.Execute(operation, ReadLoadedAddinMetadata(), new WindowsUserRegistry());
        }
        catch (Exception) { return 2; }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static AddinRegistrationMetadata ReadLoadedAddinMetadata()
    {
        var connect = typeof(Connect);
        var assembly = connect.Assembly;
        var version = assembly.GetName().Version;
        // .NET Framework RegAsm / RegistrationServices uses Assembly.CodeBase.
        // Do not substitute EscapedCodeBase, Uri.AbsoluteUri, Shell URL helpers,
        // a caller-supplied path or the registrar's own assembly identity.
        return new AddinRegistrationMetadata(connect.FullName!, connect.GUID, assembly.FullName!, version!,
            assembly.ImageRuntimeVersion, assembly.CodeBase!, new[] { typeof(ExtConnectMode), typeof(ExtDisconnectMode) }
                .Select(type => new EnumRegistrationMetadata(type.FullName!, type.GUID)).ToArray());
    }
}

internal sealed class WindowsUserRegistry : IUserRegistry
{
    public void CreateOrUpdate(RegistrationView view, RegistrationNode node)
    {
        using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, ToNativeView(view)))
        using (var key = root.CreateSubKey(node.Path, writable: true))
        {
            if (key is null) throw new InvalidOperationException("无法创建当前用户注册项。");
            foreach (var value in node.Values)
                key.SetValue(value.Name, value.Value, value.Kind == PlanValueKind.DWord ? RegistryValueKind.DWord : RegistryValueKind.String);
        }
    }

    public void DeleteOwnedTree(RegistrationView view, string path)
    {
        using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, ToNativeView(view)))
            root.DeleteSubKeyTree(path, throwOnMissingSubKey: false);
    }

    private static RegistryView ToNativeView(RegistrationView view) => view == RegistrationView.Registry32
        ? RegistryView.Registry32 : RegistryView.Registry64;
}
