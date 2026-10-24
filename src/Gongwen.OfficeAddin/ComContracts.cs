using System;
using System.Runtime.InteropServices;

namespace Gongwen.OfficeAddin;

// Standard shared Office COM contracts. No Office PIA or VSTO runtime is used.
[ComVisible(true)]
public enum ExtConnectMode
{
    AfterStartup = 0, Startup = 1, External = 2, CommandLine = 3, Solution = 4, UiSetup = 5
}

[ComVisible(true)]
public enum ExtDisconnectMode
{
    HostShutdown = 0, UserClosed = 1, UiSetupComplete = 2, SolutionClosed = 3
}

[ComVisible(true)]
[Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IDTExtensibility2
{
    [DispId(1)]
    void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, ExtConnectMode connectMode,
        [MarshalAs(UnmanagedType.IDispatch)] object addInInstance,
        [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array? custom);

    [DispId(2)]
    void OnDisconnection(ExtDisconnectMode removeMode,
        [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array? custom);

    [DispId(3)]
    void OnAddInsUpdate([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array? custom);

    [DispId(4)]
    void OnStartupComplete([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array? custom);

    [DispId(5)]
    void OnBeginShutdown([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array? custom);
}

[ComVisible(true)]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    [return: MarshalAs(UnmanagedType.BStr)]
    string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
}

// A fixed default IDispatch interface makes Ribbon callback lookup independent
// of class member ordering, automatic class interfaces and host type libraries.
[ComVisible(true)]
[Guid("C8B34047-5D6E-4A52-9D46-85B523D6F2D1")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IGongwenRibbonCallbacks
{
    [DispId(101)]
    void FormatCurrentDocument([MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(102)]
    void OpenTool([MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(103)]
    void ShowHelp([MarshalAs(UnmanagedType.IDispatch)] object control);
}
