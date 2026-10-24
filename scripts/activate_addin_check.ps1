param([Parameter(Mandatory=$true)][string]$ResultPath)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NativeRibbonProbe {
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int RibbonCall(IntPtr self, [MarshalAs(UnmanagedType.BStr)] string id, out IntPtr xml);
    public static string GetXml(IntPtr ribbon) {
        var table = Marshal.ReadIntPtr(ribbon);
        var slot = Marshal.ReadIntPtr(table, 7 * IntPtr.Size);
        var call = (RibbonCall)Marshal.GetDelegateForFunctionPointer(slot, typeof(RibbonCall));
        IntPtr xml;
        int result = call(ribbon, "Microsoft.Word.Document", out xml);
        Marshal.ThrowExceptionForHR(result);
        try { return Marshal.PtrToStringBSTR(xml); }
        finally { if (xml != IntPtr.Zero) Marshal.FreeBSTR(xml); }
    }
}
'@
$instance = New-Object -ComObject 'Gongwen.OfficeAddin.Connect'
$unknown = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($instance)
try {
    foreach ($id in @('B65AD801-ABAF-11D0-BB8B-00A0C90F2744','000C0396-0000-0000-C000-000000000046','C8B34047-5D6E-4A52-9D46-85B523D6F2D1')) {
        $iid = [Guid]$id
        $pointer = [IntPtr]::Zero
        $status = [Runtime.InteropServices.Marshal]::QueryInterface($unknown, [ref]$iid, [ref]$pointer)
        if ($status -ne 0 -or $pointer -eq [IntPtr]::Zero) { throw 'Expected COM interface was not exposed' }
        if ($id -eq '000C0396-0000-0000-C000-000000000046') {
            [xml]$ribbon = [NativeRibbonProbe]::GetXml($pointer)
            if ($ribbon.DocumentElement.LocalName -ne 'customUI') { throw 'Native GetCustomUI did not return Ribbon XML' }
        }
        [void][Runtime.InteropServices.Marshal]::Release($pointer)
    }
} finally { [void][Runtime.InteropServices.Marshal]::Release($unknown) }
@{ passed=$true; process_bits=([IntPtr]::Size*8); interfaces_queried=3; native_ribbon_vtable_verified=$true; wps_word_host_verified=$false } | ConvertTo-Json | Set-Content -Encoding UTF8 $ResultPath
Write-Output 'COM activation and interface checks passed; Office host not exercised.'
