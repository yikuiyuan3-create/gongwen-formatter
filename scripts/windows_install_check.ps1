param([Parameter(Mandatory=$true)][string]$Installer)
$ErrorActionPreference = 'Stop'
$installRoot = Join-Path $env:RUNNER_TEMP ('Gongwen 插件 # '+[Guid]::NewGuid().ToString('N').Substring(0,8))
$evidenceRoot = Join-Path (Get-Location) 'evidence'
New-Item -ItemType Directory -Force $evidenceRoot | Out-Null
$setup = Start-Process -FilePath (Resolve-Path $Installer) -ArgumentList "/S /D=$installRoot" -Wait -PassThru
if ($setup.ExitCode -ne 0) { throw 'Installer did not complete successfully' }
$version = ([xml](Get-Content Directory.Build.props)).Project.PropertyGroup.Version
$appRoot = Join-Path $installRoot "app-$version"
$addin = Join-Path $appRoot 'office-addin\Gongwen.OfficeAddin.dll'
if (!(Test-Path $addin)) { throw 'Installed add-in is missing' }
$expectedCodeBase='file:///'+$addin.Replace('\','/')
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
    $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    try {
        $key = $registry.OpenSubKey('Software\Classes\CLSID\{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}\InprocServer32')
        if (!$key -or ![string]::Equals([Uri]::UnescapeDataString([string]$key.GetValue('CodeBase')),$expectedCodeBase,[StringComparison]::OrdinalIgnoreCase)) { Write-Output ('Actual CodeBase: '+$key.GetValue('CodeBase')); Write-Output ('Expected target: '+$expectedCodeBase); throw 'COM CodeBase did not resolve to the complete expected install path' }
        $key.Close()
        $versionKey=$registry.OpenSubKey("Software\Classes\CLSID\{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}\InprocServer32\$version.0")
        if(!$versionKey -or ![string]::Equals([Uri]::UnescapeDataString([string]$versionKey.GetValue('CodeBase')),$expectedCodeBase,[StringComparison]::OrdinalIgnoreCase)){throw 'Missing exact framework version registration'}
        $versionKey.Close()
        foreach($id in @('{39F70D9C-9C3A-37DF-BE3A-F9A5813EBFCA}','{33213C29-1EF1-3941-83B7-709ED42ABA31}')) {
            $recordKey=$registry.OpenSubKey("Software\Classes\Record\$id\$version.0")
            if(!$recordKey){throw 'Missing enum record registration'}
            $recordKey.Close()
        }
        foreach ($path in @('Software\Microsoft\Office\Word\Addins\Gongwen.OfficeAddin.Connect','Software\Kingsoft\Office\WPS\Addins\Gongwen.OfficeAddin.Connect')) {
            $key = $registry.OpenSubKey($path)
            if (!$key -or $key.GetValue('LoadBehavior') -ne 3) { throw 'Expected host registration missing' }
            $key.Close()
        }
    } finally { $registry.Close() }
}
$probe = Join-Path (Get-Location) 'scripts\activate_addin_check.ps1'
foreach ($bits in @(64,32)) {
    $shellDir = if ($bits -eq 64) { 'System32' } else { 'SysWOW64' }
    $powershell = Join-Path $env:WINDIR "$shellDir\WindowsPowerShell\v1.0\powershell.exe"
    $result = Join-Path $evidenceRoot "com-activation-$bits.json"
    $stdoutFile = Join-Path $evidenceRoot "com-probe-$bits.stdout.txt"
    $stderrFile = Join-Path $evidenceRoot "com-probe-$bits.stderr.txt"
    $process = Start-Process $powershell -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -ArgumentList "-NoProfile -NonInteractive -File `"$probe`" -ResultPath `"$result`"" -Wait -PassThru
    if (Test-Path $stdoutFile) { Get-Content $stdoutFile | Write-Output }
    if ($process.ExitCode -ne 0 -or !(Test-Path $result)) {
        if (Test-Path $stderrFile) { Get-Content $stderrFile | Write-Output }
        throw "COM activation failed for $bits-bit host (exit $($process.ExitCode))"
    }
    $data = Get-Content $result | ConvertFrom-Json
    if (!$data.passed -or $data.process_bits -ne $bits) { throw 'COM probe result mismatch' }
}
# Preserve an unknown user file: uninstall must not recurse-delete it.
$unknownFile = Join-Path $appRoot 'user-retained.txt'
'fictional user file' | Set-Content $unknownFile
$uninstaller = Join-Path $installRoot '卸载公文一键排版.exe'
$uninstall = Start-Process $uninstaller -ArgumentList '/S' -Wait -PassThru
$deadline = [DateTime]::UtcNow.AddSeconds(30)
while ((Test-Path $addin) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 250 }
if (Test-Path $addin) { throw 'Uninstall did not remove the add-in payload' }
if (!(Test-Path $unknownFile)) { throw 'Uninstall deleted an unknown user file' }
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
    $registry = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    try { if ($registry.OpenSubKey('Software\Classes\CLSID\{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}')) { throw 'Uninstall retained its COM registration' } }
    finally { $registry.Close() }
}
@{ passed=$true; installer_executed=$true; com_32_and_64_verified=$true; special_character_install_path_verified=$true; uninstall_preserved_unknown_file=$true; office_host_verified=$false } | ConvertTo-Json | Set-Content -Encoding UTF8 (Join-Path $evidenceRoot 'windows-install-check.json')
