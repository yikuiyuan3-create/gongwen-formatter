$ErrorActionPreference='Stop'
$installer=(Get-ChildItem work/diagnostic-input -Recurse -Filter '*_Setup_*_x64.exe' | Select-Object -First 1).FullName
$version=([xml](Get-Content Directory.Build.props)).Project.PropertyGroup.Version
$probe=Join-Path (Get-Location) 'scripts/activate_addin_check.ps1'
$powershell=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
foreach($suffix in @('GongwenAscii','Gongwen 插件 #')) {
    $target=Join-Path $env:RUNNER_TEMP ($suffix+[Guid]::NewGuid().ToString('N').Substring(0,6))
    $p=Start-Process $installer -ArgumentList "/S /D=$target" -Wait -PassThru
    if($p.ExitCode -ne 0){throw 'Diagnostic installation failed'}
    $dll=Join-Path $target "app-$version/office-addin/Gongwen.OfficeAddin.dll"
    Write-Output "CASE $suffix"
    Write-Output ([Reflection.AssemblyName]::GetAssemblyName($dll).FullName)
    $key=Get-Item 'HKCU:/Software/Classes/CLSID/{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}/InprocServer32'
    foreach($name in @('Class','Assembly','RuntimeVersion','CodeBase')){Write-Output "$name=$($key.GetValue($name))"}
    $uri=[Uri]$key.GetValue('CodeBase')
    Write-Output "LOCAL=$($uri.LocalPath); FRAGMENT=$($uri.Fragment); EXISTS=$([IO.File]::Exists($uri.LocalPath))"
    $regfile=Join-Path $env:RUNNER_TEMP 'expected-registration.reg'
    $regasm=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/RegAsm.exe'
    & $regasm $dll "/regfile:$regfile" /codebase
    if(Test-Path $regfile){Get-Content $regfile | Write-Output}
    $result=Join-Path $env:RUNNER_TEMP 'diagnostic-com.json'
    $stdout=Join-Path $env:RUNNER_TEMP 'diagnostic-com.stdout.txt'
    $stderr=Join-Path $env:RUNNER_TEMP 'diagnostic-com.stderr.txt'
    $p=Start-Process $powershell -ArgumentList "-NoProfile -NonInteractive -File `"$probe`" -ResultPath `"$result`"" -RedirectStandardOutput $stdout -RedirectStandardError $stderr -Wait -PassThru
    Write-Output "PROBE_EXIT=$($p.ExitCode)"
    Get-Content $stdout,$stderr | Write-Output
    $loadCommand='try { $a=[Reflection.Assembly]::LoadFrom('''+$dll.Replace("'","''")+'''); [Activator]::CreateInstance($a.GetType(''Gongwen.OfficeAddin.Connect'')); ''MANAGED_LOAD_OK'' } catch { $_.Exception.ToString(); exit 1 }'
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($loadCommand))
    & $powershell -NoProfile -NonInteractive -EncodedCommand $encoded
    $uninstaller=Join-Path $target '卸载公文一键排版.exe'
    Start-Process $uninstaller -ArgumentList '/S' -Wait | Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(30)
    while((Test-Path $dll) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 200}
}
