param([Parameter(Mandatory=$true)][string]$Installer)
# CI only: create and remove a disposable non-administrator test account.
$ErrorActionPreference='Stop'
$name='gwci'+[Guid]::NewGuid().ToString('N').Substring(0,10)
$secret=ConvertTo-SecureString ([Guid]::NewGuid().ToString('N')+'aA!9') -AsPlainText -Force
$account=$null
$sandbox=Join-Path $env:RUNNER_TEMP $name
$evidenceRoot=Join-Path (Get-Location) 'evidence'
New-Item -ItemType Directory -Force $evidenceRoot | Out-Null
try {
    $account=New-LocalUser -Name $name -Password $secret -AccountNeverExpires -PasswordNeverExpires
    $group=Get-LocalGroup -SID 'S-1-5-32-545'
    if(!(Get-LocalGroupMember $group | Where-Object { $_.SID -eq $account.SID })) {
        Add-LocalGroupMember -Group $group -Member $account
    }
    New-Item -ItemType Directory -Force (Join-Path $sandbox 'scripts'),(Join-Path $sandbox 'temp'),(Join-Path $sandbox 'evidence') | Out-Null
    Copy-Item (Resolve-Path $Installer) (Join-Path $sandbox 'setup.exe')
    Copy-Item Directory.Build.props $sandbox
    Copy-Item scripts/windows_install_check.ps1,scripts/activate_addin_check.ps1 (Join-Path $sandbox 'scripts')
    $entry=Join-Path $sandbox 'entry.ps1'
    @'
$ErrorActionPreference='Stop'
Set-Location $PSScriptRoot
$env:RUNNER_TEMP=Join-Path $PSScriptRoot 'temp'
$env:TEMP=$env:RUNNER_TEMP
$env:TMP=$env:RUNNER_TEMP
$principal=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Expected a standard user'}
& ./scripts/windows_install_check.ps1 -Installer (Join-Path $PSScriptRoot 'setup.exe')
'@ | Set-Content -Encoding UTF8 $entry
    & icacls.exe $sandbox /grant "*$($account.SID.Value):(OI)(CI)M" /T /Q | Out-Null
    if($LASTEXITCODE -ne 0){throw 'Cannot grant access to isolated CI test folder'}
    $credential=New-Object System.Management.Automation.PSCredential("$env:COMPUTERNAME\$name",$secret)
    $stdout=Join-Path $env:RUNNER_TEMP ($name+'.stdout.txt')
    $stderr=Join-Path $env:RUNNER_TEMP ($name+'.stderr.txt')
    $powershell=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $process=Start-Process $powershell -Credential $credential -LoadUserProfile -WorkingDirectory $sandbox -ArgumentList "-NoProfile -NonInteractive -File `"$entry`"" -RedirectStandardOutput $stdout -RedirectStandardError $stderr -Wait -PassThru
    Get-Content $stdout,$stderr | Write-Output
    Get-ChildItem (Join-Path $sandbox 'evidence') -File | Copy-Item -Destination $evidenceRoot -Force
    if($process.ExitCode -ne 0){throw "Standard-user verification failed (exit $($process.ExitCode))"}
    $result=Get-Content (Join-Path $evidenceRoot 'windows-install-check.json') | ConvertFrom-Json
    if(!$result.passed){throw 'Standard-user receipt did not pass'}
    Write-Output 'Standard-user installation, COM activation, Ribbon and uninstall verified.'
} finally {
    if($account){Remove-LocalUser -Name $name}
    if(Test-Path $sandbox){Remove-Item -LiteralPath $sandbox -Recurse -Force}
}
