#!/usr/bin/env python3
"""Local test, offline Windows publish, exact uninstall manifest, NSIS build."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import re
from release_config import VERSION,INSTALLER,PAYLOAD

ROOT = Path(__file__).resolve().parents[1]

def run(cmd):
    print('RUN', ' '.join(map(str,cmd)), flush=True)
    env = os.environ.copy()
    env.update(DOTNET_CLI_TELEMETRY_OPTOUT='1', DOTNET_NOLOGO='1', DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1')
    result=subprocess.run(list(map(str,cmd)),cwd=ROOT,env=env,text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
    print(result.stdout,flush=True)
    result.check_returncode()
    return result.stdout

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--dotnet',required=True)
    parser.add_argument('--runtime-version',default='10.0.13')
    args=parser.parse_args()
    run([sys.executable,ROOT/'scripts/create_samples.py'])
    checks={}
    for name,project,extra in [('core','Gongwen.Tests',['--',ROOT]),('office_addin','Gongwen.OfficeAddin.Tests',[]),('local_ipc','Gongwen.Integration.Tests',[])]:
        log=run([args.dotnet,'run','--project',ROOT/'tests'/project/(project+'.csproj'),'-c','Release',*extra])
        matched=re.search(r'RESULT (\d+) passed, (\d+) failed',log)
        assert matched and int(matched[2])==0,'Missing passing test summary: '+name
        checks[name]={'passed':int(matched[1]),'failed':int(matched[2])}
    payload=PAYLOAD
    if payload.exists():raise SystemExit('Publish directory already exists; use a fresh staging path after inspecting the previous output.')
    run([args.dotnet,'publish',ROOT/'src/Gongwen.Desktop/Gongwen.Desktop.csproj','-c','Release','-r','win-x64','--self-contained','true','-o',payload,'-p:RuntimeFrameworkVersion='+args.runtime_version,'-p:PublishReadyToRun=false','-p:DebugType=None','-p:DebugSymbols=false'])
    run([args.dotnet,'build',ROOT/'src/Gongwen.OfficeAddin/Gongwen.OfficeAddin.csproj','-c','Release','-o',payload/'office-addin','-p:DebugType=None','-p:DebugSymbols=false'])
    for name in ['LICENSE','NOTICE.md']:
        shutil.copyfile(ROOT/name,payload/name)
    for name in ['samples','docs']:
        if (ROOT/name).is_dir():shutil.copytree(ROOT/name,payload/name)
    packages=Path.home()/'.nuget/packages'
    runtime_notices={
        'microsoft.netcore.app.runtime.win-x64':['LICENSE.TXT','THIRD-PARTY-NOTICES.TXT'],
        'microsoft.windowsdesktop.app.runtime.win-x64':['LICENSE']
    }
    for package,names in runtime_notices.items():
        license_dir=payload/'licenses'/package
        license_dir.mkdir(parents=True,exist_ok=True)
        for name in names:
            source=packages/package/args.runtime_version/name
            if not source.is_file():raise SystemExit('Runtime license missing: '+str(source))
            shutil.copy2(source,license_dir/name)
    files=sorted(p for p in payload.rglob('*') if p.is_file())
    lines=['; Exact installed-file manifest; never recurse-delete user directories.']
    def escape(path):return str(path).replace('/','\\').replace('$','$$').replace('"','$\\"')
    for p in files:lines.append('Delete "$INSTDIR\\app-${APPVERSION}\\'+escape(p.relative_to(payload))+'"')
    for p in sorted((p for p in payload.rglob('*') if p.is_dir()),key=lambda p:len(p.parts),reverse=True):
        lines.append('RMDir "$INSTDIR\\app-${APPVERSION}\\'+escape(p.relative_to(payload))+'"')
    lines.append('RMDir "$INSTDIR\\app-${APPVERSION}"')
    (ROOT/'packaging/uninstall-files.nsh').write_text('\n'.join(lines)+'\n',encoding='utf-8-sig')
    # Windows makensis uses native file separators; avoid non-ASCII CLI output names.
    compiled_installer=ROOT/'work'/f'GongwenFormatter_Setup_{VERSION}_x64.exe'
    assert not compiled_installer.exists() and not INSTALLER.exists(),'Preserve the previous installer before rebuilding'
    run(['makensis','-V3','-DAPPVERSION='+VERSION,'-DPAYLOAD='+str(payload),'-DPAYLOAD_FILESPEC='+str(payload/'*'),'-DOUTPUT='+str(compiled_installer),ROOT/'packaging/installer.nsi'])
    compiled_installer.rename(INSTALLER)
    manifest=[{'path':str(p.relative_to(payload)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in files]
    (ROOT/'evidence/payload-manifest.json').write_text(json.dumps({'runtime_version':args.runtime_version,'files':manifest},ensure_ascii=False,indent=2))
    (ROOT/'evidence/release-checks.json').write_text(json.dumps({'version':VERSION,'runtime_version':args.runtime_version,'checks':checks,'windows_office_host_verified':False},ensure_ascii=False,indent=2))
    print('INSTALLER',INSTALLER)

if __name__=='__main__':main()
