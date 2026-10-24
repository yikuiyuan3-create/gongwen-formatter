#!/usr/bin/env python3
"""Inspect generated NSIS bytes without executing Windows code."""
import hashlib
import json
from pathlib import Path
import struct
import subprocess
import tempfile
import shutil
import platform
from release_config import INSTALLER

ROOT=Path(__file__).resolve().parents[1]

def pe(path):
    data=path.read_bytes()
    assert data[:2]==b'MZ', str(path)+' lacks DOS header'
    offset=struct.unpack_from('<I',data,0x3c)[0]
    assert data[offset:offset+4]==b'PE\0\0', str(path)+' lacks PE header'
    machine=struct.unpack_from('<H',data,offset+4)[0]
    optional=offset+24
    magic=struct.unpack_from('<H',data,optional)[0]
    subsystem=struct.unpack_from('<H',data,optional+68)[0]
    directories=optional+(112 if magic==0x20b else 96)
    cert_offset,cert_bytes=struct.unpack_from('<II',data,directories+4*8)
    clr_rva,clr_size=struct.unpack_from('<II',data,directories+14*8)
    clr_flags=None
    if clr_rva:
        count=struct.unpack_from('<H',data,offset+6)[0]
        table=optional+struct.unpack_from('<H',data,offset+20)[0]
        for index in range(count):
            row=table+index*40
            virtual_size,virtual_address,raw_size,raw_pointer=struct.unpack_from('<IIII',data,row+8)
            if virtual_address<=clr_rva<virtual_address+max(virtual_size,raw_size):
                clr_flags=struct.unpack_from('<I',data,raw_pointer+clr_rva-virtual_address+16)[0];break
    return {'machine':hex(machine),'optional_header':hex(magic),'subsystem':subsystem,'signed_certificate_bytes':cert_bytes,'certificate_offset':cert_offset,'clr_flags':clr_flags}

def main():
    installer=INSTALLER
    manifest=json.loads((ROOT/'evidence/payload-manifest.json').read_text())
    inspection=Path(tempfile.mkdtemp(prefix='installer-inspect-',dir=ROOT/'work'))
    sevenzip=shutil.which('7zz') or shutil.which('7z')
    assert sevenzip,'Install 7-Zip (7zz or 7z) for package verification'
    subprocess.run([sevenzip,'t',str(installer)],check=True,capture_output=True)
    subprocess.run([sevenzip,'x',str(installer),'-o'+str(inspection),'-y'],check=True,capture_output=True)
    apps=list(inspection.rglob('GongwenFormatter.exe'))
    assert len(apps)==1,'Expected exactly one desktop executable'
    payload=apps[0].parent
    matched=0
    for entry in manifest['files']:
        path=payload/entry['path']
        assert path.is_file(),'Installer file missing: '+entry['path']
        assert path.stat().st_size==entry['bytes'],'Installer size mismatch: '+entry['path']
        assert hashlib.sha256(path.read_bytes()).hexdigest()==entry['sha256'],'Installer hash mismatch: '+entry['path']
        matched+=1
    actual={str(p.relative_to(payload)) for p in payload.rglob('*') if p.is_file()}
    expected={entry['path'] for entry in manifest['files']}
    assert actual==expected,'Installer payload has unexpected files'
    images={name:pe(payload/name) for name in ['GongwenFormatter.exe','coreclr.dll','hostfxr.dll','hostpolicy.dll','wpfgfx_cor3.dll']}
    assert all(info['machine']=='0x8664' for info in images.values()),'Native Windows payload is not x64'
    assert images['GongwenFormatter.exe']['subsystem']==2,'Desktop executable not Windows GUI'
    config=json.loads((payload/'GongwenFormatter.runtimeconfig.json').read_text())
    included=config['runtimeOptions']['includedFrameworks']
    assert {x['name'] for x in included}=={'Microsoft.NETCore.App','Microsoft.WindowsDesktop.App'},'Missing self-contained framework'
    assert all(x['version']==manifest['runtime_version'] for x in included),'Wrong runtime version'
    addin=pe(payload/'office-addin/Gongwen.OfficeAddin.dll')
    assert addin['machine']=='0x14c' and addin['clr_flags'] is not None and addin['clr_flags']&1 and not addin['clr_flags']&2,'Expected AnyCPU IL-only, not 32-bit-required add-in'
    registrar=pe(payload/'office-addin/Gongwen.AddinRegistrar.exe')
    assert registrar['machine']=='0x14c' and registrar['clr_flags'] is not None and registrar['clr_flags']&1 and not registrar['clr_flags']&2,'Expected AnyCPU registrar'
    assert (payload/'LICENSE').is_file() and (payload/'NOTICE.md').is_file(),'Missing open-source license notices'
    report={
        'passed':True,'installer_sha256':hashlib.sha256(installer.read_bytes()).hexdigest(),
        'installer_bytes':installer.stat().st_size,'installer_pe':pe(installer),
        'verified_payload_files':matched,'payload_bytes':sum(e['bytes'] for e in manifest['files']),
        'native_images':images,'included_frameworks':included,
        'office_addin_pe':addin,'registrar_pe':registrar,'office_addin_host_verified':False,
        'verification':'NSIS archive test/extraction, exact payload bytes and PE/CLR headers on '+platform.system(),
        'windows_install_startup_verified':False,'wps_word_compatibility_verified':False,
        'inspection_directory':str(inspection)
    }
    (ROOT/'evidence/installer-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
