#!/usr/bin/env python3
"""Copy a verified release into a fresh exFAT USB folder; full readback."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import plistlib
import shutil
import subprocess
import tempfile
import time
from zipfile import ZipFile
from release_config import ROOT,RELEASE_NAME,VERSION

def hash_file(path):
    h=hashlib.sha256()
    with path.open('rb') as stream:
        while chunk:=stream.read(1024*1024):h.update(chunk)
    return h.hexdigest()

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--volume',required=True)
    parser.add_argument('--outputs',default=str(ROOT/'work/releases'))
    args=parser.parse_args()
    outputs=Path(args.outputs).resolve();release=outputs/RELEASE_NAME;volume=Path(args.volume).resolve()
    data=plistlib.loads(subprocess.run(['diskutil','info','-plist',str(volume)],capture_output=True,check=True).stdout)
    assert data['MountPoint']==str(volume) and data['FilesystemType']=='exfat' and data['Internal'] is False and data['BusProtocol']=='USB' and data['WritableVolume'] is True,'Unexpected delivery medium'
    for line in (release/'SHA256SUMS.txt').read_text().splitlines():
        digest,relative=line.split('  ',1);assert hash_file(release/relative)==digest,'Local checksum failed: '+relative
    with ZipFile(release/f'源代码_公文一键排版_v{VERSION}.zip') as z:assert z.testzip() is None
    files=sorted(p for p in release.rglob('*') if p.is_file());total=sum(p.stat().st_size for p in files)
    assert shutil.disk_usage(volume).free>total+10_000_000,'Not enough USB space'
    folder=volume/RELEASE_NAME
    if folder.exists():folder=volume/(RELEASE_NAME+'_'+time.strftime('%H%M%S'))
    folder.mkdir(exist_ok=False)
    for source in files:
        dest=folder/source.relative_to(release);dest.parent.mkdir(parents=True,exist_ok=True)
        with source.open('rb') as input_stream,dest.open('xb') as target:
            shutil.copyfileobj(input_stream,target,1024*1024);target.flush();os.fsync(target.fileno())
    expected={str(f.relative_to(release)) for f in files}
    extras=[f for f in folder.rglob('*') if f.is_file() and str(f.relative_to(folder)) not in expected]
    outer=volume/('._'+folder.name)
    if outer.exists():extras.append(outer)
    backup=Path(tempfile.mkdtemp(prefix='usb-generated-metadata-',dir=ROOT/'work'))
    for extra in extras:
        # Remove only system-generated AppleDouble for this freshly created delivery.
        raw=extra.read_bytes();assert extra.name.startswith('._') and raw[:4]==bytes.fromhex('00051607'),'Unexpected extra file preserved'
        dest=backup/extra.relative_to(volume);dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(raw)
        assert dest.read_bytes()==raw;extra.unlink()
    os.sync()
    records=[]
    for source in files:
        relative=str(source.relative_to(release));dest=folder/relative;expected_hash=hash_file(source);actual=hash_file(dest)
        assert source.stat().st_size==dest.stat().st_size and expected_hash==actual,'USB readback failed: '+relative
        records.append({'path':relative,'bytes':dest.stat().st_size,'sha256':actual,'status':'matched'})
    assert {str(f.relative_to(folder)) for f in folder.rglob('*') if f.is_file()}==expected,'Unexpected USB delivery entries'
    report={'status':'passed','copied_file_count':len(records),'total_bytes':total,'source_folder':str(release),'usb_folder':str(folder),'media':{'name':data['VolumeName'],'format':data['FilesystemName'],'external_usb':True},'method':'Fresh folder; flushed writes; full size and SHA-256 USB readback','generated_mac_metadata_cleaned':len(extras),'existing_usb_data_modified':False,'windows_runtime_verified':False,'wps_word_compatibility_verified':False,'files':records}
    receipt=outputs/(RELEASE_NAME+'_U盘交付校验.json');receipt.write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps({k:v for k,v in report.items() if k!='files'},ensure_ascii=False,indent=2))
    print('RECEIPT',receipt)

if __name__=='__main__':main()
