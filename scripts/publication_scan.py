#!/usr/bin/env python3
"""Whitelist public source; reject private paths, credentials and non-fictional samples."""
from pathlib import Path
import re
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
FOLDERS=('src','tests','scripts','packaging','docs','samples','.github')
ROOT_FILES=('.gitignore','Directory.Build.props','README.md','LICENSE','NOTICE.md','SECURITY.md')
EXCLUDED={'bin','obj','work','evidence','.git','__pycache__','.DS_Store'}

def source_files():
    for name in ROOT_FILES:
        p=ROOT/name
        if p.is_file():yield p
    for name in FOLDERS:
        for p in sorted((ROOT/name).rglob('*')):
            if p.is_file() and not any(part in EXCLUDED or part.startswith('._') for part in p.relative_to(ROOT).parts):
                assert not p.is_symlink(),'Symlink outside public source contract'
                yield p

def scan():
    findings=[];count=0
    markers=['/'+'Users/','/'+'Volumes/','wxid'+'_','xwechat'+'_files']
    secrets=[r'gh[pousr]_[A-Za-z0-9]{20,}',r'github_pat_[A-Za-z0-9_]{20,}',r'sk-proj-[A-Za-z0-9_-]{20,}',r'AKIA[A-Z0-9]{16}',r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----']
    for p in source_files():
        count+=1;relative=str(p.relative_to(ROOT))
        if p.suffix=='.docx':
            assert p.parent==ROOT/'samples','Only independently generated fictional samples may be published'
            with ZipFile(p) as z:
                text=z.read('word/document.xml').decode('utf-8')
                if '虚构' not in text:findings.append((relative,'sample lacks fictional marker'))
            continue
        try:text=p.read_text(encoding='utf-8')
        except UnicodeError:findings.append((relative,'unexpected binary'));continue
        if any(marker in text for marker in markers):findings.append((relative,'private machine/source path'))
        if any(re.search(pattern,text) for pattern in secrets):findings.append((relative,'credential-like material'))
    if findings:
        # Paths and category only; never echo matching secret values.
        for file,category in findings:print('FAIL',file,category)
        raise SystemExit(1)
    print(f'PASS public source whitelist: {count} files; no private path or credential patterns; samples fictional')
    return count

if __name__=='__main__':scan()
