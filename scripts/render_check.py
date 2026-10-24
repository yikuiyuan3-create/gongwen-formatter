#!/usr/bin/env python3
"""Supplementary macOS document rendering; never WPS/Word acceptance."""
import json
import hashlib
import os
from pathlib import Path
import re
import subprocess
import shutil
from release_config import ROOT,VERSION

def main():
    tests=json.loads((ROOT/'evidence/core-tests.json').read_text())
    assert tests['failed']==0,'Requires a fresh passing test output'
    tested=Path(tests['outputDirectory'])
    qa=ROOT/f'work/render-font-check-v{VERSION}';qa.mkdir(exist_ok=False)
    profile=qa/'lo-profile';profile.mkdir()
    config=qa/'fonts.conf'
    config.write_text('<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd"><fontconfig><dir>/System/Library/Fonts</dir><dir>/Library/Fonts</dir><cachedir>'+str(qa/'font-cache')+'</cachedir><alias><family>仿宋</family><prefer><family>Songti SC</family></prefer></alias><alias><family>方正小标宋简体</family><prefer><family>Songti SC</family></prefer></alias><alias><family>黑体</family><prefer><family>Heiti SC</family></prefer></alias><alias><family>楷体</family><prefer><family>Songti SC</family></prefer></alias><alias><family>微软雅黑</family><prefer><family>Heiti SC</family></prefer></alias><alias><family>sans-serif</family><prefer><family>Heiti SC</family></prefer></alias></fontconfig>')
    env=os.environ.copy();env['FONTCONFIG_FILE']=str(config)
    soffice=os.environ.get('SOFFICE') or shutil.which('soffice')
    if not soffice:raise SystemExit('Set SOFFICE to a LibreOffice executable; supplemental rendering is optional.')
    inputs=[tested/(stem+'.docx') for stem in ['formatted','edited-body','table-acknowledged','protected-edited','protected-numbered-shipping-output','numbered-tabs-output']]
    result=subprocess.run([soffice,'-env:UserInstallation='+profile.as_uri(),'--headless','--convert-to','pdf','--outdir',str(qa),*map(str,inputs)],capture_output=True,text=True,env=env,timeout=60,check=True)
    (ROOT/f'evidence/render-v{VERSION}.log').write_text(result.stdout+result.stderr)
    records=[]
    for source in inputs:
        pdf=qa/(source.stem+'.pdf')
        info=subprocess.run(['pdfinfo',str(pdf)],check=True,capture_output=True,text=True).stdout
        pages=int(re.search(r'^Pages:\s+(\d+)',info,re.M).group(1))
        subprocess.run(['pdftoppm','-png','-scale-to','1600',str(pdf),str(qa/source.stem)],check=True,capture_output=True)
        records.append({'sample':source.stem,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'pdf_sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'pdf_pages':pages,'page_size':re.search(r'^Page size:\s+(.+)',info,re.M).group(1),'images':[f'{source.stem}-{i}.png' for i in range(1,pages+1)],'visual_inspection':'pending'})
    report={'renderer':'LibreOfficeDev 26.8.0.0.alpha0 on macOS','render_completed':True,'visual_inspection_completed':False,'fonts':'Explicit macOS Songti SC / Heiti SC substitutions; not Windows required fonts','wps_word_compatibility_verified':False,'records':records}
    (ROOT/'evidence/渲染检查.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
