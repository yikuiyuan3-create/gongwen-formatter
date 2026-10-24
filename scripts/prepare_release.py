#!/usr/bin/env python3
"""Create a portable public release from fresh test and package verification."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
from zipfile import ZipFile,ZIP_DEFLATED
from release_config import ROOT,VERSION,INSTALLER,RELEASE_NAME,RELEASE_DATE
from publication_scan import source_files,scan

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--outputs',type=Path,default=ROOT/'work/releases')
    args=parser.parse_args()
    core=json.loads((ROOT/'evidence/core-tests.json').read_text())
    checks=json.loads((ROOT/'evidence/release-checks.json').read_text())
    package=json.loads((ROOT/'evidence/installer-verification.json').read_text())
    assert core['failed']==0 and checks['version']==VERSION and package['passed']
    assert all(v['failed']==0 for v in checks['checks'].values())
    digest=hashlib.sha256(INSTALLER.read_bytes()).hexdigest()
    assert digest==package['installer_sha256'],'Installer changed after verification'
    scan()
    destination=args.outputs/RELEASE_NAME;destination.mkdir(parents=True,exist_ok=False)
    shutil.copyfile(INSTALLER,destination/INSTALLER.name)
    for name in ['LICENSE','NOTICE.md','README.md']:
        shutil.copyfile(ROOT/name,destination/name)
    shutil.copyfile(ROOT/'docs/插件安装与使用.md',destination/'安装与使用.md')
    source=destination/f'gongwen-formatter-v{VERSION}-source.zip'
    with ZipFile(source,'w',ZIP_DEFLATED) as z:
        for p in sorted(source_files()):z.write(p,'gongwen-formatter/'+str(p.relative_to(ROOT)))
    with ZipFile(destination/'虚构示例.zip','w',ZIP_DEFLATED) as z:
        for p in sorted((ROOT/'samples').glob('*.docx')):z.write(p,p.name)
        tested=Path(core['outputDirectory'])
        z.write(tested/'formatted.docx','05_通知_排版结果.docx')
        z.write(tested/'protected-numbered-shipping-output.docx','08_保护编号_编辑排版结果.docx')
    evidence={'version':VERSION,'date':RELEASE_DATE,'checks':checks['checks'],'package':{k:v for k,v in package.items() if k!='inspection_directory'},'wps_word_toolbar_verified':False,'windows_runtime_verified':False}
    (destination/'verification.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2))
    counts='；'.join(f"{k}: {v['passed']}通过/{v['failed']}失败" for k,v in checks['checks'].items())
    report=f'''# 公文一键排版 v{VERSION} 公开发布验证记录

日期：{RELEASE_DATE}。许可：MIT。状态：测试版。

- 新增 .NET Framework 4.8 AnyCPU COM工具栏，Word/WPS共用入口；只读活动文档的保存状态与路径，不调用Save/SaveAs。
- 新增 --open 与仅当前用户、本机会话IPC；已有窗口接收请求，保留未导出编辑、模态和忙碌边界。
- 实际统一检查：{counts}。其中IPC检查使用真实本机管道；COM策略检查使用虚构文件与宿主替身，不冒充Office执行。
- Release发布与自包含Windows x64打包成功，运行组件{checks['runtime_version']}；{package['verified_payload_files']}个载荷文件逐项核验长度与SHA-256。
- 安装包大小{package['installer_bytes']:,}字节，SHA-256：`{digest}`。
- 公开源码使用明确白名单，扫描本机路径和凭证模式，仅附虚构示例。
- Windows安装、32/64位COM激活、特殊字符安装路径及卸载以GitHub Actions的独立Windows检查为准，不将macOS构建当作Windows运行证明。
- 尚未完成真实WPS/Word工具栏按钮、不同宿主版本、分页/字体/打印与双端另存验收；不支持的COM环境不得视为已兼容。
- 安装器/应用未商业代码签名；不附商业字体，不破解打开密码，不改变单位安全策略。

安装与使用见同目录文档。源码和下载：https://github.com/yikuiyuan3-create/gongwen-formatter
'''
    (destination/'验证报告.md').write_text(report)
    files=sorted(p for p in destination.iterdir() if p.is_file())
    (destination/'SHA256SUMS.txt').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in files))
    print(destination);print('Public release files:',len(files)+1)

if __name__=='__main__':main()
