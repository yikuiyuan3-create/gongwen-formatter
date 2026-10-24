#!/usr/bin/env python3
"""Independently authored, fictional DOCX acceptance samples; no private input."""
import base64
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('w', W)
ET.register_namespace('r', R)

TEXTS = [
    '关于开展办公文档整理工作的通知',
    '各示例部门：',
    '为规范内部文档管理，现将有关事项通知如下。样例内容完全虚构，仅用于软件测试。',
    '一、工作安排',
    '（一）整理范围',
    '1. 收集公开的制度文件。',
    '（1）保留原始文件，另存整理后的版本。',
    '本次演示预算为1234.56元，编号为TEST-2026-001，英文缩写为DOCX。',
    '附件：示例清单（虚构）',
    '示例公司办公室（虚构）',
    '2026年10月1日',
]

def paragraph(text):
    p = ET.Element(f'{{{W}}}p')
    props = ET.SubElement(p, f'{{{W}}}pPr')
    ET.SubElement(props, f'{{{W}}}spacing', {f'{{{W}}}line': '240', f'{{{W}}}lineRule':'auto'})
    run = ET.SubElement(p, f'{{{W}}}r')
    rp = ET.SubElement(run, f'{{{W}}}rPr')
    ET.SubElement(rp, f'{{{W}}}rFonts', {f'{{{W}}}eastAsia':'微软雅黑', f'{{{W}}}ascii':'Arial'})
    ET.SubElement(rp, f'{{{W}}}sz', {f'{{{W}}}val':'24'})
    ET.SubElement(run, f'{{{W}}}t').text = text
    return p

def create(path, title=None, table=False, protection=False, revision=False, multisection=False, numbering=False):
    doc = ET.Element(f'{{{W}}}document')
    body = ET.SubElement(doc, f'{{{W}}}body')
    for i, text in enumerate(TEXTS):
        body.append(paragraph(title if i == 0 and title else text))
    if numbering:
        for i in [2,7]:
            props=list(body)[i].find(f'{{{W}}}pPr')
            num=ET.SubElement(props,f'{{{W}}}numPr')
            ET.SubElement(num,f'{{{W}}}ilvl',{f'{{{W}}}val':'0'})
            ET.SubElement(num,f'{{{W}}}numId',{f'{{{W}}}val':'1'})
            ET.SubElement(props,f'{{{W}}}ind',{f'{{{W}}}left':'720',f'{{{W}}}hanging':'360'})
    if table:
        tbl = ET.SubElement(body, f'{{{W}}}tbl')
        tp = ET.SubElement(tbl, f'{{{W}}}tblPr')
        ET.SubElement(tp, f'{{{W}}}tblW', {f'{{{W}}}w':'0', f'{{{W}}}type':'auto'})
        grid = ET.SubElement(tbl, f'{{{W}}}tblGrid')
        for _ in range(2): ET.SubElement(grid, f'{{{W}}}gridCol', {f'{{{W}}}w':'3500'})
        for cells in [('示例事项','金额（虚构）'),('文档整理','1234.56')]:
            row = ET.SubElement(tbl, f'{{{W}}}tr')
            for cell in cells:
                tc = ET.SubElement(row, f'{{{W}}}tc')
                tc.append(paragraph(cell))
    if revision:
        p = list(body)[2]
        ins = ET.SubElement(p, f'{{{W}}}ins', {f'{{{W}}}id':'1',f'{{{W}}}author':'虚构测试',f'{{{W}}}date':'2026-10-01T00:00:00Z'})
        r = ET.SubElement(ins,f'{{{W}}}r');ET.SubElement(r,f'{{{W}}}t').text='修订内容（虚构）'
    if multisection:
        pp = list(body)[1].find(f'{{{W}}}pPr')
        ET.SubElement(pp, f'{{{W}}}sectPr')
    sect = ET.SubElement(body, f'{{{W}}}sectPr')
    ET.SubElement(sect, f'{{{W}}}pgSz', {f'{{{W}}}w':'11906',f'{{{W}}}h':'16838'})
    ET.SubElement(sect, f'{{{W}}}pgMar', {f'{{{W}}}top':'1440',f'{{{W}}}bottom':'1440',f'{{{W}}}left':'1440',f'{{{W}}}right':'1440',f'{{{W}}}header':'851',f'{{{W}}}footer':'992',f'{{{W}}}gutter':'0'})
    types='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
    rels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    document_rels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''
    settings=ET.Element(f'{{{W}}}settings')
    if protection:ET.SubElement(settings,f'{{{W}}}documentProtection',{f'{{{W}}}edit':'readOnly',f'{{{W}}}enforcement':'1'})
    if numbering:
        types=types.replace('</Types>',f'<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
        document_rels=document_rels.replace('</Relationships>','<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',types)
        z.writestr('_rels/.rels',rels)
        z.writestr('word/document.xml',ET.tostring(doc,encoding='utf-8',xml_declaration=True))
        z.writestr('word/_rels/document.xml.rels',document_rels)
        z.writestr('word/styles.xml',f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
        z.writestr('word/settings.xml',ET.tostring(settings,encoding='utf-8',xml_declaration=True))
        if numbering:
            z.writestr('word/numbering.xml',f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

def main():
    samples=ROOT/'samples';samples.mkdir(exist_ok=True)
    create(samples/'01_通知_虚构示例.docx')
    create(samples/'02_请示_虚构示例.docx','关于办公文档整理方案的请示')
    create(samples/'03_报告_虚构示例.docx','关于办公文档整理情况的报告')
    create(samples/'04_含表格_需要人工复核.docx',table=True)
    create(samples/'07_编辑保护与自动编号_虚构示例.docx',protection=True,numbering=True)
    fixtures=ROOT/'work/fixtures';fixtures.mkdir(parents=True,exist_ok=True)
    create(fixtures/'protected.docx',protection=True)
    create(fixtures/'revisions.docx',revision=True)
    create(fixtures/'multisection.docx',multisection=True)
    print('Created 5 fictional handoff samples and 3 restricted fixtures.')

if __name__=='__main__':main()
