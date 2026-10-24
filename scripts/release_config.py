from pathlib import Path
import xml.etree.ElementTree as ET
import os
from datetime import date,datetime

ROOT=Path(__file__).resolve().parents[1]
VERSION=ET.parse(ROOT/'Directory.Build.props').findtext('./PropertyGroup/Version')
assert VERSION and all(part.isdigit() for part in VERSION.split('.'))
INSTALLER_NAME=f'公文一键排版_Setup_{VERSION}_x64.exe'
INSTALLER=ROOT/'work'/INSTALLER_NAME
PAYLOAD=ROOT/'work/publish'/f'win-x64-{VERSION}'
RELEASE_DATE=os.environ.get('GONGWEN_RELEASE_DATE',date.today().strftime('%Y%m%d'))
datetime.strptime(RELEASE_DATE,'%Y%m%d')
RELEASE_NAME=f'公文一键排版_Windows_v{VERSION}_{RELEASE_DATE}'
