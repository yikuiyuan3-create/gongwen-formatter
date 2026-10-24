Unicode true
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "WinVer.nsh"
!include "x64.nsh"
!ifndef APPVERSION
!define APPVERSION "0.4.0"
!endif
!ifndef PAYLOAD
!define PAYLOAD "../work/publish/win-x64-${APPVERSION}"
!endif
!ifndef OUTPUT
!define OUTPUT "../work/公文一键排版_Setup_${APPVERSION}_x64.exe"
!endif
Name "公文一键排版"
OutFile "${OUTPUT}"
InstallDir "$LOCALAPPDATA\Programs\GongwenFormatter"
InstallDirRegKey HKCU "Software\GongwenFormatter" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma
!include "office-addin.nsh"
BrandingText "公文一键排版 · 本地离线处理"
VIProductVersion "${APPVERSION}.0"
VIAddVersionKey /LANG=2052 "ProductName" "公文一键排版"
VIAddVersionKey /LANG=2052 "FileDescription" "公文一键排版 Windows x64 离线安装程序"
VIAddVersionKey /LANG=2052 "FileVersion" "${APPVERSION}"
VIAddVersionKey /LANG=2052 "ProductVersion" "${APPVERSION}"
VIAddVersionKey /LANG=2052 "LegalCopyright" "公文一键排版"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "安装公文一键排版"
!define MUI_WELCOMEPAGE_TEXT "安装WPS/Word公文排版工具栏插件和独立排版窗口。先关闭WPS、Word和旧版工具。$\r$\n$\r$\n本地离线处理、不上传正文、原稿保留。MIT免费开源。工具栏需要.NET Framework 4.8及允许COM的办公软件版本；独立窗口自带.NET 10运行组件。$\r$\n$\r$\n适用于64位Windows。请先用虚构示例完成目标WPS/Word验收。"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "INSTALL-NOTICE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TITLE "安装文件已写入"
!define MUI_FINISHPAGE_TEXT "重新打开WPS文字或Word，查找“公文排版”工具栏。也可从桌面打开独立窗口。$\r$\n$\r$\n先将文档保存为DOCX并按Ctrl+S，点击“排版当前文档”；在工具窗口导出新的DOCX。先用虚构示例验收。"
!define MUI_FINISHPAGE_RUN "$INSTDIR\app-${APPVERSION}\GongwenFormatter.exe"
!define MUI_FINISHPAGE_RUN_TEXT "打开公文一键排版"
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "SimpChinese"

Function .onInit
  ${IfNot} ${AtLeastWin10}
    MessageBox MB_ICONSTOP "此安装包需要 Windows 10 或 Windows 11。" /SD IDOK
    Abort
  ${EndIf}
  ${IfNot} ${RunningX64}
    MessageBox MB_ICONSTOP "此安装包需要64位 Windows。32位系统请使用专门适配的安装包。" /SD IDOK
    Abort
  ${EndIf}
  SetRegView 64
  SetShellVarContext current
  ReadRegDWORD $0 HKLM "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"
  ${If} $0 < 528040
    MessageBox MB_ICONSTOP "工具栏插件需要.NET Framework 4.8或更新4.x。请由单位IT从微软官方配置后重试。" /SD IDOK
    Abort
  ${EndIf}
  Call CheckOfficeClosed
  ReadRegStr $2 HKLM "SOFTWARE\Microsoft\Windows NT\CurrentVersion" "CurrentBuildNumber"
  ${If} $2 < 14393
    MessageBox MB_ICONSTOP "当前 Windows 系统版本过早，无法安装。请将系统版本信息交给维护人员。" /SD IDOK
    Abort
  ${EndIf}
  System::Call 'kernel32::OpenMutexW(i 0x100000, i 0, w "Local\GongwenFormatter.SingleInstance") p.r0'
  ${If} $0 != 0
    System::Call 'kernel32::CloseHandle(p r0)'
    MessageBox MB_ICONEXCLAMATION "请先退出公文一键排版，再安装或更新。" /SD IDOK
    Abort
  ${EndIf}
FunctionEnd

Section "公文一键排版" SEC_MAIN
  SetOutPath "$INSTDIR\app-${APPVERSION}"
  ClearErrors
  File /r "${PAYLOAD}/*"
  ${If} ${Errors}
    MessageBox MB_ICONSTOP "组件未完整写入，安装尚未完成。请检查空间和权限后重试。原稿未被修改。" /SD IDOK
    Abort
  ${EndIf}
  Call RegisterOfficeAddin
  SetOutPath "$INSTDIR"
  File /oname=安装须知.txt "INSTALL-NOTICE.txt"
  WriteUninstaller "$INSTDIR\卸载公文一键排版.exe"
  WriteRegStr HKCU "Software\GongwenFormatter" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "DisplayName" "公文一键排版"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "DisplayVersion" "${APPVERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "Publisher" "公文一键排版"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "DisplayIcon" "$INSTDIR\app-${APPVERSION}\GongwenFormatter.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "UninstallString" '$\"$INSTDIR\卸载公文一键排版.exe$\"'
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter" "NoRepair" 1
  CreateDirectory "$SMPROGRAMS\公文一键排版"
  CreateShortcut "$SMPROGRAMS\公文一键排版\公文一键排版.lnk" "$INSTDIR\app-${APPVERSION}\GongwenFormatter.exe"
  CreateShortcut "$SMPROGRAMS\公文一键排版\使用说明.lnk" "$INSTDIR\app-${APPVERSION}\docs\使用说明.html"
  CreateShortcut "$DESKTOP\公文一键排版.lnk" "$INSTDIR\app-${APPVERSION}\GongwenFormatter.exe"
SectionEnd

Function un.onInit
  SetRegView 64
  SetShellVarContext current
  Call un.CheckOfficeClosed
  System::Call 'kernel32::OpenMutexW(i 0x100000, i 0, w "Local\GongwenFormatter.SingleInstance") p.r0'
  ${If} $0 != 0
    System::Call 'kernel32::CloseHandle(p r0)'
    MessageBox MB_ICONEXCLAMATION "请先退出公文一键排版，再卸载。" /SD IDOK
    Abort
  ${EndIf}
FunctionEnd
Section "Uninstall"
  !insertmacro UnregisterAddinView 32
  !insertmacro UnregisterAddinView 64
  ; Exact list of installed files, preserve user documents and unknown files.
  !include "uninstall-files.nsh"
  ; If upgraded from v0.1, remove only its known installed files as well.
  !include "uninstall-files-v0.1.0.nsh"
  !include "uninstall-files-v0.2.0.nsh"
  !include "uninstall-files-v0.3.0.nsh"
  Delete "$INSTDIR\安装须知.txt"
  Delete "$INSTDIR\卸载公文一键排版.exe"
  Delete "$SMPROGRAMS\公文一键排版\公文一键排版.lnk"
  Delete "$SMPROGRAMS\公文一键排版\使用说明.lnk"
  RMDir "$SMPROGRAMS\公文一键排版"
  Delete "$DESKTOP\公文一键排版.lnk"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\GongwenFormatter"
  DeleteRegKey HKCU "Software\GongwenFormatter"
  RMDir "$INSTDIR"
  MessageBox MB_ICONINFORMATION "程序已卸载。原稿、排版输出和本地模板配置已保留。" /SD IDOK
SectionEnd
