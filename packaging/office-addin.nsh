; Only this product's per-user registration. Never change Office security policy.
!define ADDIN_PROGID "Gongwen.OfficeAddin.Connect"
!define ADDIN_CLSID "{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}"
!define ADDIN_CLASS "Gongwen.OfficeAddin.Connect"
!define ADDIN_ASSEMBLY "Gongwen.OfficeAddin, Version=${APPVERSION}.0, Culture=neutral, PublicKeyToken=null"
Var AddinCodeBase

!macro RegisterAddinView VIEW
  SetRegView ${VIEW}
  WriteRegStr HKCU "Software\Classes\${ADDIN_PROGID}" "" "公文一键排版"
  WriteRegStr HKCU "Software\Classes\${ADDIN_PROGID}\CLSID" "" "${ADDIN_CLSID}"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}" "" "公文一键排版"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\ProgId" "" "${ADDIN_PROGID}"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "" "mscoree.dll"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "ThreadingModel" "Both"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "Class" "${ADDIN_CLASS}"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "Assembly" "${ADDIN_ASSEMBLY}"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "RuntimeVersion" "v4.0.30319"
  WriteRegStr HKCU "Software\Classes\CLSID\${ADDIN_CLSID}\InprocServer32" "CodeBase" "$AddinCodeBase"
  WriteRegStr HKCU "Software\Microsoft\Office\Word\Addins\${ADDIN_PROGID}" "FriendlyName" "公文一键排版"
  WriteRegStr HKCU "Software\Microsoft\Office\Word\Addins\${ADDIN_PROGID}" "Description" "本地公文排版，另存新DOCX，原稿保留"
  WriteRegDWORD HKCU "Software\Microsoft\Office\Word\Addins\${ADDIN_PROGID}" "LoadBehavior" 3
  WriteRegStr HKCU "Software\Kingsoft\Office\WPS\Addins\${ADDIN_PROGID}" "FriendlyName" "公文一键排版"
  WriteRegStr HKCU "Software\Kingsoft\Office\WPS\Addins\${ADDIN_PROGID}" "Description" "本地公文排版，另存新DOCX，原稿保留"
  WriteRegDWORD HKCU "Software\Kingsoft\Office\WPS\Addins\${ADDIN_PROGID}" "LoadBehavior" 3
!macroend

!macro UnregisterAddinView VIEW
  SetRegView ${VIEW}
  DeleteRegKey HKCU "Software\Microsoft\Office\Word\Addins\${ADDIN_PROGID}"
  DeleteRegKey HKCU "Software\Kingsoft\Office\WPS\Addins\${ADDIN_PROGID}"
  DeleteRegKey HKCU "Software\Classes\CLSID\${ADDIN_CLSID}"
  DeleteRegKey HKCU "Software\Classes\${ADDIN_PROGID}"
!macroend

Function RegisterOfficeAddin
  ; URL-encode the real path, including spaces, Chinese characters and '#'.
  System::Call 'shlwapi::UrlCreateFromPathW(w "$INSTDIR\app-${APPVERSION}\office-addin\Gongwen.OfficeAddin.dll", w .r0, *i ${NSIS_MAX_STRLEN} r1, i 0) i.r2'
  ${If} $2 != 0
    MessageBox MB_ICONSTOP "无法创建加载项路径，安装未完成。请选择本机文件夹后重试。" /SD IDOK
    Abort
  ${EndIf}
  StrCpy $AddinCodeBase $0
  ClearErrors
  !insertmacro RegisterAddinView 32
  !insertmacro RegisterAddinView 64
  ${If} ${Errors}
    !insertmacro UnregisterAddinView 32
    !insertmacro UnregisterAddinView 64
    MessageBox MB_ICONSTOP "当前用户的插件注册未完成，请检查单位权限策略后重试。" /SD IDOK
    Abort
  ${EndIf}
FunctionEnd

; NSIS is x86 Unicode: PROCESSENTRY32W is 556 bytes, szExeFile offset 36.
; Native read-only enumeration avoids PowerShell and third-party process plugins.
!macro DefineOfficeCheck PREFIX
Function ${PREFIX}CheckOfficeClosed
  System::Call 'kernel32::CreateToolhelp32Snapshot(i 2, i 0) p.r0'
  ${If} $0 == -1
    MessageBox MB_ICONSTOP "无法检查办公软件状态，请关闭WPS与Word后重试。" /SD IDOK
    Abort
  ${EndIf}
  System::Alloc 556
  Pop $1
  ${If} $1 == 0
    System::Call 'kernel32::CloseHandle(p r0)'
    Abort
  ${EndIf}
  System::Call '*$1(i 556)'
  System::Call 'kernel32::Process32FirstW(p r0, p r1) i.r2'
  loop:
    ${If} $2 == 0
      Goto done
    ${EndIf}
    IntOp $3 $1 + 36
    System::Call '*$3(&w260 .r4)'
    ${If} $4 == "WINWORD.EXE"
    ${OrIf} $4 == "wps.exe"
      System::Free $1
      System::Call 'kernel32::CloseHandle(p r0)'
      MessageBox MB_ICONEXCLAMATION "请先关闭WPS文字和Microsoft Word，再安装、更新或卸载插件。" /SD IDOK
      Abort
    ${EndIf}
    System::Call 'kernel32::Process32NextW(p r0, p r1) i.r2'
    Goto loop
  done:
  System::Free $1
  System::Call 'kernel32::CloseHandle(p r0)'
FunctionEnd
!macroend
!insertmacro DefineOfficeCheck ""
!insertmacro DefineOfficeCheck "un."
