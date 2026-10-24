; Only this product's per-user registration. Never change Office security policy.
!define ADDIN_PROGID "Gongwen.OfficeAddin.Connect"
!define ADDIN_CLSID "{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}"
!define ADDIN_CLASS "Gongwen.OfficeAddin.Connect"
!define ADDIN_ASSEMBLY "Gongwen.OfficeAddin, Version=${APPVERSION}.0, Culture=neutral, PublicKeyToken=null"
Function RegisterOfficeAddin
  ClearErrors
  ExecWait '"$INSTDIR\app-${APPVERSION}\office-addin\Gongwen.AddinRegistrar.exe" --register' $0
  ${If} ${Errors}
  ${OrIf} $0 != 0
    MessageBox MB_ICONSTOP "当前用户插件注册未完成，请检查.NET Framework与目录权限后重试。" /SD IDOK
    Abort
  ${EndIf}
FunctionEnd

Function un.UnregisterOfficeAddin
  ClearErrors
  ExecWait '"$INSTDIR\app-${APPVERSION}\office-addin\Gongwen.AddinRegistrar.exe" --unregister' $0
  ${If} ${Errors}
  ${OrIf} $0 != 0
    MessageBox MB_ICONSTOP "插件注册项尚未移除，卸载已停止。请重新安装同版本后重试。" /SD IDOK
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
