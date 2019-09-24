Attribute VB_Name = "mdlLogon"
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Type LUID
LowPart As Long
HighPart As Long
End Type
Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type
Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

'===============================================
'set the shut down privilege for the current application
Public Sub EnableShutDown()
Dim hProc As Long
Dim hToken As Long
Dim mLUID As LUID
Dim mPriv As TOKEN_PRIVILEGES
Dim mNewPriv As TOKEN_PRIVILEGES
hProc = GetCurrentProcess()
OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
mPriv.PrivilegeCount = 1
mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
mPriv.Privileges(0).pLuid = mLUID
' enable shutdown privilege for the current application
AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
End Sub
' Shut Down NT
Public Sub ShutDownNT(Force As Boolean)
Dim ret As Long
Dim Flags As Long
Flags = EWX_SHUTDOWN
If Force Then Flags = Flags + EWX_FORCE
If IsWinNT Then EnableShutDown
ExitWindowsEx Flags, 0
End Sub
'Restart NT
Public Sub RebootNT(Force As Boolean)
Dim ret As Long
Dim Flags As Long
Flags = EWX_REBOOT
If Force Then Flags = Flags + EWX_FORCE
If IsWinNT Then EnableShutDown
ExitWindowsEx Flags, 0
End Sub

'Log off the current user
Public Sub LogOffNT(Force As Boolean)
Dim ret As Long
Dim Flags As Long
Flags = EWX_LOGOFF
If Force Then Flags = Flags + EWX_FORCE
ExitWindowsEx Flags, 0
End Sub

'crie um formulário com 3 botões

'public Sub Command1_Click()
'    LogOffNT True
'End Sub
'public Sub Command2_Click()
'RebootNT True
'End Sub
'public Sub Command3_Click()
'ShutDownNT True
'End Sub

'public Sub Form_Load()
'KPD-Team 2000
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net
'Command1.Caption = "Log Off NT"
'Command2.Caption = "Reboot NT"
'Command3.Caption = "Shutdown NT"
'End Sub

'======================================
'Detect if the program is running under NT platform
Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

' Log off the current user
Public Sub LogOff()
    Dim ret As Long
    Dim Flags As Long
    Flags = EWX_LOGOFF
    If IsWinNT Then
    Flags = Flags + EWX_FORCE
    ExitWindowsEx Flags, 0
    End If
End Sub

'LogOff Windows
'======================================================
