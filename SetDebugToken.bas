Attribute VB_Name = "SetDebugToken"
Option Explicit

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Long, ReturnLength As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const ANYSIZE_ARRAY = 1
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
        pLuid As LARGE_INTEGER
        Attributes As Long
End Type
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Sub SetDebugToken()
    Dim hToken As Long
    Dim tkp As TOKEN_PRIVILEGES
    OpenProcessToken GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue vbNullString, SE_DEBUG_NAME, tkp.Privileges(0).pLuid
    tkp.PrivilegeCount = 1
    tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    AdjustTokenPrivileges hToken, False, tkp, 0, 0, 0
    CloseHandle hToken
End Sub

