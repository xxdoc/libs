'==========mNTPrivileges.bas==========
'Module to obtain NT privilegest to access remote process
Option Explicit

Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8

Private Type LARGE_INTEGER
  LowPart As Long
  HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  pLuid As LARGE_INTEGER
  Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
  ByVal ProcessHandle As Long, _
  ByVal DesiredAccess As Long, _
  TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
  Alias "LookupPrivilegeValueA" ( _
  ByVal lpSystemName As String, _
  ByVal lpName As String, _
  lpLuid As LARGE_INTEGER) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" ( _
  ByVal TokenHandle As Long, _
  ByVal DisableAllPrivileges As Long, _
  ByRef NewState As TOKEN_PRIVILEGES, _
  ByVal BufferLength As Long, _
  ByRef PreviousState As Any, _
  ByRef ReturnLength As Any) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private bDone As Boolean

Public Function EnableDebugPrivNT() As Boolean
  If bDone Then
     EnableDebugPrivNT = True
     Exit Function
  End If
 
  Dim hToken As Long
  Dim li As LARGE_INTEGER
  Dim tkp As TOKEN_PRIVILEGES
 
  If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES _
                      Or TOKEN_QUERY, hToken) = 0 Then Exit Function
 
  If LookupPrivilegeValue("", SE_DEBUG_NAME, li) = 0 Then Exit Function
 
  tkp.PrivilegeCount = 1
  tkp.Privileges(0).pLuid = li
  tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
 
  bDone = AdjustTokenPrivileges(hToken, False, tkp, 0, ByVal 0&, 0)
  EnableDebugPrivNT = bDone
End Function
