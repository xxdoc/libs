Attribute VB_Name = "modVista"
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'vista+ only
Private Type TOKEN_ELEVATION
    TokenIsElevated As Long
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const TOKEN_READ As Long = &H20008
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ELEVATION_TYPE As Long = 18
Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Enum TOKEN_INFORMATION_CLASS
    TokenUser = 1
    TokenGroups
    TokenPrivileges
    TokenOwner
    TokenPrimaryGroup
    TokenDefaultDacl
    TokenSource
    TokenType
    TokenImpersonationLevel
    TokenStatistics
    TokenRestrictedSids
    TokenSessionId
    TokenGroupsAndPrivileges
    TokenSessionReference
    TokenSandBoxInert
    TokenAuditPolicy
    TokenOrigin
    tokenElevationType
    TokenLinkedToken
    TokenElevation
    TokenHasRestrictions
    TokenAccessInformation
    TokenVirtualizationAllowed
    TokenVirtualizationEnabled
    TokenIntegrityLevel
    TokenUIAccess
    TokenMandatoryPolicy
    TokenLogonSid
    MaxTokenInfoClass  'MaxTokenInfoClass should always be the last enum
End Enum

Private Declare Function ChangeWindowMessageFilter Lib "user32" (ByVal msg As Long, ByVal flag As Long) As Long 'Vista+
Const WM_COPYDATA = &H4A
Const WM_COPYGLOBALDATA = &H49
Const MSGFLT_ADD = 1
Const MSGFLT_REMOVE = 2

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Enum dataType
    REG_BINARY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
    'REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
    'REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    'REG_MULTI_SZ = 7                   ' Multiple Unicode strings
    REG_SZ = 1                         ' Unicode nul terminated string
End Enum

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Function AllowCopyDataAcrossUIPI(Optional allow As Boolean)
    Dim a, b, c
    Dim action As Long
    action = IIf(allow, MSGFLT_ADD, MSGFLT_REMOVE)
    b = ChangeWindowMessageFilter(WM_COPYDATA, action) 'we still need this for IPC to get hook data...
    c = ChangeWindowMessageFilter(WM_COPYGLOBALDATA, action)
    'MsgBox a & " " & b & " " & c
End Function


Public Function IsVistaPlus() As Boolean
    Dim osVersion As OSVERSIONINFO
    osVersion.dwOSVersionInfoSize = Len(osVersion)
    If GetVersionEx(osVersion) = 0 Then Exit Function
    If osVersion.dwPlatformId <> VER_PLATFORM_WIN32_NT Or osVersion.dwMajorVersion < 6 Then Exit Function
    IsVistaPlus = True
End Function

Function IsProcessElevated() As Boolean

    Dim fIsElevated As Boolean
    Dim dwError As Long
    Dim hToken As Long

    'Open the primary access token of the process with TOKEN_QUERY.
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hToken) = 0 Then GoTo cleanup
     
    Dim elevation As TOKEN_ELEVATION
    Dim dwSize As Long
    If GetTokenInformation(hToken, TOKEN_INFORMATION_CLASS.TokenElevation, elevation, Len(elevation), dwSize) = 0 Then
        'When the process is run on operating systems prior to Windows Vista, GetTokenInformation returns FALSE with the
        'ERROR_INVALID_PARAMETER error code because TokenElevation is not supported on those operating systems.
         dwError = Err.LastDllError
         GoTo cleanup
    End If

    fIsElevated = IIf(elevation.TokenIsElevated = 0, False, True)

cleanup:
    If hToken Then CloseHandle (hToken)
    'if ERROR_SUCCESS <> dwError then err.Raise
    IsProcessElevated = fIsElevated
End Function

Function isUACEnabled() As Boolean
    Dim v
    v = ReadRegValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA")
    If v = -1 Then Form1.List1.AddItem "Failed to read UAC setting from registry"
    isUACEnabled = (v = 1)
End Function

Function ReadRegValue(hive As hKey, path, ByVal KeyName)
     
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    Dim ret As Long
    'retrieve nformation about the key
    Dim p As String
    Dim handle As Long
    
    ReadRegValue = -1
    p = path
    RegOpenKeyEx hive, p, 0, KEY_READ, handle
    lResult = RegQueryValueEx(handle, CStr(KeyName), 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then ReadRegValue = Replace(strBuf, Chr$(0), "")
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, strData, lDataBufSize)
            If lResult = 0 Then ReadRegValue = strData
        ElseIf lValueType = REG_DWORD Then
            Dim x As Long
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, x, lDataBufSize)
            ReadRegValue = x
        ElseIf lValueType = REG_EXPAND_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then ReadRegValue = Replace(strBuf, Chr$(0), "")

        'Else
        '    MsgBox "UnSupported Type " & lValueType
        End If
    End If
    RegCloseKey handle
    
End Function
