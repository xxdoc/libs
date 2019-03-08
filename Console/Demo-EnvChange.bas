Attribute VB_Name = "MDemoEnvChange"
' *************************************************************************
'  Copyright ©2007 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Function apiRegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function apiRegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function apiRegFlushKey Lib "advapi32.dll" Alias "RegFlushKey" (ByVal hKey As Long) As Long
Private Declare Function apiRegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SHDeleteValue Lib "shlwapi" Alias "SHDeleteValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String) As Long

' Constants used with registry calls
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ As Long = 1&
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

' Constants used to broadcast messages
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const WM_SETTINGCHANGE As Long = &H1A

' Reg Key Security Options
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))


Public Sub Main()
   Dim nResult As Long
   Const lmKey = "System\CurrentControlSet\Control\Session Manager\Environment"
   Const cuKey = "Environment"
   Const lmEvar = "ClassicVB-lm"
   Const cuEvar = "ClassicVB-cu"
   
   ' Required in all MConsole.bas supported apps!
   Con.Initialize

   ' Check whether to clear or set variables.
   Con.WriteLine "Writing to registry...  ", False
   If InStr(1, Command$, "/clear", vbTextCompare) Then
      ' Clear e-vars at the User and Machine levels.
      ' HKLM calls may fail if not running as admin.
      Call RegDeleteValue(HKEY_CURRENT_USER, cuKey, cuEvar)
      Call RegDeleteValue(HKEY_LOCAL_MACHINE, lmKey, lmEvar)
   Else
      ' Set e-vars at the User and Machine levels.
      ' HKLM calls may fail if not running as admin.
      Call RegSetStringValue(HKEY_CURRENT_USER, cuKey, cuEvar, "Rocks!")
      Call RegSetStringValue(HKEY_LOCAL_MACHINE, lmKey, lmEvar, "Rocks!")
   End If
   
   ' Tell the world what we've done.
   Con.WriteLine "Broadcasting change..."
   Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, "Environment", SMTO_ABORTIFHUNG, 5000, nResult)
   
   ' Demonstrate success, or lack thereof
   Con.WriteLine cuEvar & "=" & Environ$(cuEvar)
   Con.WriteLine lmEvar & "=" & Environ$(lmEvar)
   
   ' Allow user to see output if launched from Explorer.
   If Con.LaunchMode = conLaunchExplorer Then
      Con.PressAnyKey
   End If
End Sub

Private Function RegDeleteValue(ByVal RootKey As Long, ByVal Key As String, ByVal Value As String) As Boolean
   Dim nRet As Long
   ' Just delete this single value.
   nRet = SHDeleteValue(RootKey, Key, Value)
   ' Return result of SHDeleteValue call.
   RegDeleteValue = (nRet = ERROR_SUCCESS)
End Function

Private Function RegSetStringValue(ByVal RootKey As Long, ByVal Key As String, ByVal ValueName As String, ByVal Value As String) As Boolean
   Dim nRet As Long
   Dim hKey As Long
   ' Open a key and set a value within it.
   If apiRegOpenKeyEx(RootKey, Key, 0&, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
      ' If NULL, the default value will be written.
      If ValueName = "*" Then ValueName = vbNullString
      ' Attempt to write data - Always a string.
      nRet = apiRegSetValueEx(hKey, ValueName, 0&, REG_SZ, ByVal Value, Len(Value))
      Call apiRegFlushKey(hKey)
      Call apiRegCloseKey(hKey)
      ' Return result of RegSetValueEx call.
      RegSetStringValue = (nRet = ERROR_SUCCESS)
   End If
End Function
