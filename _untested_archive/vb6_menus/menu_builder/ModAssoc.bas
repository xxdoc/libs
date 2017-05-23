Attribute VB_Name = "ModAssoc"
Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" _
(ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Function Associate(ByVal apPath As String, ByVal Ext As String) As Boolean
'Borrowed this association function from a submission by
' Insomniaque modified by Dj's Computer Labs
'Da rest is all Bobo Enterprises copyright
  Dim sKeyName As String
  Dim sKeyValue As String
  Dim ret&
  Dim lphKey&
  Dim apTitle As String
  apTitle = ParseName(apPath)
  If InStr(Ext, ".") = 0 Then Ext = "." & Ext
   sKeyName = Ext
  sKeyValue = apTitle
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
  If ret& <> 0 Then GoTo AssocFailed
   sKeyName = apTitle
  sKeyValue = apPath & " %1"
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
  If ret& <> 0 Then GoTo AssocFailed
    sKeyValue = apPath
  ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
  If ret& <> 0 Then GoTo AssocFailed
  ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
  If ret& <> 0 Then GoTo AssocFailed
   Associate = True
  Exit Function
AssocFailed:
  Associate = False
End Function
Public Function ParseName(ByVal sPath As String) As String
  Dim strX As String
  Dim intX As Integer
  intX = InStrRev(sPath, "\")
  strX = Trim(Right(sPath, Len(sPath) - intX))
  If Right(strX, 1) = Chr(0) Then
    ParseName = Left(strX, Len(strX) - 1)
  Else
    ParseName = strX
  End If
End Function
Public Sub FileSave(Text As String, FilePath As String)
On Error Resume Next
Dim Directory As String
              Directory$ = FilePath
              Open Directory$ For Output As #1
           Print #1, Text
       Close #1
Exit Sub
End Sub
Function TrimVoid(Expre)
  On Error Resume Next
  Dim i As Integer
  Dim beg As String
  Dim expr As String
  For i = 1 To Len(Expre)
        beg = Mid(Expre, i, 1)
        If beg Like "[a-zA-Z0-9]" Then expr = expr & beg
    Next
    TrimVoid = expr
End Function

Public Function GetShortCut(cboindex As Integer) As String
'Shortened version thanks to a friend from Brazil
   Select Case cboindex
      Case 1 To 26
         GetShortCut = "^" + Chr$(64 + cboindex)
      Case 27 To 38
         GetShortCut = "{F" + Format(cboindex - 26) + "}"
      Case 39 To 50
         GetShortCut = "^{F" + Format(cboindex - 38) + "}"
      Case 51 To 62
         GetShortCut = "+{F" + Format(cboindex - 50) + "}"
      Case 63 To 74
         GetShortCut = "+^{F" + Format(cboindex - 62) + "}"
      Case 75
         GetShortCut = "^{INSERT}"
      Case 76
         GetShortCut = "+{INSERT}"
      Case 77
         GetShortCut = "{DEL}"
      Case 78
         GetShortCut = "+{DEL}"
      Case 79
         GetShortCut = "%{BKSP}"
   End Select
End Function

