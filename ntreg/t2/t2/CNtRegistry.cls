VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNtRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hive As hKey
Private USER_SID As String

Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    'HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    'HKEY_DYN_DATA = &H80000006
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

Private Declare Function keyExists_ Lib "ntreg.dll" (ByVal path As String) As Boolean
Private Declare Function deleteValue_ Lib "ntreg.dll" (ByVal path As String, ByVal valueName As String) As Boolean
Private Declare Function createKey_ Lib "ntreg.dll" (ByVal path As String) As Boolean


'int __stdcall userSID(char* sid, int *size){
Private Declare Function userSID Lib "ntreg.dll" (ByVal sid As String, ByRef size As Long) As Boolean

Private Sub Class_Initialize()
    Dim buf As String, size As Long
    size = 500
    buf = Space(size)
    If userSID(buf, size) Then
        USER_SID = Left(buf, size)
    Else
        If size > 0 Then
            buf = Space(size)
            If userSID(buf, size) Then USER_SID = Left(buf, size)
        End If
    End If
    'MsgBox USER_SID
End Sub

Private Function GetRootPathFor(hRoot As hKey)
 
    If hRoot = HKEY_LOCAL_MACHINE Then
        GetRootPathFor = "\Registry\Machine"
    ElseIf hRoot = HKEY_CLASSES_ROOT Then
        GetRootPathFor "\Registry\Machine\SOFTWARE\Classes"
    ElseIf (hRoot = HKEY_CURRENT_CONFIG) Then
        GetRootPathFor = "\Registry\Machine\System\CurrentControlSet\Hardware Profiles\Current"
    ElseIf hRoot = HKEY_USERS Then
        GetRootPathFor = "\Registry\User"
    Else
        GetRootPathFor = "\Registry\User\" & USER_SID
    End If

End Function

Private Sub SetPath(ByRef path As String)
    
    If InStr(path, "\") < 1 Then Exit Sub
    
    Dim p() As String

    p = Split(path, "\")
    p(0) = UCase(p(0))
    
    If p(0) = "HKLM" Or p(0) = "HKEY_LOCAL_MACHINE" Then
        Me.hive = HKEY_LOCAL_MACHINE
        p(0) = GetRootPathFor(Me.hive)
    ElseIf p(0) = "HKCR" Or p(0) = "HKEY_CLASSES_ROOT" Then
        Me.hive = HKEY_CLASSES_ROOT
        p(0) = GetRootPathFor(Me.hive)
    ElseIf p(0) = "HKCU" Or p(0) = "HKEY_CURRENT_USER" Then
         Me.hive = HKEY_CURRENT_USER
        p(0) = GetRootPathFor(Me.hive)
    ElseIf p(0) = "HKU" Or p(0) = "HKEY_USERS" Then
        Me.hive = HKEY_USERS
        p(0) = GetRootPathFor(Me.hive)
    ElseIf p(0) = "HKCC" Or p(0) = "HKEY_CURRENT_CONFIG" Then
        Me.hive = HKEY_CURRENT_CONFIG
        p(0) = GetRootPathFor(Me.hive)
    Else
        Exit Sub
    End If
    
    path = Join(p, "\")
    
End Sub

Function keyExists(ByVal path As String) As Boolean
    SetPath path
    keyExists = keyExists_(path)
End Function

Function DeleteValue(ByVal path As String, valueName) As Boolean
    SetPath path
    DeleteValue = deleteValue_(path, valueName)
End Function

'Function DeleteKey(path) As Boolean
'
'End Function

'Function CreateKey(ByVal path As String) As Boolean
'    SetPath path
'    CreateKey = createKey_(path)
'End Function

Function SetValue(ByVal path As String, ByVal KeyName As String, Data, dType As dataType) As Boolean


End Function

Function ReadValue(path, ByVal KeyName)
'
'    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
'    Dim ret As Long
'    'retrieve nformation about the key
'    Dim p As String
'
'    p = stdPath(path)
'    RegOpenKeyEx hive, p, 0, KEY_READ, Handle
'    lResult = RegQueryValueEx(Handle, CStr(KeyName), 0, lValueType, ByVal 0, lDataBufSize)
'    If lResult = 0 Then
'        If lValueType = REG_SZ Then
'            strBuf = String(lDataBufSize, Chr$(0))
'            lResult = RegQueryValueEx(Handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
'            If lResult = 0 Then ReadValue = Replace(strBuf, Chr$(0), "")
'        ElseIf lValueType = REG_BINARY Then
'            Dim strData As Integer
'            lResult = RegQueryValueEx(Handle, CStr(KeyName), 0, 0, strData, lDataBufSize)
'            If lResult = 0 Then ReadValue = strData
'        ElseIf lValueType = REG_DWORD Then
'            Dim x As Long
'            lResult = RegQueryValueEx(Handle, CStr(KeyName), 0, 0, x, lDataBufSize)
'            ReadValue = x
'        ElseIf lValueType = REG_EXPAND_SZ Then
'            strBuf = String(lDataBufSize, Chr$(0))
'            lResult = RegQueryValueEx(Handle, CStr(KeyName), 0, 0, ByVal strBuf, lDataBufSize)
'            If lResult = 0 Then ReadValue = Replace(strBuf, Chr$(0), "")
'
'        'Else
'        '    MsgBox "UnSupported Type " & lValueType
'        End If
'    End If
'    RegCloseKey Handle
    
End Function

Function EnumKeys(path) As String()
    
    
End Function

Function EnumValues(path) As String()
   
   
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    'Search the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    StripTerminator = sInput
    If ZeroPos > 0 Then StripTerminator = Left$(sInput, ZeroPos - 1)
End Function

Private Function stdPath(sIn) As String
    stdPath = Replace(sIn, "/", "\")
    If Left(stdPath, 1) = "\" Then stdPath = Mid(stdPath, 2, Len(stdPath))
    If Right(stdPath, 1) <> "\" Then stdPath = stdPath & "\"
End Function



