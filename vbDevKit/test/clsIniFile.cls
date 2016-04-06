VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private Type Sect
  name As String
  Key() As String
  value() As String
End Type

Private IniObj() As Sect
Private IniFile As String

Public Sub LoadFile(INIFileLoc As String)
   
   
   IniFile = INIFileLoc
   If Dir(INIFileLoc) = "" Then Exit Sub
   inidata = readFile(INIFileLoc)
   
   tmp = Split(inidata, vbCrLf)
   For i = 0 To UBound(tmp)
      If Left(tmp(i), 1) = "[" Then tmp(i) = Replace(tmp(i), "[", Chr(5))
   Next
   
   inidata = Join(tmp, vbCrLf)
   sec = Split(inidata, Chr(5))
   ReDim IniObj(UBound(sec))
   
    For i = 1 To UBound(sec)
      IniObj(i).name = Mid(sec(i), 1, InStr(1, sec(i), "]") - 1)
      subs = Split(sec(i), vbCrLf)
      ReDim IniObj(i).Key(UBound(subs))
      ReDim IniObj(i).value(UBound(subs))
      For j = 1 To UBound(subs)
        If Trim(subs(j)) <> "" Then
          a = Split(subs(j), "=")
          IniObj(i).Key(j) = a(0)
          IniObj(i).value(j) = a(1)
        End If
      Next
    Next
End Sub

Public Function GetValue(Section, Key) As String
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    GetValue = IniObj(s).value(k)
End Function

Public Function EnumSections() As String()
    Dim r() As String
    
    If SectAryIsEmpty() Then Exit Function
    
    ReDim r(UBound(IniObj))
    For i = 0 To UBound(IniObj)
        r(i) = IniObj(i).name
    Next
    EnumSections = r()
    
End Function

Public Function EnumKeys(Section) As String()
    Dim r() As String
    If SectAryIsEmpty() Then Exit Function
    
    s = GetSectionIndex(Section)
    ReDim r(UBound(IniObj(s).Key))
    
    For j = 0 To UBound(IniObj(s).Key)
       r(j) = IniObj(s).Key(j)
    Next
    EnumKeys = r()
    
End Function


Public Function AddKey(Section, Key, value) As Boolean
    s = GetSectionIndex(Section)
    
    If s = -1 Then Exit Function
    If GetKeyIndex(s, Key) <> -1 Then Exit Function
    
    ub = UBound(IniObj(s).Key) + 1
    ReDim Preserve IniObj(s).Key(ub)
    ReDim Preserve IniObj(s).value(ub)
    IniObj(s).Key(ub) = Key
    IniObj(s).value(ub) = value
    AddKey = True
    
End Function

Function SectionExists(SectionName) As Boolean
    If GetSectionIndex(SectionName) <> -1 Then SectionExists = True
End Function

Function KeyExists(Section, Key) As Boolean
    s = GetSectionIndex(Section)
    
    If s = -1 Then Exit Function
    If GetKeyIndex(s, Key) <> -1 Then KeyExists = True
    
End Function

Public Function AddSection(SectionName) As Boolean
        If GetSectionIndex(SectionName) <> -1 Then AddSection = False: Exit Function
        ub = SectPush(IniObj)
        IniObj(ub).name = SectionName
        ReDim IniObj(ub).Key(0)
        ReDim IniObj(ub).value(0)
        AddSection = True
End Function

Public Function DeleteSection(Section) As Boolean
    s = GetSectionIndex(Section)
    If s = -1 Then DeleteSection = False: Exit Function
    ReDim IniObj(s).Key(0)
    ReDim IniObj(s).value(0)
    IniObj(s).name = ""
    DeleteSection = True
End Function

Public Function DeleteKey(Section, Key) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    If s = -1 Then DeleteKey = False: Exit Function
    If k = -1 Then DeleteKey = False: Exit Function
    IniObj(s).Key(k) = ""
    IniObj(s).value(k) = ""
    DeleteKey = True
End Function

Public Function SetValue(Section, Key, newVal) As Boolean
    s = GetSectionIndex(Section)
    k = GetKeyIndex(s, Key)
    If s = -1 Then SetValue = False: Exit Function
    If k = -1 Then SetValue = False: Exit Function
    IniObj(s).value(k) = CStr(newVal)
    SetValue = True
End Function

Public Sub Save()
    If SectAryIsEmpty Then
        MsgBox "There are no Sections in INI File, cannot save", vbInformation
        Exit Sub
    End If
    
    For i = 0 To UBound(IniObj)
      If IniObj(i).name <> "" Then
        tmp = tmp & "[" & IniObj(i).name & "]" & vbCrLf
          For j = 0 To UBound(IniObj(i).Key)
             If IniObj(i).Key(j) <> "" Then
               tmp = tmp & IniObj(i).Key(j) & "=" & IniObj(i).value(j) & vbCrLf
             End If
          Next
        tmp = tmp & vbCrLf
      End If
    Next
    
    Call WriteFile(CStr(tmp))
    ReDim IniObj(0) 'just to free up memory
    
End Sub

Public Sub Release()
    ReDim IniObj(0) 'just to free up memory
End Sub

Private Function GetSectionIndex(Section) As Integer
    On Error GoTo hell
    For i = 0 To UBound(IniObj)
       If LCase(IniObj(i).name) = LCase(Section) Then
          GetSectionIndex = CInt(i)
          Exit Function
       End If
    Next
hell:
    GetSectionIndex = -1
End Function

Private Function GetKeyIndex(SectionIndex, KeyName) As Integer
    On Error GoTo hell
    For i = 0 To UBound(IniObj(SectionIndex).Key)
        If LCase(IniObj(SectionIndex).Key(i)) = LCase(KeyName) Then
            GetKeyIndex = CInt(i)
            Exit Function
        End If
    Next
hell:
    GetKeyIndex = -1
End Function

Private Function readFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   readFile = temp
End Function

Private Sub WriteFile(it As String)
    If Len(IniFile) = 0 Then Err.Raise 1, , "No File name set, load file first"
    
    f = FreeFile
    Open IniFile For Output As #f
      Print #f, it
    Close f
End Sub


Private Function SectAryIsEmpty() As Boolean
  On Error GoTo oops
    i = UBound(IniObj)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Function SectPush(ary() As Sect) As Integer  'this modifies parent ary object
    On Error GoTo init                       'returns ubound
    x = UBound(ary)                          '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    SectPush = x
    Exit Function
init:     ReDim ary(0): SectPush = 0
End Function






Private Sub Class_Initialize()
'    If Not isRegistered And Not isInitalized Then TellThemAllAboutIt
End Sub
