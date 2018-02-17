VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public CompanyName As String
Public FileDescription As String
Public FileVersion As String
Public InternalName As String
Public LegalCopyright As String
Public OrigionalFileName As String
Public ProductName As String
Public ProductVersion As String
Public LanguageID As String


Property Get asStr()

    Dim x() As String
    
    If Len(CompanyName) > 0 Then push x, "CompanyName      " & CompanyName
    If Len(FileDescription) > 0 Then push x, "FileDescription  " & FileDescription
    If Len(FileVersion) > 0 Then push x, "FileVersion      " & FileVersion
    If Len(InternalName) > 0 Then push x, "InternalName     " & InternalName
    If Len(LegalCopyright) > 0 Then push x, "LegalCopyright   " & LegalCopyright
    If Len(OrigionalFileName) > 0 Then push x, "OriginalFilename " & OrigionalFileName
    If Len(ProductName) > 0 Then push x, "ProductName      " & ProductName
    If Len(ProductVersion) > 0 Then push x, "ProductVersion   " & ProductVersion
               
    asStr = Join(x, vbCrLf)

End Property

Friend Sub Load(strVersionInfo() As String, strLangCharset2 As String)

    CompanyName = strVersionInfo(0)
    FileDescription = strVersionInfo(1)
    FileVersion = strVersionInfo(2)
    InternalName = strVersionInfo(3)
    LegalCopyright = strVersionInfo(4)
    OrigionalFileName = strVersionInfo(5)
    ProductName = strVersionInfo(6)
    ProductVersion = strVersionInfo(7)
    LanguageID = strLangCharset2
    
End Sub
