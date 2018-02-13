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
Attribute VB_Exposed = False
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

    
    asStr = "CompanyName      " & CompanyName & vbCrLf & _
                "FileDescription  " & FileDescription & vbCrLf & _
                "FileVersion      " & FileVersion & vbCrLf & _
                "InternalName     " & InternalName & vbCrLf & _
                "LegalCopyright   " & LegalCopyright & vbCrLf & _
                "OriginalFilename " & OrigionalFileName & vbCrLf & _
                "ProductName      " & ProductName & vbCrLf & _
                "ProductVersion   " & ProductVersion
                

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
