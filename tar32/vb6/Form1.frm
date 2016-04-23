VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tar32.dll VB6 demo"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2265
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3870
      Width           =   10140
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tar As New CTarFile

Private Sub Form_Load()
    
    Dim c As New Collection, cc, src, tst, tst2, lst() As String, f As String
    
    f = GetParentFolder(App.path) & "\test.tar.gz"
    
    If Not tar.isInitilized Then
        List1.AddItem "Could not load tar32.dll?"
        Exit Sub
    End If
    
    If Not tar.FileExists(f) Then
        List1.AddItem "Could not find test tar.gz file?"
        Exit Sub
    End If
    
    List1.AddItem "Tar32.dll Version: " & tar.Version()
    List1.AddItem "Check Archive: " & tar.isValidArchive(f)
    List1.AddItem "File count: " & tar.FileCount(f)
    
    Set c = tar.EnumFiles(f, "*.php")
    List1.AddItem "Enumerating php files in archive...Found " & c.Count
    
    'For Each cc In c
    '    List1.AddItem cc
    'Next
    
    List1.AddItem "Creating archive of tar cpp files"
    src = GetParentFolder(App.path) & "\src"
    tst = App.path & "\test.tgz"
    tst2 = App.path & "\test.tar"
    
    tar.KeepDirectoryStructure = False
    
    If Not tar.ArchiveFromFolder(src, tst, "*.cpp", atTgz) Then
        List1.AddItem "Failed! " & tar.LastError
    Else
        List1.AddItem "Success! File Size: 0x" & Hex(FileLen(tst))
    End If
    
    List1.AddItem "Creating archive from file list.."
    
    push lst, App.path & "\CTarFile.cls"
    push lst, App.path & "\Form1.frm"
    
    If Not tar.ArchiveFromFiles(lst, tst2, atTar) Then
        List1.AddItem "Failed! " & tar.LastError
    Else
        List1.AddItem "Success! File Size: 0x" & Hex(FileLen(tst2))
    End If
    
    List1.AddItem "Extracting first php file to c:\"
    If Not tar.ExtractFile(f, c(1), "C:\") Then
        List1.AddItem "Failed! " & tar.LastError
    Else
        List1.AddItem "Success!"
    End If
    
    List1.AddItem "Test complete!"
    
    
End Sub




Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function GetParentFolder(path) As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function
