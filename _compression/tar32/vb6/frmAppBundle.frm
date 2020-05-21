VERSION 5.00
Begin VB.Form frmBatchUntar 
   Caption         =   "Batch UnTar"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDir 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   180
      Width           =   375
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   780
      TabIndex        =   3
      Top             =   720
      Width           =   7755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "Dir"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   495
   End
End
Attribute VB_Name = "frmBatchUntar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tar As New CTarFile

Private Sub Form_Load()

    If Not tar.isInitilized Then
        List1.AddItem "Could not load tar32.dll?"
        Command1.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub cmdDir_Click()
    Text1 = fso.dlg.FolderDialog2()
End Sub

Private Sub Command1_Click()
    
    Dim c As Collection
    Dim toDir As String, f, ok As Boolean
    
    Set c = fso.GetFolderFiles(Text1)
    List1.Clear
    
    If c.Count = 0 Then
        List1.AddItem "no files"
        Exit Sub
    End If
    
    For Each f In c
        toDir = fso.GetParentFolder(f) & "\_" & fso.GetBaseName(f)
        If Not fso.FolderExists(toDir) Then MkDir toDir
        List1.AddItem "Processing " & Replace(f, Text1, "./")
        List1.AddItem "   Files: " & tar.FileCount(f)
        ok = tar.ExtractTo(f, toDir)
        List1.AddItem "   Extracted: " & ok & " " & tar.LastError
    Next
    
    List1.AddItem "done"
        
End Sub


