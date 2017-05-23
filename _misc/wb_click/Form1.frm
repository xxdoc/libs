VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      ExtentX         =   10186
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    URL = "file://" & App.Path & "\page.html"
    Debug.Print "Html Page: " & URL
    wb.Navigate2 URL
    
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    On Error Resume Next
    
    Debug.Print "Url: " & URL
    If Len(PostData) > 0 Then Debug.Print "PostData: " & PostData
    
    If InStr(1, URL, "btn1.click", vbTextCompare) > 0 Then
        Cancel = True
        Command1_Click
    End If
    
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    txt1 = wb.Document.getElementById("txt1").Value
    MsgBox "VB6 Code: User clicked the web browser button txt1 = " & txt1
End Sub

