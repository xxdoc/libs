VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   15855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   6735
      Left            =   11400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   180
      Width           =   4395
   End
   Begin VB.TextBox Text3 
      Height          =   1875
      Left            =   11400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":3067
      Top             =   7200
      Width           =   4395
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   180
      Top             =   9900
      _ExtentX        =   1005
      _ExtentY        =   1005
      Language        =   "jscript"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   10020
      TabIndex        =   4
      Top             =   9960
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "Form1.frx":306D
      Top             =   7200
      Width           =   11175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10020
      TabIndex        =   2
      Top             =   6660
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6660
      Width           =   9615
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11175
      ExtentX         =   19711
      ExtentY         =   11456
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Command1_Click()
    wb.Navigate2 Text1.Text
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Text3 = Empty
    sc.Reset
    sc.AddObject "p", Me, True
    sc.AddObject "wb", wb, True
    sc.AddCode Text2.Text
End Sub

Private Sub Form_Load()
    Dim y()
    x = Split(Text4, vbCrLf)
    For Each xx In x
        a = InStrRev(xx, "#")
        If a > 0 Then
           b = Mid(xx, a + 1)
           push y, Replace(b, "post", "")
        End If
    Next
    Me.Caption = UBound(y)
    Text4 = Join(y, vbCrLf)
End Sub

Sub delay(x)
    Dim dt As Date
    dt = Now
    Do While DateDiff("s", dt, Now) < x
        DoEvents
        Sleep 50   ' put your app to sleep in small increments
                   ' to avoid having CPU go to 100%
    Loop
End Sub

Private Sub sc_Error()
    Text3 = sc.Error.Description
End Sub

Sub alert(x)
    MsgBox x
End Sub
