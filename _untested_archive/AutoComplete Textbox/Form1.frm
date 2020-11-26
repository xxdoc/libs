VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RTB 
      Height          =   360
      Left            =   420
      TabIndex        =   1
      Top             =   1860
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   420
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   420
      TabIndex        =   4
      Top             =   2400
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Dummy TextBox to test tabbing, focus gain/loss, etc.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   270
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Auto-Complete RichTextBox:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   1590
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRS As ADODB.Recordset
Private WithEvents mAutoCompleteTextBox As cAutoCompleteTextBox
Attribute mAutoCompleteTextBox.VB_VarHelpID = -1
Private Sub Form_Load()
   Set mRS = New ADODB.Recordset
   With mRS
      .Fields.Append "ID", adBigInt
      .Fields.Append "Name", adVarChar, 255
      
      .Open
      .AddNew Array("ID", "Name"), Array(1, "Adams, Bob")
      .AddNew Array("ID", "Name"), Array(2, "Fisher, John")
      .AddNew Array("ID", "Name"), Array(3, "Williams, Ted")
      .AddNew Array("ID", "Name"), Array(4, "Williams, Fred")
      .AddNew Array("ID", "Name"), Array(5, "Jones, William")
      .AddNew Array("ID", "Name"), Array(6, "Thomson, Adam")
      .Update
      .Sort = "Name ASC"
  End With
  
  Label2.Caption = "You can use UP/DOWN arrows to move through the items that match the typed characters (as well as END. HOME). " & vbCrLf & vbCrLf & _
                  "DELETE or ESC will both remove all typed chars (allowing you to do a new search)." & vbCrLf & vbCrLf & _
                  "BACKSPACE removes the last-typed character."
  
  Set mAutoCompleteTextBox = New cAutoCompleteTextBox
  mAutoCompleteTextBox.Init RTB, mRS, "ID", "Name", "Enter a Staff member name", "Start Typing..."
End Sub

Private Sub mAutoCompleteTextBox_MatchSelected(MatchedID As String, MatchedText As String)
   If MatchedID = vbNullString Then
      'MsgBox "This demo requires that you make a selection!"
      'RTB.SetFocus
   Else
      Debug.Print "User chose ", MatchedID, MatchedText
   End If
End Sub
