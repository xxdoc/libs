VERSION 5.00
Object = "{13721F52-9B62-4CFD-B602-B9C73642064A}#1.0#0"; "kroolCtls.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin kroolCtls.RichTextBox rtf 
      Height          =   2985
      Left            =   5265
      TabIndex        =   3
      Top             =   4815
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   5265
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   3
      WantReturn      =   -1  'True
      Text            =   "Form1.frx":0000
      TextRTF         =   "Form1.frx":0039
   End
   Begin kroolCtls.ProgressBar pb 
      Height          =   420
      Left            =   5220
      Top             =   3690
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   741
      Step            =   10
   End
   Begin kroolCtls.TabStrip TabStrip1 
      Height          =   2805
      Left            =   360
      TabIndex        =   2
      Top             =   3195
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4948
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitTabs        =   "Form1.frx":01AD
   End
   Begin kroolCtls.TreeView tv 
      Height          =   2940
      Left            =   4950
      TabIndex        =   1
      Top             =   180
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5186
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   6
      LineStyle       =   1
      LabelEdit       =   1
      Indentation     =   1
   End
   Begin kroolCtls.ListView lv 
      Height          =   2715
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   4789
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      LabelEdit       =   1
      HideSelection   =   0   'False
      GroupView       =   -1  'True
      GroupSubsetCount=   2
      UseColumnChevron=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Form_Load()
    
    
    Dim li As LvwListItem
    lv.ColumnHeaders.Add , , "Col1"
    lv.ColumnHeaders.Add , , "Col2"
    Set li = lv.ListItems.Add(, , "test")
    li.SubItems(1) = "fart"
    
    Dim n As kroolCtls.TvwNode
    Set n = tv.Nodes.Add(, , "top-key", "top")
    tv.Nodes.Add n.Key, TvwNodeRelationshipChild, , "child"
    n.Expanded = True
    
    Me.Visible = True
    
    pb.Max = 100
    For i = 0 To 101
        Me.Refresh
        Sleep 10
        DoEvents
        pb.Value = pb.Value + 1
    Next
    
     
    
End Sub

Private Sub lv_ItemClick(ByVal Item As kroolCtls.LvwListItem, ByVal Button As Integer)
    MsgBox Item.Text
End Sub

Private Sub TabStrip1_TabClick(ByVal TabItem As kroolCtls.TbsTab)
    MsgBox TabStrip1.SelectedItem.Caption
End Sub
