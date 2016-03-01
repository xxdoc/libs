VERSION 5.00
Object = "{B3802EF7-CC1A-4294-B64B-33354DC196B7}#1.1#0"; "kroolCtls.ocx"
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
   Begin kroolCtls.TreeView TreeView1 
      Height          =   3030
      Left            =   8010
      TabIndex        =   4
      Top             =   2385
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5345
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin kroolCtls.RichTextBox RichTextBox1 
      Height          =   3075
      Left            =   1440
      TabIndex        =   3
      Top             =   4365
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   5424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      WantReturn      =   -1  'True
      Text            =   "Form1.frx":0000
      TextRTF         =   "Form1.frx":0039
   End
   Begin kroolCtls.IPAddress IPAddress1 
      Height          =   555
      Left            =   7470
      TabIndex        =   2
      Top             =   1260
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin kroolCtls.ProgressBar ProgressBar1 
      Height          =   555
      Left            =   7425
      Top             =   270
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   979
      Step            =   10
   End
   Begin kroolCtls.ListView ListView1 
      Height          =   3480
      Left            =   4725
      TabIndex        =   1
      Top             =   540
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   6138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin kroolCtls.TabStrip TabStrip1 
      Height          =   3390
      Left            =   315
      TabIndex        =   0
      Top             =   405
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   5980
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InitTabs        =   "Form1.frx":01B9
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
