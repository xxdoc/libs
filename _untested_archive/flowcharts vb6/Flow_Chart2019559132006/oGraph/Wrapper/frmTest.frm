VERSION 5.00
Object = "{24B00F00-B508-4B8E-84F8-CB55079786FD}#3.0#0"; "Graph.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H8000000A&
   Caption         =   "Flow Chart Designer : By Rajneesh"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   21690
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   26.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   180.75
   StartUpPosition =   1  'CenterOwner
   Begin oGraph.oConvas Convas1 
      Height          =   5535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9763
      ForeColor       =   0
      BackColor       =   16777215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   15000
      ScaleHeight     =   585
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmTest.frx":0000
      Top             =   240
      Width           =   6255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   750
      Max             =   40
      Min             =   12
      TabIndex        =   4
      Top             =   5850
      Value           =   20
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   3
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7290
      TabIndex        =   2
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6300
      TabIndex        =   1
      Top             =   5820
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5310
      TabIndex        =   0
      Top             =   5820
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmTest.frx":0288
      Top             =   5730
      Width           =   480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim p1 As oGraph.oPicture, p2 As oGraph.oPicture, p3 As oGraph.oPicture, p4 As oGraph.oPicture
Dim t1 As oGraph.oText, t2 As oText, t3 As oText, t4 As oText
Dim s1 As oGraph.oLine, s2 As oGraph.oLine, s3 As oGraph.oLine



Private Sub HScroll1_Change()
    Convas1.Zoom = HScroll1.Value * 5
    Convas1.Paint
End Sub


Private Sub Command1_Click(index As Integer)
    Dim Dummy As Variant, iFile As Integer, data() As Byte
    Dim UNCFile As String

    Select Case index
        
        Case 2: 'Save
            Dummy = Convas1.BinaryData
            UNCFile = App.Path & "\Convas.blg"
            iFile = FreeFile
            Open UNCFile For Binary As iFile
            Put iFile, , Dummy
            Close iFile '
        Case 3: 'Clear
            Convas1.ClearWorkSheet
        Case 4: 'Load
            iFile = FreeFile
            UNCFile = App.Path & "\Convas.blg"
            
            Open UNCFile For Binary As iFile
            Get iFile, , Dummy
            Close iFile
            'Assign the Variant to a bytearray to the bag.contents
            
            If Len(Dummy) > 0 Then
                'Convas1.BinaryData = ""
                Convas1.ClearWorkSheet
                Convas1.BinaryData = Dummy
            End If
            'Convas1.Paint
        Case 5: '100 %
            Convas1.Zoom = 100
            Convas1.Paint
        Case 6:
        Case 7:
    End Select
End Sub


    




Private Sub Form_Load()
  
    Me.Show
    
    Dim block As CBlock
    Dim cfg As New CFlowGraph
    Dim i As Long
    Dim p() As oGraph.oPicture
    
    cfg.ParseAsm Text1.Text
    
    With Convas1
        For i = 1 To cfg.blocks.Count
            Set block = cfg.blocks(i)
            cfg.GenerateNodeImage block, Picture1
            Set p1 = .AddNode("p" & i)
            Set block.node = p1
            Set block.node.Image = Picture1.Picture
            
            Set p1 = block.node
            p1.Visible = True
            p1.Activate
        Next
    End With
    
    For i = 1 To cfg.blocks.Count
        
    Next
    
'    With Convas1
'        Set p1 = .AddNode("p1")
'        Set p2 = .AddNode("p2")
'        Set p3 = .AddNode("p3")
'        Set p4 = .AddNode("p4")
'        Set t1 = .AddText("t1")
'        Set t2 = .AddText("t2")
'        Set t3 = .AddText("t3")
'        'Set t4 = .AddText("t4")
'
'        Set s1 = .AddStep(OnTCompletion, "p1", "p2", "s1")
'        Set s2 = .AddStep(OnTFail, "p2", "p3", "s2")
'        Set s3 = .AddStep(OnTSuccess, "p2", "p4", "s3")
'    End With
'
'
'
'    With p1
'
'        .Caption = asm1
'        .CentreX = 1200
'        .CentreY = 1200
'        .ToolTipText = "I am " & .Caption
'        Set .Image = Image1.Picture
'    End With
'
'    With p2
'        .Caption = asm2
'
'        .CentreX = 4200
'        .CentreY = 1200
'        .ToolTipText = "I am " & .Caption
'        Set .Image = Image1.Picture
'    End With
'
'    With p3
'        .Caption = asm3
'        .CentreX = 1200
'        .CentreY = 4200
'        .ToolTipText = "I am " & .Caption
'        Set .Image = Image1.Picture
'    End With
'
'    With p4
'        .Caption = "All Gone Well - Winner"
'        .CentreX = 6200
'        .CentreY = 4200
'        .ToolTipText = "I am " & .Caption
'        Set .Image = Image1.Picture
'    End With
'
'    With t1
'        .Caption = asm1
'        .CentreX = 4500
'        .CentreY = 300
'        MsgBox .tHeight
'        .Font.Size = 20
'        .ToolTipText = "I am " & .Caption
'        .Font.Bold = True
'    End With
'
'    With t2
'        .Caption = asm2
'        .CentreX = 4500
'        .CentreY = 300
'        .Font.Size = 20
'        .ToolTipText = "I am " & .Caption
'        .Font.Bold = True
'    End With
'
'
'    With t3
'        .Caption = asm3
'        .CentreX = 4500
'        .CentreY = 300
'        .Font.Size = 20
'        .ToolTipText = "I am " & .Caption
'        .Font.Bold = True
'    End With
'
'
'     s1.LayereLineType = OnTCompletion
'     s2.LayereLineType = OnTFail
'     s3.LayereLineType = OnTSuccess
'
'     s1.ToolTipText = "I am line s1"
'     s2.ToolTipText = "I am line s2"
'     s3.ToolTipText = "I am line s3"
'
'
'
''     p1.Visible = True
''     p2.Visible = True
''     p3.Visible = True
''     p4.Visible = True
'
'     t1.Visible = True
'     t2.Visible = True
'     t3.Visible = True
'
'     s1.Visible = True
'     s2.Visible = True
'     s3.Visible = True
'     p1.Activate
'     p2.Activate
'     p3.Activate
'     p4.Activate
'     t1.Activate
     
     
     
End Sub



'Private Sub Form_Resize()
'    If Me.WindowState <> vbMinimized Then
'        Convas1.Width = Me.Width - 220
'        Convas1.Height = Me.Height - 1000
'        HScroll1.Top = Convas1.Height + Convas1.Top + 50
'        Command1(2).Top = HScroll1.Top
'        Command1(3).Top = HScroll1.Top
'        Command1(4).Top = HScroll1.Top
'        Command1(5).Top = HScroll1.Top
'    End If
'End Sub



