VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin oGraph.oConvas oConvas1 
      Height          =   6015
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p1 As oGraph.oPicture, p2 As oGraph.oPicture, p3 As oGraph.oPicture, p4 As oGraph.oPicture
Dim t1 As oGraph.oText, t2 As oText, t3 As oText, t4 As oText
Dim s1 As oGraph.oLine, s2 As oGraph.oLine, s3 As oGraph.oLine



'Private Sub HScroll1_Change()
'    Convas1.Zoom = HScroll1.value * 5
'    Convas1.Paint
'End Sub


'Private Sub Command1_Click(index As Integer)
'    Dim Dummy As Variant, iFile As Integer, data() As Byte
'    Dim UNCFile As String
'
'    Select Case index
'
'        Case 2: 'Save
'            Dummy = Convas1.BinaryData
'            UNCFile = App.Path & "\Convas.blg"
'            iFile = FreeFile
'            Open UNCFile For Binary As iFile
'            Put iFile, , Dummy
'            Close iFile '
'        Case 3: 'Clear
'            Convas1.ClearWorkSheet
'        Case 4: 'Load
'            iFile = FreeFile
'            UNCFile = App.Path & "\Convas.blg"
'
'            Open UNCFile For Binary As iFile
'            Get iFile, , Dummy
'            Close iFile
'            'Assign the Variant to a bytearray to the bag.contents
'
'            If Len(Dummy) > 0 Then
'                'Convas1.BinaryData = ""
'                Convas1.ClearWorkSheet
'                Convas1.BinaryData = Dummy
'            End If
'            'Convas1.Paint
'        Case 5: '100 %
'            Convas1.Zoom = 100
'            Convas1.Paint
'        Case 6:
'        Case 7:
'    End Select
'End Sub


    




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




