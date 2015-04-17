VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Save Image"
      Height          =   465
      Left            =   7830
      TabIndex        =   3
      Top             =   315
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Demo Graph"
      Height          =   465
      Left            =   4635
      TabIndex        =   2
      Top             =   270
      Width           =   2760
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   5055
      Left            =   180
      ScaleHeight     =   4995
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   4860
      Width           =   4515
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   990
      Width           =   11625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'note dot.Validate doesnt seem to work?

Dim img As BinaryImage

Private Sub Command1_Click()
   
   Dim g As New CGraph
   Dim n0 As CNode, n1 As CNode, n2 As CNode, n3 As CNode, n4 As CNode, n5 As CNode
   
   Set n0 = g.AddNode("this is my" & vbCrLf & "multiline\nnode")
   n0.shape = "box"
   n0.style = "filled"
   n0.color = "lightyellow"
   n0.fontcolor = "#c0c0c0"
   
   Set n1 = g.AddNode
   Set n2 = g.AddNode
   Set n3 = g.AddNode
   Set n4 = g.AddNode
   Set n5 = g.AddNode
   
   n0.ConnectTo n2
   n1.ConnectTo n2
   n2.ConnectTo n3
   n1.ConnectTo n4
   n0.ConnectTo n5
   
   Call g.GenerateGraph
   Text1.Text = g.lastGraph
   
   Set img = g.dot.ToGIF(g.lastGraph)
   If img Is Nothing Then Exit Sub
   
   Set Picture1.Picture = img.Picture
    
End Sub

Private Sub Command2_Click()

   If img Is Nothing Then Exit Sub
   
   pth = App.Path & "\sample.gif"
   If img.Save(pth) Then
        MsgBox "Saved to " & pth, vbInformation
   Else
        MsgBox "Save failed", vbExclamation
   End If
   
   'or SavePicture Picture1, App.Path & "\sample.bmp"
    
End Sub



'some sample dot files:

'digraph G {A -> B -> C -> D;}

'digraph G {
'    size ="4,4";
'    main [shape=box]; /*
'    this is a comment
'    */
'    main -> parse [weight=8];
'    parse -> execute;
'    main -> init [style=dotted];
'    main -> cleanup;
'    execute -> { make_string; printf}
'    init -> make_string;
'    edge [color=red]; // so is this
'    main -> printf [style=bold,label="100 times"];
'    make_string [label="make a\nstring"];
'    node [shape=box,style=filled,color=".7 .3 1.0"];
'    execute -> compare;
'}

'digraph G {
'
'    subgraph cluster_0 {
'        style=filled;
'        color=lightgrey;
'        node [style=filled,color=white];
'        a0 -> a1 -> a2 -> a3;
'        label = "process #1";
'    }
'
'    subgraph cluster_1 {
'        node [style=filled];
'        b0 -> b1 -> b2 -> b3;
'        label = "process #2";
'        color = blue
'    }
'    start -> a0;
'    start -> b0;
'    a1 -> b3;
'    b2 -> a3;
'    a3 -> a0;
'    a3 -> end;
'    b3 -> end;
'
'    start [shape=Mdiamond];
'    end [shape=Msquare];
'}
