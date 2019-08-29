VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Default"
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   6975
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
      Height          =   375
      Left            =   8775
      TabIndex        =   3
      Top             =   6930
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   6585
      Left            =   3960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   270
      Width           =   6405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   330
      Left            =   4455
      TabIndex        =   1
      Top             =   6930
      Width           =   1500
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6630
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   11695
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents st As CSaveTree
Attribute st.VB_VarHelpID = -1

'save whatever data we want for the node however you want based on the tag value we give it (note not tied to node.tag)
Private Sub st_Serialize(n As MSComctlLib.Node, appendTag As String, ByVal index As Long)
    'in real world examples node.tag is probably an object to serialize to file then store the file name in appendTag
    appendTag = n.tag
End Sub

'restore whatever data we want for the node however you want based on the tag value we give it (note not tied to node.tag)
Private Sub st_DeSerialize(n As MSComctlLib.Node, ByVal appendTag As String, ByVal index As Long)
    'in real world example createobject from data in appendTag file path and set n.tag = reloaded obj
    'if desired we could set a progress bar based on index and st.NodeCount
    n.tag = appendTag
End Sub

'example of actual use in large project how to save/restore node icon and serialized data
'Private Sub saveTree_Serialize(n As MSComctlLib.Node, appendTag As String, ByVal index As Long)
'
'    'the treeview knows how to display target data on click
'    tvProject_NodeClick n
'    If txtCode.Text <> lastText Then
'        lastText = txtCode.Text
'        tmp = fso.GetFreeFileName(exportDir)
'        fso.writeFile tmp, lastText
'        appendTag = n.Image & ":" & fso.FileNameFromPath(tmp)
'    Else
'        appendTag = n.Image & ":" & fso.FileNameFromPath(tmp)
'    End If
'
'End Sub
'
'Private Sub saveTree_DeSerialize(n As MSComctlLib.Node, ByVal tag As String, ByVal index As Long)
'    Dim tmp() As String
'    Dim fName As String
'
'    On Error Resume Next
'
'    If InStr(tag, ":") > 0 Then
'        tmp() = Split(tag, ":", 2)
'        n.Image = CLng(tmp(0))
'        fName = tmp(1)
'    Else
'        fName = tag
'    End If
'
'    If Len(fName) > 0 Then n.tag = saveTree.BaseDir & fName
'
'    DoEvents
'
'End Sub


Private Sub Command1_Click()
    st.saveTree TreeView1, App.path & "\treeSave.txt"
    Me.Caption = "Tree saved " & Now
End Sub

Private Sub Command2_Click()
    st.RestoreTree TreeView1, App.path & "\treeSave.txt"
    ExpandAll
    Me.Caption = "Tree restored " & Now
End Sub

Private Sub Form_Load()
    Set st = New CSaveTree
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Text1 = Node.tag
End Sub

Private Sub Command3_Click()
    Dim n As Node
 
    Set n = TreeView1.Nodes.Add(, , "TopLevelBranch1", "TopLevelBranch1")
    n.tag = "Tag: " & n.Text
    n.Expanded = True
    
    Set n = TreeView1.Nodes.Add("TopLevelBranch1", tvwChild, "SubBranch1_1", "SubBranch1_1")
    n.tag = "Tag: " & n.Text
    n.Expanded = True
    
    Set n = TreeView1.Nodes.Add("SubBranch1_1", tvwChild, "Node1_1_1", "Node1_1_1")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add("SubBranch1_1", tvwChild, "Node1_1_2", "Node1_1_2")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add("SubBranch1_1", tvwChild, "Node1_1_3", "Node1_1_3")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add("TopLevelBranch1", tvwChild, "SubBranch1_2", "SubBranch1_2")
    n.tag = "Tag: " & n.Text
    n.Expanded = True
    
    Set n = TreeView1.Nodes.Add("SubBranch1_2", tvwChild, "Node1_2_1", "Node1_2_1")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add("SubBranch1_2", tvwChild, "Node1_2_2", "Node1_2_2")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add(, , "TopLevelBranch2", "TopLevelBranch2")
    n.tag = "Tag: " & n.Text
    n.Expanded = True
    
    Set n = TreeView1.Nodes.Add("TopLevelBranch2", tvwChild, "SubBranch2_1", "SubBranch2_1")
    n.tag = "Tag: " & n.Text
    
    Set n = TreeView1.Nodes.Add("TopLevelBranch2", tvwChild, "SubBranch2_2", "SubBranch2_2")
    n.tag = "Tag: " & n.Text
    n.Expanded = True
    
    Set n = TreeView1.Nodes.Add("SubBranch2_2", tvwChild, "Node2_2_1", "Node2_2_1")
    n.tag = "Tag: " & n.Text
    
End Sub

Sub ExpandAll()
    Dim n As Node
    For Each n In TreeView1.Nodes
        n.Expanded = True
    Next
End Sub
