VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TvwNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ShadowTreeView As TreeView
Private PropNode As Collection

Private Sub Class_Initialize()
Set PropNode = New Collection
End Sub

Private Sub Class_Terminate()
Set PropNode = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropNode.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As TreeView)
Set ShadowTreeView = Object
End Sub

Public Function Add(Optional ByVal Relative As Variant, Optional ByVal Relationship As TvwNodeRelationshipConstants = TvwNodeRelationshipNext, Optional ByVal Key As String, Optional ByVal Text As String, Optional ByVal Image As Long, Optional ByVal SelectedImage As Long) As TvwNode
Attribute Add.VB_Description = "Adds a node to the nodes collection."
If Image < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
If SelectedImage < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim NewNode As New TvwNode, RelativeNode As TvwNode
If Not IsMissing(Relative) Then
    Select Case VarType(Relative)
        Case vbString, vbLong, vbInteger, vbByte
            Set RelativeNode = PropNode(Relative)
        Case Else
            Err.Raise 13
    End Select
End If
NewNode.FInit ShadowTreeView, Key, Image, SelectedImage
If Key = vbNullString Then
    PropNode.Add NewNode
Else
    PropNode.Add NewNode, Key
End If
NewNode.Index = PropNode.Count
ShadowTreeView.FNodesAdd NewNode, Key, RelativeNode, Relationship, Text, Image, SelectedImage
Set Add = NewNode
End Function

Public Property Get Item(ByVal Index As Variant) As TvwNode
Attribute Item.VB_Description = "Returns a node given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropNode(Index)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of nodes in the collection."
Count = PropNode.Count
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Removes all nodes from the collection."
ShadowTreeView.FNodesClear
Set PropNode = New Collection
End Sub

Public Sub Remove(ByVal Index As Variant)
Attribute Remove.VB_Description = "Removes a node from the collection."
ShadowTreeView.FNodesRemove PropNode(Index).Handle
PropNode.Remove Index
Dim i As Long
For i = 1 To PropNode.Count
    Me.Item(i).Index = i
Next i
End Sub
