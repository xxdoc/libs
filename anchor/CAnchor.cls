VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author:  David Zimmer <dzzie@yahoo.com>
'site:    http://sandsprite.com
'license: free for any use

Public Items As New Collection
Public WithEvents owner As Form
Attribute owner.VB_VarHelpID = -1


Sub AddItem(c As Control, Optional anchorTopLeft As Boolean = True, Optional fixedHeight As Boolean, Optional fixedWidth As Boolean)

    Dim ce As New CAnchorElement
    
    If owner Is Nothing Then
        MsgBox "Must set owner form before adding items to CAnchor class", vbInformation, "Dev Error"
        Exit Sub
    End If
    
    ce.name = c.name
    ce.anchorTopLeft = anchorTopLeft
    ce.fixedHeight = fixedHeight
    ce.fixedWidth = fixedWidth
    ce.MarginRight = owner.Width - c.Left - c.Width
    ce.MarginBottom = owner.Height - c.Top - c.Height
    
    Items.Add ce

End Sub


Private Sub owner_Resize()

    On Error Resume Next
    
    Dim c As Control
    Dim ce As CAnchorElement
    
    For Each c In owner.Controls
        For Each ce In Me.Items
            If c.name = ce.name Then
                HandleElementResize c, ce
                Exit For
            End If
        Next
    Next
        
End Sub

Private Sub HandleElementResize(c As Control, e As CAnchorElement)
    On Error Resume Next
    
    If Not e.fixedWidth Then
        c.Width = owner.Width - c.Left - e.MarginRight
    ElseIf Not e.anchorTopLeft Then
        c.Left = owner.Width - c.Width - e.MarginRight
    End If
        
    If Not e.fixedHeight Then
        c.Height = owner.Height - c.Top - e.MarginBottom
    ElseIf Not e.anchorTopLeft Then
        c.Top = owner.Height - c.Height - e.MarginBottom
    End If
    
End Sub
