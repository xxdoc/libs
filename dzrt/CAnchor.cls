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
Attribute VB_Exposed = True
'author:  David Zimmer <dzzie@yahoo.com>
'site:    http://sandsprite.com
'license: free for any use

Public Items As New Collection
Private WithEvents m_owner As Form
Attribute m_owner.VB_VarHelpID = -1

Property Let owner(v As Object)
    Set m_owner = v
    Set Items = New Collection
End Property

Sub LoadSettingsFromTags(parentForm As Object)
    
    Dim c As Control
    Dim tmp() As String
    Dim b(3) As Boolean
    
    'note if .tag is missing elements
    'the on error will let them default to false
    'lazy programming but reliable..
    On Error Resume Next
    
    Set m_owner = parentForm
    
    'Debug.Print "name  |  Tag  | Boolean  (order: anchorTopLeft,fixedHeight,fixedWidth)"
    
    For Each c In m_owner.Controls
        'If c.name = "Text2" Then Stop 'to watch on Error default to false for sure..
        If Len(c.Tag) = 0 Then
            b(0) = True
        Else
            tmp() = split(c.Tag, ",")
            For i = 0 To 2
                b(i) = IIf(tmp(i) = "1", True, False)
            Next
        End If
        'Debug.Print c.name & " " & c.Tag & " " & b(0) & " " & b(1) & " " & b(2)
        AddItem c, b(0), b(1), b(2)
        Erase tmp()
        Erase b() 'they all default to false again
    Next
    
End Sub

Sub AddItem(objControl As Object, Optional anchorTopLeft As Boolean = True, Optional fixedHeight As Boolean, Optional fixedWidth As Boolean)

    Dim ce As New CAnchorElement
    
    If m_owner Is Nothing Then
        MsgBox "Must set m_owner form before adding items to CAnchor class", vbInformation, "Dev Error"
        Exit Sub
    End If
    
    ce.name = objControl.name
    ce.anchorTopLeft = anchorTopLeft
    ce.fixedHeight = fixedHeight
    ce.fixedWidth = fixedWidth
    ce.MarginRight = m_owner.Width - objControl.Left - objControl.Width
    ce.MarginBottom = m_owner.Height - objControl.Top - objControl.Height
    
    Items.Add ce

End Sub


Private Sub m_owner_Resize()

    On Error Resume Next
    
    Dim c As Control
    Dim ce As CAnchorElement
    
    For Each c In m_owner.Controls
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
        c.Width = m_owner.Width - c.Left - e.MarginRight
    ElseIf Not e.anchorTopLeft Then
        c.Left = m_owner.Width - c.Width - e.MarginRight
    End If
        
    If Not e.fixedHeight Then
        c.Height = m_owner.Height - c.Top - e.MarginBottom
    ElseIf Not e.anchorTopLeft Then
        c.Top = m_owner.Height - c.Height - e.MarginBottom
    End If
    
End Sub
