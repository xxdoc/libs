VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ePinType
  InPin
  OutPin
End Enum

Public PinType As ePinType, X As Single, Y As Single, MouseOver As Boolean
Public ParentKey As String, Key As String
Public TargetParentKey As String, TargetKey As String

Public Sub ConnectTo(OtherPin As cPin)
  If Me.PinType = OtherPin.PinType Then Err.Raise vbObjectError, , "Pin-Types don't match"
  If Me.PinType = OutPin Then
    Me.TargetParentKey = OtherPin.ParentKey
    Me.TargetKey = OtherPin.Key
  Else 'OtherPin is of type OutPin, so we set its Targetkeys to this Pins Keys
    OtherPin.TargetParentKey = Me.ParentKey
    OtherPin.TargetKey = Me.Key
  End If
End Sub

Public Function GetConnectionString() As String
  GetConnectionString = ParentKey & ".Out_" & Key & " --> " & TargetParentKey & ".In_" & TargetKey
End Function

Public Sub DrawConnection(CC As cCairoContext, Widgets As cWidgets, Parent As cwNode, ParentPinsIdx, ParentPinsCount)
Dim TargetCont As cwNode, TargetPin As cPin, TargetPinsIdx
  If Widgets.Exists(TargetParentKey) Then Set TargetCont = Widgets(TargetParentKey) Else Exit Sub

  For Each TargetPin In TargetCont.InPins
    If TargetPin.Key = TargetKey Then
      DrawBezier CC, Parent.Widget, TargetCont.Widget, ParentPinsIdx, ParentPinsCount, TargetPinsIdx, TargetCont.InPins.Count
      Exit Sub
    End If
    TargetPinsIdx = TargetPinsIdx + 1
  Next
End Sub

Private Sub DrawBezier(CC As cCairoContext, WPar As cWidgetBase, WTgt As cWidgetBase, _
                       ParentPinsIdx, ParentPinsCount, TargetPinsIdx, TargetPinsCount)
Dim x0, y0, x1, y1
  x0 = WPar.Left + WPar.Width: y0 = WPar.Top + WPar.Height / (ParentPinsCount + 1) * (ParentPinsIdx + 1)
  x1 = WTgt.Left:              y1 = WTgt.Top + WTgt.Height / (TargetPinsCount + 1) * (TargetPinsIdx + 1)
  
  CC.MoveTo x0, y0: CC.CurveTo x0 + 80, y0, x1 - 80, y1, x1, y1
End Sub
