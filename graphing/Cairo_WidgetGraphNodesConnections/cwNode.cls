VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase, dx As Single, dy As Single, Pin As cPin
Attribute W.VB_VarHelpID = -1
 
Const Padding = 13
Public InPins As cCollection, OutPins As cCollection, ActivePin As cPin

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.ForeColor = &HBBBBBB
      W.Alpha = 0.8
      W.Moveable = True
  Set InPins = New_c.Collection(False)
  Set OutPins = New_c.Collection(False)
End Sub

Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property
 
Public Function PinAdd(Key As String, ByVal PinType As ePinType) As cPin
Set PinAdd = New cPin
    PinAdd.ParentKey = W.Key
    PinAdd.Key = Key
    PinAdd.PinType = PinType
    IIf(PinType = InPin, InPins, OutPins).Add PinAdd, Key
End Function

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As cWidgetBase)
  For Each Pin In InPins:  Pin.MouseOver = False: Next
  For Each Pin In OutPins: Pin.MouseOver = False: Next
  W.Refresh
End Sub
 
Private Sub W_DblClick()
  MsgBox "DblClick-Event on: " & W.Key
End Sub
 
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Set Pin = PinHitTest(X, Y, InPins)
  If Pin Is Nothing Then Set Pin = PinHitTest(X, Y, OutPins)
  If ActivePin Is Pin Then Exit Sub Else Set ActivePin = Pin: W.Refresh
  If Pin Is Nothing Then W.ToolTip = W.Key: Exit Sub
  W.ToolTip = W.Key & vbCrLf & IIf(Pin.PinType, "Out", "In") & "Pin: " & Pin.Key
End Sub

Private Function PinHitTest(X, Y, Pins As cCollection) As cPin
  For Each Pin In Pins
    Pin.MouseOver = W.MouseOver And Abs(Pin.X + Padding \ 2 - X - 1) < Padding \ 2 And Abs(Pin.Y + Padding \ 2 - Y) < Padding \ 2
    If Pin.MouseOver Then Set PinHitTest = Pin
  Next
End Function

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  dx = dx_Aligned: dy = dy_Aligned
  Draw CC
  DrawPins CC, InPins
  DrawPins CC, OutPins
End Sub

Private Sub Draw(CC As cCairoContext)
  W.SelectionColor = IIf(W.Focused, vbMagenta, IIf(W.MouseOver, vbCyan, W.BorderColor))
  CC.SetLineWidth 2
  CC.RoundedRect Padding, 0, dx - 2 * Padding, dy, Padding \ 2, True
  CC.Stroke , Cairo.CreateSolidPatternLng(W.SelectionColor, W.Alpha)

  W.SelectFontSettingsInto CC
  CC.DrawText 0, 4, dx, dy, W.Key, True, vbCenter
 
  If Len(W.ImageKey) Then CC.RenderSurfaceContent W.ImageKey, (dx - 48) / 2, (dy + 12 - 48) / 2, 48, 48
End Sub

Private Sub DrawPins(CC As cCairoContext, Pins As cCollection)
Dim i, Pin As cPin
  For Each Pin In Pins
    i = i + 1
    Pin.X = IIf(Pins Is OutPins, dx - Padding + 1, 0)
    Pin.Y = dy / (Pins.Count + 1) * i - Padding / 2
    W.FocusColor = IIf(Pin.MouseOver, vbYellow, vbRed)
    Cairo.Theme.DrawTo CC, W, thmTypeArrow, 2, Pin.X, Pin.Y, Padding, Padding, , thmDirectionRight
  Next
End Sub
