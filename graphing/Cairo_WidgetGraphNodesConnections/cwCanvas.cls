VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public ActivePin As cPin 'determinded by HitTesting in W_MouseMove

Private Sub Class_Initialize()
  Set W = Cairo.WidgetBase
      W.ForwardFocus = True
      W.BackColor = -1 'no BackColor (telling the WidgetEngine, to make this Widget use the BackGround of its Parent
End Sub
 
Public Property Get Widget() As cWidgetBase: Set Widget = W: End Property
Public Property Get Widgets() As cWidgets: Set Widgets = W.Widgets: End Property
 
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
  DrawConnectionPaths W.MeasureContext, True, X, Y 'perform HitTesting using the W.MeasureContext
  If ActivePin Is Nothing Then W.ToolTip = "" Else W.ToolTip = ActivePin.GetConnectionString
End Sub

Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
  W.Refresh
End Sub
Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
  W.Refresh
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
  DrawConnectionPaths CC
End Sub

Private Sub DrawConnectionPaths(CC As cCairoContext, Optional ByVal HitTesting As Boolean, Optional ByVal X, Optional ByVal Y)
Dim Node As cwNode, Pin As cPin, OutPinsIdx As Long, PinInStroke As cPin
  CC.SetLineWidth IIf(HitTesting, 8, 2) 'make the "stroke-thickness" a bit larger (+- 4 Pixels) in case of HitTesting
    For Each Node In Widgets
      OutPinsIdx = 0
      For Each Pin In Node.OutPins
        Pin.DrawConnection CC, Widgets, Node, OutPinsIdx, Node.OutPins.Count
        
        If HitTesting Then 'use the CCs InStroke-Method to check, if the x,y-Coords are "on the Path"
          If CC.InStroke(X, Y) Then Set PinInStroke = Pin
          CC.ClearPath
        Else 'normal coloring of the Path per Stroke-Command
          CC.Stroke , Cairo.CreateSolidPatternLng(IIf(Pin Is ActivePin And W.MouseOver, vbGreen, &HBBBBBB))
        End If
        
        OutPinsIdx = OutPinsIdx + 1
      Next
    Next
  If HitTesting And Not ActivePin Is PinInStroke Then Set ActivePin = PinInStroke: W.Refresh
End Sub
