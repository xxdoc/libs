VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1
Private Canvas As cwCanvas, N(1 To 8) As cwNode

Private Sub Class_Initialize()
  Set Form = Cairo.WidgetForms.Create(vbSizable, "Main-Form with connectable Widgets (Tab-Key for Focus, MouseWheel for Zoom)", , 900, 700)
      Form.IconImageKey = "FormIco"
      Form.WidgetRoot.ImageKey = "BgBlack"
      Form.WidgetRoot.ImageKeyRenderBehaviour = ImgKeyRenderRepeat
      
  Set Canvas = Form.Widgets.Add(New cwCanvas, "Canvas", 0, 0, Form.ScaleWidth, Form.ScaleHeight)

  Set N(1) = AddNode("Node1", 60, 40, "NodeIco")
  Set N(2) = AddNode("Node2", 60, 120, "NodeIco")
  Set N(3) = AddNode("Node3", 60, 380, "NodeIco")
  Set N(4) = AddNode("Node4", 160, 490, "NodeIco")
  Set N(5) = AddNode("Node5", 290, 200, "NodeIco")
  Set N(6) = AddNode("Node6", 290, 300, "NodeIco")
  Set N(7) = AddNode("Node7", 515, 420, "NodeIco")
  Set N(8) = AddNode("Node8", 730, 220, "NodeIco")
  
  'Add and connect Pins from Node 1 and Node 2 to Node 5
  N(1).PinAdd("Pin1", OutPin).ConnectTo N(5).PinAdd("Pin1", InPin)
  N(2).PinAdd("Pin1", OutPin).ConnectTo N(5).PinAdd("Pin2", InPin)
  
  'Add and connect Pins from Node 3 and Node 4 to Node 6
  N(3).PinAdd("Pin1", OutPin).ConnectTo N(6).PinAdd("Pin1", InPin)
  N(4).PinAdd("Pin1", OutPin).ConnectTo N(6).PinAdd("Pin2", InPin)
  
  'Add and connect Pins from Node 7 to Nodes 5, 6 and 4 (now in "reverse", from InPins to OutPins)
  N(7).PinAdd("Pin1", InPin).ConnectTo N(5).PinAdd("Pin1", OutPin)
  N(7).PinAdd("Pin2", InPin).ConnectTo N(6).PinAdd("Pin1", OutPin)
  N(7).PinAdd("Pin3", InPin).ConnectTo N(4).PinAdd("Pin2", OutPin)
  
  'Add and connect Pins from Node 8 to Nodes 1 and 7
  N(8).PinAdd("Pin1", InPin).ConnectTo N(1).PinAdd("Pin2", OutPin)
  N(8).PinAdd("Pin2", InPin).ConnectTo N(7).PinAdd("Pin2", OutPin)
End Sub

Public Function AddNode(Key As String, X, Y, ImageKey) As cwNode
  Set AddNode = Canvas.Widgets.Add(New cwNode, Key, X, Y, 96, 68)
      AddNode.Widget.ToolTip = Key
      AddNode.Widget.ToolTipDelay = 300
      AddNode.Widget.ImageKey = ImageKey
End Function
 
Private Sub Form_ResizeWithDimensions(ByVal NewWidth As Long, ByVal NewHeight As Long)
  Canvas.Widget.Move 0, 0, NewWidth / Form.WidgetRoot.Zoom, NewHeight / Form.WidgetRoot.Zoom
End Sub

Private Sub Form_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
Dim Zoom: Zoom = Form.WidgetRoot.Zoom + 0.1 * Sgn(Rotation)
  If Zoom < 0.5 Then Zoom = 0.5 Else If Zoom > 2 Then Zoom = 2 'set some Limits for that
  Form.WidgetRoot.Zoom = Zoom
  Canvas.Widget.Move 0, 0, Form.ScaleWidth / Zoom, Form.ScaleHeight / Zoom
  Form.WidgetRoot.Refresh
End Sub
 
