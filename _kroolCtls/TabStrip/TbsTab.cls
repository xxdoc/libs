VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbsTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ObjectPointer As Long
Private PropIndex As Long, PropKey As String
Private PropTag As String
Private PropToolTipText As String

Private Function ShadowTabStrip() As TabStrip
Set ShadowTabStrip = PtrToObj(ObjectPointer)
End Function

Friend Sub FInit(ByVal Object As TabStrip, ByVal Key As String)
ObjectPointer = ObjPtr(Object)
PropKey = Key
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the tab."
Index = PropIndex
End Property

Friend Property Let Index(ByVal Value As Long)
PropIndex = Value
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Get Tag() As String
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
Tag = PropTag
End Property

Public Property Let Tag(ByVal Value As String)
PropTag = Value
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the caption of the tab."
Attribute Caption.VB_UserMemId = 0
Caption = ShadowTabStrip.FTabCaption(PropIndex)
End Property

Public Property Let Caption(ByVal Value As String)
ShadowTabStrip.FTabCaption(PropIndex) = Value
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the tool tip text of the tab. Requires that the show tips property is set to true."
ToolTipText = PropToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
PropToolTipText = Value
End Property

Public Property Get Image() As Long
Attribute Image.VB_Description = "Returns/sets the index of an image in an associated image list control."
Image = ShadowTabStrip.FTabImage(PropIndex)
End Property

Public Property Let Image(ByVal Value As Long)
If Value < 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
ShadowTabStrip.FTabImage(PropIndex) = Value
End Property

Public Property Get Selected() As Boolean
Attribute Selected.VB_Description = "Returns/sets a value that determines whether the tab is selected or not."
Selected = ShadowTabStrip.FTabSelected(PropIndex)
End Property

Public Property Let Selected(ByVal Value As Boolean)
ShadowTabStrip.FTabSelected(PropIndex) = Value
End Property

Public Property Get Pressed() As Boolean
Attribute Pressed.VB_Description = "Returns/sets a value that determines whether the tab is pressed or not. This is only meaningful if the style property is set to buttons or flat buttons."
Pressed = ShadowTabStrip.FTabPressed(PropIndex)
End Property

Public Property Let Pressed(ByVal Value As Boolean)
ShadowTabStrip.FTabPressed(PropIndex) = Value
End Property

Public Property Get HighLighted() As Boolean
Attribute HighLighted.VB_Description = "Returns/sets a value that determines whether the tab is highlighted or not. This has no visible effect if the version of comctl32.dll is 6.0 or higher and the visual styles property is set to true."
HighLighted = ShadowTabStrip.FTabHighLighted(PropIndex)
End Property

Public Property Let HighLighted(ByVal Value As Boolean)
ShadowTabStrip.FTabHighLighted(PropIndex) = Value
End Property

Public Property Get Left() As Single
Attribute Left.VB_Description = "Returns the distance between the internal left edge of an object and the left edge of its container."
Left = ShadowTabStrip.FTabLeft(PropIndex)
End Property

Public Property Let Left(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Top() As Single
Attribute Top.VB_Description = "Returns the distance between the internal top edge of an object and the top edge of its container."
Top = ShadowTabStrip.FTabTop(PropIndex)
End Property

Public Property Let Top(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Width() As Single
Attribute Width.VB_Description = "Returns the width of the tab."
Width = ShadowTabStrip.FTabWidth(PropIndex)
End Property

Public Property Let Width(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property

Public Property Get Height() As Single
Attribute Height.VB_Description = "Returns the height of the tab."
Height = ShadowTabStrip.FTabHeight(PropIndex)
End Property

Public Property Let Height(ByVal Value As Single)
Err.Raise Number:=383, Description:="Property is read-only"
End Property
