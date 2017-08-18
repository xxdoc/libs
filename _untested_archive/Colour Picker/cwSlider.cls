VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event ValueChanged(ByVal NewValue As Double)

Private Radius As Double
Private mTrackStartX As Double, mTrackWidth As Double
Private mMax As Double
Private mValue As Double

Private mCaption As String

Private Controlpoints As cControlPoints
Private WithEvents CenterDrag As cControlPoint
Attribute CenterDrag.VB_VarHelpID = -1
Private SliderTrack As cControlPoint
Attribute SliderTrack.VB_VarHelpID = -1

Private WithEvents w As cWidgetBase
Attribute w.VB_VarHelpID = -1

Public Sub Init(ByVal sCaption As String, ByVal Min As Double, Optional ByVal Max As Double = 100, Optional ByVal pValue As Double)
Dim CaptionSpace As Long

   Radius = w.Height \ 2

   mCaption = sCaption
   If Len(mCaption) Then
      With Cairo.CreateSurface(w.Width, w.Height).CreateContext
          .SelectFont "Tahoma", 9
          CaptionSpace = .GetTextExtents(mCaption & "  ")
      End With
   End If
   
   mTrackStartX = Radius + CaptionSpace
   mTrackWidth = w.Width - CaptionSpace - (2 * Radius)
    
   mMax = Max
    
   SliderTrack.Move (mTrackStartX + mTrackWidth / 2), (w.Height / 2)
   SliderTrack.Radius = (2 * Radius)
   SliderTrack.SetRectangularStyle mTrackWidth + (2 * Radius), (3 * Radius)
    
   CenterDrag.Move mTrackStartX, (w.Height / 2)
   CenterDrag.Radius = Radius + 0.5
   Value = pValue

End Sub
Public Property Let Max(ByVal NewValue As Double)
    If mMax = NewValue Then Exit Property
    mMax = NewValue
End Property
Public Property Get Value() As Double
    Value = mValue
End Property

Public Property Let Value(ByVal NewValue As Double)
      
   If NewValue = mValue Then Exit Property
   If NewValue < 0 Then NewValue = 0
   If NewValue > mMax Then NewValue = mMax
   mValue = NewValue
     
   If mMax = 0 Then
       CenterDrag.X = mTrackStartX
   Else
       CenterDrag.X = mTrackStartX + mValue * mTrackWidth / mMax
   End If
   w.Refresh
   RaiseEvent ValueChanged(mValue)

End Property
Private Sub DrawControlBody(CC As cCairoContext)
   With CC
      .Rectangle mTrackStartX - Radius * 0.5, (w.Height - Radius) / 2, mTrackWidth + Radius, Radius
      .SetSourceColor RGB(80, 80, 80)
      .Fill
   End With
End Sub
Private Sub CenterDrag_PositionChanging(NewX As Double, NewY As Double)
Dim SliderIntervalPixels As Double
    
    If mMax = 0 Then
        SliderIntervalPixels = mTrackWidth
    Else
        SliderIntervalPixels = mTrackWidth / mMax
    End If
    NewY = w.Height / 2
    If NewX < mTrackStartX Then NewX = mTrackStartX
    If NewX > mTrackStartX + mTrackWidth Then NewX = mTrackStartX + mTrackWidth
  
    Value = CInt((NewX - mTrackStartX) / SliderIntervalPixels)
End Sub

Private Sub Class_Initialize()
   Set w = Cairo.WidgetBase
   Set Controlpoints = New_c.Controlpoints
   Set CenterDrag = Controlpoints.Add("CenterDrag", 0, 0, , , 0, True)
   Set SliderTrack = Controlpoints.Add("SliderTrack", 0, 0, , , 0, True)
   w.CanGetFocus = False
End Sub
Public Property Get Widget() As cWidgetBase
   Set Widget = w
End Property
Public Property Get Widgets() As cWidgets
   Set Widgets = w.Widgets
End Property
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim CP As cControlPoint
    
   If Not Button = vbLeftButton Then Exit Sub
   
   Set CP = Controlpoints.CheckControlPointUnderCursor(X, Y)
   If Not CP Is Nothing Then
      CenterDrag_PositionChanging CDbl(X), CDbl(Y)
      CenterDrag.SetMouseDownPoint X, Y
   End If
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim CP As cControlPoint, MOverStateChanged As Boolean
    Set CP = Controlpoints.CheckControlPointUnderCursor(X, Y, True, MOverStateChanged)
    If MOverStateChanged Then w.Refresh
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim CP As cControlPoint
    Set CP = Controlpoints.CheckControlPointUnderCursor(X, Y)
    Controlpoints.EnsureMouseUpState
    w.Refresh
End Sub
Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal DX As Single, ByVal DY As Single, UserObj As Object)
   Draw CC
End Sub

Private Sub Draw(CC As cCairoContext)
   With CC
      '.SetSourceColor vbRed:.Paint
      
      'slider track
      .Rectangle mTrackStartX - Radius * 0.5, (w.Height - Radius) / 2, mTrackWidth + Radius, Radius
      .SetSourceColor RGB(80, 80, 80)
      .Fill
      
      'captions
      .SelectFont "Tahoma", 9, &H404040, True
      .DrawText 0, 0, 100, w.Height, mCaption, , vbLeftJustify, , True
      
      'Slider Ball
      .Save
         .TranslateDrawings CenterDrag.X, CenterDrag.Y
         .ARC 0, 0, Radius - 3
         .SetSourceColor &H404040
         .SetLineWidth 6
         .Stroke True
         .SetSourceColor vbWhite
         .SetLineWidth 1.5
         .Stroke
      .Restore
       
       Controlpoints.Draw CC
   End With
   
End Sub

