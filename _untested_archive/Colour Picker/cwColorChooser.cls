VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwColorChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long

Event ColorChanged(PickedColour As Long)

Const NumberOfRings As Long = 9 'we actually draw one less than this (+ two half-size rings at the inner and outer edges, for white and black)
Const DonutRadius& = 280, DonutInnerRadius& = DonutRadius * 0.3, RingWidth As Long = (DonutRadius - DonutInnerRadius) / (NumberOfRings - 1)
Const DonutRenderScale As Double = 0.43

Private xCenter As Single, yCenter As Single
Private DonutSrf As cCairoSurface, ColorLut(0 To 1440 - 1) As Long
Private mHueLocked As Boolean, mLumLocked As Boolean
Private mCurrentRadius As Double, mCurrentAngleIndex As Long

Private WithEvents w As cWidgetBase
Attribute w.VB_VarHelpID = -1
Private Type RGB
   R As Long
   G As Long
   B As Long
End Type
Private Type HLS
   H As Long
   L As Long
   S As Long
End Type
Private mRGB As RGB, mHLS As HLS, mColour As Long
Private Sub Class_Initialize()
  Set w = Cairo.WidgetBase
      w.BackColor = vbWhite
      w.ImplementsWheelMessages = True
   mHLS.S = 240: mRGB.R = 255
   Set DonutSrf = CreateDonutSpectrumSurface(DonutRadius * 2, RingWidth)
   mCurrentRadius = DonutRadius - ((NumberOfRings - 1) \ 2) * RingWidth
End Sub
Public Property Let Hue(pHue As Long)
   mCurrentAngleIndex = (1440 + pHue * 6) Mod 1440
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Property
Public Property Let Luminance(pLuminance As Long)
   mCurrentRadius = DonutInnerRadius + pLuminance * (DonutRadius - DonutInnerRadius) / 240
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Property
Public Property Let Saturation(pSaturation As Long)
   If pSaturation < 0 Then pSaturation = 0
   If pSaturation > 240 Then pSaturation = 240
   mHLS.S = pSaturation
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Property
Public Property Get Widget() As cWidgetBase
  Set Widget = w
End Property
Public Property Get Widgets() As cWidgets
  Set Widgets = w.Widgets
End Property
Public Property Get Colour() As Long
  Colour = mColour
End Property
Private Sub SetColour(pColour As Long, Optional RecalcHLS As Boolean)
Dim R As Double, G As Double, B As Double, DummyH As Long
   'If mColour = pColour Then Exit Property
   mColour = pColour
   Cairo.ColorSplit mColour, R, G, B
   mRGB.R = CInt(R * 255): mRGB.G = CInt(G * 255): mRGB.B = CInt(B * 255)
   If mColour = vbBlack Then
      mHLS.L = 0
   ElseIf mColour = vbWhite Then
      mHLS.L = 240
   ElseIf R = G And G = B Then
      'not interested in the Hue the API call spits out; it's always 160 for greys. We'll use the value the user chose instead
      ColorRGBToHLS mColour, DummyH, mHLS.L, mHLS.S
   Else
      If RecalcHLS Then ColorRGBToHLS mColour, mHLS.H, mHLS.L, mHLS.S
   End If
   If Not mHueLocked Then mCurrentAngleIndex = (mHLS.H * 1440 / 240) Mod 1440
   If Not mLumLocked Then mCurrentRadius = DonutInnerRadius + mHLS.L * (DonutRadius - DonutInnerRadius) / 240
   w.Refresh
   RaiseEvent ColorChanged(mColour)
End Sub
Public Property Let Colour(pColour As Long)
   SetColour pColour, True
End Property
Public Property Let Red(pRed As Long)
   mRGB.R = pRed
   Colour = RGB(mRGB.R, mRGB.G, mRGB.B)
End Property
Public Property Let Green(pGreen As Long)
   mRGB.G = pGreen
   Colour = RGB(mRGB.R, mRGB.G, mRGB.B)
End Property
Public Property Let Blue(pBlue As Long)
   mRGB.B = pBlue
   Colour = RGB(mRGB.R, mRGB.G, mRGB.B)
End Property
Public Property Get Hue() As Long
   Hue = mHLS.H
End Property
Public Property Get Sat() As Long
   Sat = mHLS.S
End Property
Public Property Get Lum() As Long
   Lum = mHLS.L
End Property

Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim MouseRadius As Single
   
   MouseRadius = Sqr((Abs(yCenter - Y) ^ 2 + Abs(xCenter - X) ^ 2)) / DonutRenderScale
   If MouseRadius < DonutInnerRadius Then Exit Sub
   mCurrentAngleIndex = (Cairo.CalcArc(yCenter - Y, xCenter - X) * 720 / Cairo.PI + 1080) Mod 1440
   mLumLocked = MouseRadius > DonutRadius
   If Not mLumLocked Then mHueLocked = MouseRadius > DonutInnerRadius: mCurrentRadius = MouseRadius
   
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Sub
Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim MouseRadius As Single, MouseSegment As Long
   If Button <> 1 Then Exit Sub
   
   MouseRadius = Sqr((Abs(yCenter - Y) ^ 2 + Abs(xCenter - X) ^ 2)) / DonutRenderScale
   MouseSegment = (Cairo.CalcArc(yCenter - Y, xCenter - X) * 720 / Cairo.PI + 1080) Mod 1440
   
   If mHueLocked Then
      If MouseRadius < DonutInnerRadius Then
         mCurrentRadius = DonutInnerRadius
      ElseIf MouseRadius > DonutRadius Then
         mCurrentRadius = DonutRadius
      Else
         mCurrentRadius = MouseRadius
      End If
   ElseIf mLumLocked Then
      mCurrentAngleIndex = MouseSegment
   End If
   
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Sub
Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal X As Single, ByVal Y As Single)
   mHueLocked = False: mLumLocked = False
End Sub
Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
   mCurrentAngleIndex = (mCurrentAngleIndex + UBound(ColorLut) + 1 + LineOffset * 8) Mod 1440
   SetColour ColourFromPoint(mCurrentAngleIndex, mCurrentRadius)
End Sub
Private Sub W_Resize()
  xCenter = w.ScaleWidth / 2: yCenter = w.ScaleHeight / 2
  w.Refresh
End Sub

Private Sub W_Paint(CC As cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
Dim pat As cCairoPattern, i As Long, RenderSize As Single, SatSrf As cCairoSurface, tmpL As Long
   
   RenderSize = w.ScaleWidth * DonutRenderScale
   
   Set pat = Cairo.CreateLinearPattern(0, -DonutRadius * DonutRenderScale, 0, -DonutInnerRadius * DonutRenderScale)
   tmpL = mHLS.L
   For i = 0 To 20
      'Debug.Print "add stop at "; i / 20 & " for radius"; DonutInnerRadius + (DonutRadius - DonutInnerRadius) * i / 20
      pat.AddColorStop 1 - (i / 20), ColourFromPoint(mCurrentAngleIndex, DonutInnerRadius + (DonutRadius - DonutInnerRadius) * i / 20)
   Next i
   mHLS.L = tmpL
   
   DrawBorderAndBackGround CC, dx_Aligned, dy_Aligned
   
   CC.Save
      CC.TranslateDrawings xCenter + 0.5, yCenter + 0.5
      
      'the RingSurface
      Set SatSrf = DonutSrf.CreateSimilar(, , , True)
      With SatSrf.CreateContext
         .SetLineWidth 1
         .Ellipse DonutRadius, DonutRadius, DonutRadius * 2, DonutRadius * 2, True
         .SetSourceColor ColorHLSToRGB(1, 120, mHLS.S) 'Olaf! This line!
         .Operator = CAIRO_OPERATOR_HSL_SATURATION
         .Fill True
         .Operator = CAIRO_OPERATOR_OVER
         .SetSourceColor vbBlack
         .Stroke
      End With
      
      CC.RenderSurfaceContent SatSrf, -RenderSize, -RenderSize, 2 * RenderSize, 2 * RenderSize
      
      'the Luminance track...
      CC.RotateDrawingsDeg mCurrentAngleIndex / 4
      CC.MoveTo 0, -DonutRadius * DonutRenderScale
      CC.LineTo 0, -DonutInnerRadius * DonutRenderScale
      CC.SetSourceColor vbBlack, 0.7
      CC.SetLineWidth 18
      CC.Stroke True
      CC.SetLineWidth 16
      CC.Stroke , pat
    
      '...and its small black/white "inner-circle"
      CC.ARC 0, -mCurrentRadius * DonutRenderScale, 12
      CC.SetLineWidth 8
      CC.SetSourceColor vbBlack, 0.55
      CC.Stroke True 'intermediate stroke (the path is still there)
      CC.SetLineWidth 1.5
      CC.SetSourceColor vbWhite ', 0.7
      CC.Stroke 'final stroke
   
      'the outer donut 'grab-handle'
      CC.ARC 0, 0, DonutRadius * DonutRenderScale + 10, -1.7, -1.44
      CC.SetSourceColor ColorLut(mCurrentAngleIndex)
      CC.SetLineWidth 14
      CC.Stroke
   
   CC.Restore
  
   'the centered Circle-Fill with the current color
   CC.TranslateDrawings 0.5, 0.5
     CC.ARC xCenter, yCenter, -5 + DonutInnerRadius * DonutRenderScale
       CC.SetSourceColor mColour
       CC.Fill True

       CC.SetLineWidth 2.5
       CC.SetSourceColor ColorLut(mCurrentAngleIndex)
     CC.Stroke

End Sub

Private Sub DrawBorderAndBackGround(CC As cCairoContext, DX, DY)
   CC.SetLineWidth 1, True
   CC.Rectangle 0, 0, DX, DY, True
   CC.SetSourceColor w.BackColor, w.Alpha
   CC.Fill
End Sub

Private Function CreateDonutSpectrumSurface(Size, RingWidth) As cCairoSurface
Dim i&, Lst As cArrayList, LUTSrf As cCairoSurface, LUT() As Byte, R As Double, j As Long
   Set LUTSrf = Cairo.CreateSurface(1440, 1) 'a long, but only one Pixel high "Stripe-Surface" (1/4 of a degree -> 360*4=1440)
   With Cairo.CreateLinearPattern(0, 0, LUTSrf.Width - 1, 0)
      Set Lst = New_c.ArrayList(vbLong, vbRed, vbYellow, vbGreen, vbCyan, vbBlue, vbMagenta, vbRed)
      For i = 0 To Lst.Count - 1
         .AddColorStop i / 6, Lst(i)
      Next i
      LUTSrf.CreateContext.Paint 1, .This
   End With
  
   Set CreateDonutSpectrumSurface = Cairo.CreateSurface(Size, Size)
   With CreateDonutSpectrumSurface.CreateContext
      .SetLineCap CAIRO_LINE_CAP_ROUND
      .SetLineWidth 1.5 '(Size - RingWidth - 0.5) * 2 * Cairo.PI / LUTSrf.Width
      .TranslateDrawings Size \ 2, Size \ 2
       LUTSrf.BindToArray LUT
       
      'our white outer half-ring
      .Ellipse 0, 0, Size * 2, Size * 2, True
      .SetSourceColor vbWhite
      .Fill
    
      For j = NumberOfRings - 1 To 1 Step -1
         R = DonutInnerRadius + j * RingWidth - RingWidth / 2
         For i = 0 To 1440 - 1
            ColorLut(i) = RGB(LUT(i * 4 + 2, 0), LUT(i * 4 + 1, 0), LUT(i * 4, 0))
            .SetSourceColor ColourFromPoint(i, R, True)
            .MoveTo 0, -R + RingWidth
            .RelLineTo 0, -RingWidth
            .Stroke
            .RotateDrawingsDeg 0.25 'rotate by one quarter of a degree (360 / (1440 steps in the LUT))
         Next i
      Next j
       
      LUTSrf.ReleaseArray LUT
      
      'erase half the width of our inner-most circle by painting over it
      .ARC xCenter, yCenter, DonutInnerRadius
      .SetSourceColor vbWhite
      .Fill
  End With

End Function
Private Function ColourFromPoint(pAngleIndex As Long, ByVal pRadius As Double, Optional Imprecise As Boolean) As Long
Dim thisRing As Long

   mHLS.H = pAngleIndex / 4 / 360 * 240
   
   If pRadius > DonutRadius Then pRadius = DonutRadius
   
   If Imprecise Then
      'Imprecise fetches the ring colour drawn at the radius (ignoring the saturation, i.e. always 240) rather than a precise colour
      thisRing = (pRadius - DonutInnerRadius) \ RingWidth
      pRadius = DonutInnerRadius + thisRing * RingWidth
   End If
   mHLS.L = (pRadius - DonutInnerRadius) / (DonutRadius - DonutInnerRadius) * 240
   If mHLS.S = 0 Then
      'when Sat is 0, ColorHLSToRGB spits out a Luminance of zero for some reason. Since we're talking shades of grey, I'll do it myself...
      ColourFromPoint = RGB(mHLS.L / 240 * 255, mHLS.L / 240 * 255, mHLS.L / 240 * 255)
   Else
      ColourFromPoint = ColorHLSToRGB(mHLS.H, mHLS.L, mHLS.S)
   End If
   
End Function
