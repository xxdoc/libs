VERSION 5.00
Begin VB.UserControl Text 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   FillStyle       =   0  'Solid
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Text.ctx":0000
   ScaleHeight     =   1080
   ScaleWidth      =   2220
   Windowless      =   -1  'True
End
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'--------------------------Object Declaration----------------------------------
Private Enum PointerLocation
    None = 0
    TopLine = 1
    BottomLine = 2
    LeftLine = 3
    RightLine = 4
    TopLeftLine = 5
    TopRightLine = 6
    BottomLeftLine = 7
    BottomRightLine = 8
End Enum

Private m_PointerLocation As PointerLocation
Private WithEvents m_ConvasPicBox As PictureBox
Attribute m_ConvasPicBox.VB_VarHelpID = -1

Private m_IGraphics As IGraphics
Private m_Region As Long
Private m_RegionOuter As Long
Private m_Path As Long
Private m_PathText As Long
Private m_PathBack As Long
Private m_Height As Single
Private m_Width As Single
Private m_MouseX As Double
Private m_MouseY As Double
Private m_TempX As Double
Private m_TempY As Double
Private m_TempPx1() As Double
Private m_TempPy1() As Double
Private m_TempPx2() As Double
Private m_TempPy2() As Double
Private m_TempUbound  As Single
Private m_CentreOldX As Double
Private m_CentreOldY As Double
Private m_ResizingObject As Boolean
Private m_ChangesAccepted As Boolean
'--------------------------Private Variables-----------------------------------
Private m_Selected As Boolean
Private m_NodeKey As String
Private m_ToolTipText As String
Private m_Visible As Boolean
'--------------------------Control Events---------------------------
Public Event UnSelectAll(ByVal pExceptNodeKey As String)
Public Event RightClick(ByVal Key As String, ByVal X As Double, ByVal Y As Double)
Public Event Click(ByVal Key As String)
Public Event DoubleClick(ByVal Key As String)
Public Event MouseDown(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event MouseMove(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event MouseUp(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event Selected(ByVal Key As String)
Public Event Deleted(ByVal Key As String)
Private m_RaiseFromInside As Boolean
Private m_blnPointMoving As Boolean
Private m_WidthChanged As Boolean
Private m_HeightChanged As Boolean
Private m_DblClicked As Boolean
'--------------------------Events End---------------------------------
'Default Property Values:
Const m_def_ForeColor = &HF000000
Const m_def_IsTransparent = 0
Const m_def_Transparency = 255
Const m_def_IsOutlined = 1
Const m_def_BackColor = &HFFFFFF

Const m_def_CentreX = 0
Const m_def_CentreY = 0
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_IsTransparent As Boolean
Dim m_Transparency As Byte
Dim m_IsOutlined As Boolean
Dim m_BackColor As OLE_COLOR


Private m_Font As Font
Private m_Caption As String
Dim m_CentreX As Double
Dim m_CentreY As Double
'
'
''--------------------------Properties---------------------------------
Public Property Let tHeight(ByVal pHeight As Single)
    m_Height = pHeight
    m_HeightChanged = True
    If Not m_IGraphics Is Nothing Then
        Call CreatePath
        Call CreateRegion
    End If
    m_HeightChanged = False
    Me.Paint
    PropertyChanged "tHeight"
End Property

Public Property Get tHeight() As Single
    tHeight = m_Height
End Property

Public Property Let tWidth(ByVal pWidth As Single)
    m_Width = pWidth
    m_WidthChanged = True
    If Not m_IGraphics Is Nothing Then
        Call CreatePath
        Call CreateRegion
    End If
    m_WidthChanged = False
    Me.Paint
    PropertyChanged "tWidth"
End Property

Public Property Get tWidth() As Single
    tWidth = m_Width
End Property


Public Property Let oSelected(ByVal pSelected As Boolean)
    m_Selected = pSelected
    If m_RaiseFromInside Then
        RaiseEvent Selected(m_NodeKey)
        Me.Paint
        m_RaiseFromInside = False
    End If
End Property

Public Property Get oSelected() As Boolean
    oSelected = m_Selected
End Property

Public Property Let NodeKey(ByVal pNodeKey As String)
    m_NodeKey = pNodeKey
End Property

Public Property Get NodeKey() As String
    NodeKey = m_NodeKey
End Property
'
Public Property Let ToolTipText(ByVal pToolTipText As String)
    m_ToolTipText = pToolTipText
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = "TextProp"
    ToolTipText = m_ToolTipText
End Property


Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal pVisible As Boolean)
    m_Visible = pVisible
    If m_Visible Then
        Me.Paint
    End If
End Property

'--------------------------Properties End-------------------------------

Private Sub CreatePath()
    Dim pRect As RECTF
    Dim rct As RECTF
    Dim fontFam          As Long
    Dim curFont          As Long
    Dim strFormat        As Long
    Dim box              As RECTF
    Dim rct2             As RECTF
    Dim FS               As Long
    Dim IsAvailable      As Long

    If m_Path <> 0 Then
        GdipDeletePath (m_Path)
    End If


    If m_PathBack <> 0 Then
        GdipDeletePath (m_PathBack)
    End If

    If m_PathText <> 0 Then
        GdipDeletePath (m_PathText)
    End If

    Call GdipCreatePath(FillModeWinding, m_PathText)
    Call GdipCreatePath(FillModeWinding, m_PathBack)
    Call GdipCreatePath(FillModeWinding, m_Path)

    'Text stuff follows

   ' Set the Text Rendering Quality

   ' Create a font family object to allow us to create a font
   ' We have no font collection here, so pass a NULL for that parameter

   GdipCreateFontFamilyFromName m_Font.Name, 0, fontFam
   GdipIsStyleAvailable fontFam, FS, IsAvailable
   If IsAvailable = 0 Then
      Dim Msg              As String
      Msg = "Font family " & m_Font.Name & " not available under GDI+.Please select another font."
      MsgBox Msg, vbExclamation, "Font not supported"
      Me.Font = Ambient.Font
      Exit Sub
   End If
   ' Create the font from the specified font family name
   ' >> Note that we have changed the drawing Unit from pixels to points!!
   FS = FontStyleRegular 'not really needed since it's zero
   If m_Font.Bold Then FS = FS + FontStyleBold
   If m_Font.Italic Then FS = FS + FontStyleItalic
   If m_Font.Strikethrough Then FS = FS + FontStyleStrikeout
   If m_Font.Underline Then FS = FS + FontStyleUnderline

   GdipCreateFont fontFam, m_Font.Size, FS, UnitPoint, curFont
   ' Create the StringFormat object
   ' We can pass NULL for the flags and language id if we want
   GdipCreateStringFormat 0, 0, strFormat

   ' Justify each line of text
   GdipSetStringFormatAlign strFormat, StringAlignmentCenter

   ' Justify the block of text (top to bottom) in the rectangle.
   GdipSetStringFormatLineAlign strFormat, StringAlignmentCenter

   GdipMeasureString m_IGraphics.GetGraphicsHandle, m_Caption, -1, curFont, _
   rct, strFormat, pRect, 0, 0
   LSet rct = pRect
   m_ChangesAccepted = True
   If m_Height < rct.Height Then
        m_Height = rct.Height
        m_ChangesAccepted = False
   End If

   If Not (m_HeightChanged Or m_WidthChanged) Then
        m_Height = rct.Height
        m_ChangesAccepted = False
   End If

   rct.Width = rct.Width - rct.Width / 5

   If Not (m_HeightChanged Or m_WidthChanged) Then
        m_Width = rct.Width
        m_ChangesAccepted = False
   End If

   If m_Width < rct.Width Then
        m_Width = rct.Width
        m_ChangesAccepted = False
   End If

   rct.Height = m_Height
   rct.Width = m_Width
   rct.Top = m_CentreY / 15 - m_Height / 2
   rct.Left = m_CentreX / 15 - rct.Width / 2


   GdipAddPathString m_PathText, _
      m_Caption, -1, _
      fontFam, _
      FS, _
      m_Font.Size, _
      rct, _
      strFormat
   Call GdipAddPathRectangle(m_PathBack, rct.Left, rct.Top, rct.Width, rct.Height)
   If m_IsOutlined Then
     Call GdipAddPathRectangle(m_Path, rct.Left, rct.Top, rct.Width, rct.Height)
   End If

   GdipDeleteStringFormat strFormat
   GdipDeleteFont curFont
   GdipDeleteFontFamily fontFam

End Sub


Private Sub CreateRegion()
    Dim pRect As RECTF
    If m_Region <> 0 Then
        GdipDeleteRegion m_Region
        m_Region = 0
    End If

    If m_RegionOuter <> 0 Then
        GdipDeleteRegion m_RegionOuter
    End If

    If m_Height = 0 Or m_Width = 0 Then CreatePath

    With pRect
        .Height = m_Height
        .Width = m_Width
        .Top = m_CentreY / 15 - .Height / 2
        .Left = m_CentreX / 15 - .Width / 2
    End With
    Call GdipCreateRegionRect(pRect, m_Region)
    Call InflateRectF(pRect, -2, -2)
    Call GdipCreateRegionRect(pRect, m_RegionOuter)
    Call GdipCombineRegionRegion(m_RegionOuter, m_Region, CombineModeXor)
End Sub

Public Sub Paint()
    Dim pPen As Long
    Dim pTextPen As Long
    Dim pWidth As Single
    Dim pTextvbColor As Long
    Dim pColor As Colors
    Dim pTextBrush As Long
    Dim pSelectedBrush As Long
    Dim pRect As RECTF
    If m_Visible Then
        If m_IGraphics Is Nothing Then Call InitilizeEnviornment
        If m_IGraphics Is Nothing Then Exit Sub
        If m_Selected Then
            pWidth = 3
            pColor = DarkBlue
        Else
            pWidth = 1
            pColor = Black
        End If

        Call GdipCreatePen1(pColor, pWidth, UnitPixel, pPen)
        If m_IsTransparent Then
            Call GdipCreateSolidFill(GetGDIColorFromOLE(m_BackColor, m_Transparency), pTextBrush)
        Else
            Call GdipCreateSolidFill(GetGDIColorFromOLE(m_BackColor), pTextBrush)
        End If
            
        Call GdipCreateSolidFill(GetGDIColorFromOLE(&H8000000D, 140), pSelectedBrush)

        Call GdipCreatePen1(Colors.Black, 1, UnitPixel, pTextPen)
        Call GdipSetPenLineJoin(pPen, LineJoinRound)
        Call GdipDrawPath(m_IGraphics.GetGraphicsHandle, pPen, m_Path)
        If m_Selected Then
            Call GdipFillPath(m_IGraphics.GetGraphicsHandle, pTextBrush, m_PathBack)
            Call GdipFillPath(m_IGraphics.GetGraphicsHandle, pSelectedBrush, m_Path)

            Call GdipDeleteBrush(pSelectedBrush)
        Else
            Call GdipFillPath(m_IGraphics.GetGraphicsHandle, pTextBrush, m_PathBack)
        End If
        Call GdipDrawPath(m_IGraphics.GetGraphicsHandle, pTextPen, m_Path)

        Call GdipDeleteBrush(pTextBrush)
        Call GdipCreateSolidFill(GetGDIColorFromOLE(m_ForeColor), pTextBrush)
        Call GdipFillPath(m_IGraphics.GetGraphicsHandle, pTextBrush, m_PathText)
        
        'Call GdipFillRegion(m_IGraphics.GetGraphicsHandle, pTextBrush, m_RegionOuter)

        Call GdipDeletePen(pPen)
        Call GdipDeletePen(pTextPen)
        Call GdipDeleteBrush(pTextBrush)
        m_IGraphics.RePaintConvas

    End If
End Sub


Private Sub m_ConvasPicBox_DblClick()
    m_DblClicked = True
End Sub

Private Sub m_ConvasPicBox_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And m_Selected Then
        RaiseEvent Deleted(m_NodeKey)
    End If
End Sub

Private Sub m_ConvasPicBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pResult  As Long
    m_MouseX = 0
    m_MouseY = 0
    
    If m_Visible Then
        If m_IGraphics.IsConvasLocked Then Exit Sub
        If m_Region <> 0 Then
            Call GdipIsVisibleRegionPoint(m_Region, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)

            If pResult = 1 Or (m_IGraphics.IsSelected(m_NodeKey) And Shift = vbCtrlMask) Then

                m_MouseX = X / 15
                m_MouseY = Y / 15
                Call AlignPointerLocation(X, Y)
                If Shift = vbCtrlMask Then
                    m_PointerLocation = None
                End If
                If (m_PointerLocation = None Or Shift = vbCtrlMask) And Button = vbLeftButton Then
                    m_blnPointMoving = True
                    m_ResizingObject = False

                    If Not (m_IGraphics.IsSelected(m_NodeKey) And Shift = vbCtrlMask) Then

                    Else
                        m_CentreOldX = m_CentreX
                        m_CentreOldY = m_CentreY
                    End If
                    If Not m_Selected Then
                        m_RaiseFromInside = True
                        Me.oSelected = True
                    End If
                    RaiseEvent Click(m_NodeKey)
                    If Button = vbRightButton Then
                        RaiseEvent RightClick(m_NodeKey, X, Y)
                    End If
                    RaiseEvent MouseDown(m_NodeKey, Button, Shift, CDbl(X), CDbl(Y))
                Else
                    If Not m_Selected Then
                        m_RaiseFromInside = True
                        Me.oSelected = True
                    End If
                    m_ResizingObject = True
                    RaiseEvent MouseDown(m_NodeKey, Button, Shift, CDbl(X), CDbl(Y))
                End If
            Else
                    If m_Selected Then
                        m_RaiseFromInside = True
                        If Shift <> vbCtrlMask Then
                            Me.oSelected = False
                        End If
                    End If
            End If

        End If

    End If
End Sub

Private Sub m_ConvasPicBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pResult  As Long
    Dim pShiftX As Double
    Dim pShiftY As Double


    pShiftX = m_CentreX - m_MouseX * 15
    pShiftY = m_CentreY - m_MouseY * 15

    If m_Visible Then
        If m_Region <> 0 And Not m_blnPointMoving And Button = 0 Then
            Call GdipIsVisibleRegionPoint(m_Region, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)
            If pResult = 1 Then
               Call GdipIsVisibleRegionPoint(m_RegionOuter, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)

               If pResult = 1 Then

                    If Shift <> vbCtrlMask Then
                          If m_ConvasPicBox.Tag = m_NodeKey Then
                                m_ConvasPicBox.ToolTipText = m_ToolTipText
                                m_ConvasPicBox.MousePointer = GetPointer(X, Y)
                          End If
                    Else
                         m_ConvasPicBox.ToolTipText = ""
                         m_ConvasPicBox.MousePointer = MousePointerConstants.vbCustom
                    End If
               Else
                    m_TempX = X
                    m_TempY = Y
                    m_ConvasPicBox.Tag = m_NodeKey
                    If Shift <> vbCtrlMask Then
                         m_ConvasPicBox.ToolTipText = m_ToolTipText
                         m_ConvasPicBox.MousePointer = MousePointerConstants.vbSizeAll 'GetPointer(X / 15, Y / 15)
                    Else
                         m_ConvasPicBox.ToolTipText = ""
                         m_ConvasPicBox.MousePointer = MousePointerConstants.vbCustom
                    End If
               End If
               RaiseEvent MouseMove(m_NodeKey, Button, Shift, CDbl(X + pShiftX), CDbl(Y + pShiftY))
            Else
                    If Shift <> vbCtrlMask Then
                        If m_ConvasPicBox.Tag = m_NodeKey Then
                            m_ConvasPicBox.ToolTipText = ""
                            m_ConvasPicBox.MousePointer = vbDefault
                        End If
                    Else
                         m_ConvasPicBox.ToolTipText = ""
                         m_ConvasPicBox.MousePointer = MousePointerConstants.vbCustom
                    End If

           End If
        ElseIf m_blnPointMoving And Button = vbLeftButton Then
           ' If pLastX <> X And pLastY <> Y Then
    
                m_TempX = X
                m_TempY = Y
                Call DrawTempLine(X / 15, Y / 15)

                RaiseEvent MouseMove(m_NodeKey, Button, Shift, CDbl(X + pShiftX), CDbl(Y + pShiftY))

            'End If
        Else
            
            If m_ResizingObject Then
                m_TempX = X
                m_TempY = Y
                Call DrawTempLine(X / 15, Y / 15)

                RaiseEvent MouseMove(m_NodeKey, Button, Shift, CDbl(X + pShiftX), CDbl(Y + pShiftY))
            Else

                m_TempX = 0
                m_TempY = 0
            End If
        End If
    End If

End Sub

Private Function GetPointer(ByVal X As Single, Y As Single) As MousePointerConstants
    Call AlignPointerLocation(X, Y)
    Select Case m_PointerLocation
        Case PointerLocation.None:
            GetPointer = vbDefault
        Case PointerLocation.LeftLine:
            GetPointer = vbSizeWE
        Case PointerLocation.RightLine:
            GetPointer = vbSizeWE
        Case PointerLocation.TopLine:
            GetPointer = vbSizeNS
        Case PointerLocation.BottomLine:
            GetPointer = vbSizeNS
        Case PointerLocation.TopLeftLine:
            GetPointer = vbSizeNWSE
        Case PointerLocation.TopRightLine:
            GetPointer = vbSizeNESW
        Case PointerLocation.BottomLeftLine:
            GetPointer = vbSizeNESW
        Case PointerLocation.BottomRightLine:
            GetPointer = vbSizeNWSE
    End Select
End Function

Private Sub AlignPointerLocation(ByVal X As Single, ByVal Y As Single)
    Dim pRect As RECTF
    pRect.Left = m_CentreX / 15 - m_Width / 2
    pRect.Top = m_CentreY / 15 - m_Height / 2
    pRect.Height = m_Height
    pRect.Width = m_Width
    X = X / 15
    Y = Y / 15
    Call InflateRectF(pRect, -2, -2)
    If X < pRect.Left Then
        If Y < pRect.Top Then
            m_PointerLocation = TopLeftLine
        ElseIf Y > (pRect.Top + pRect.Height) Then
            m_PointerLocation = BottomLeftLine
        Else
            m_PointerLocation = LeftLine
        End If
    ElseIf X > (pRect.Left + pRect.Width) Then
        If Y < pRect.Top Then
            m_PointerLocation = TopRightLine
        ElseIf Y > (pRect.Top + pRect.Height) Then
            m_PointerLocation = BottomRightLine
        Else
            m_PointerLocation = RightLine
        End If
    Else
        If Y < pRect.Top Then
            m_PointerLocation = TopLine
        ElseIf Y > (pRect.Top + pRect.Height) Then
            m_PointerLocation = BottomLine
        Else
            m_PointerLocation = None
        End If
    End If
End Sub

Private Function DrawLine(ByRef pLineIndex As Single, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, Optional ByVal ErasePresious As Boolean = True) As LINE_POINT()


        Dim pIndex As Single
        Dim pSearchX As Double
        Dim pSearchY As Double
        Dim pFound As Boolean
        If X1 = X2 And Y1 = Y2 And X1 = Y1 And X1 = 0 Then

            ReDim pPoints(m_TempUbound + 1)
            pIndex = 0
            pLineIndex = 0
            pFound = False
            m_ConvasPicBox.DrawMode = DrawModeConstants.vbCopyPen
            If m_TempUbound > 0 Then
                For pLineIndex = 0 To m_TempUbound

                    m_ConvasPicBox.Line (m_TempPx1(pLineIndex) * 15, m_TempPy1(pLineIndex) * 15)-(m_TempPx2(pLineIndex) * 15, m_TempPy2(pLineIndex) * 15)
                Next
                If ErasePresious Then
                    pIndex = 1
                    CentreX = (IIf(m_TempPx1(pIndex) < m_TempPx2(pIndex), m_TempPx1(pIndex), m_TempPx2(pIndex)) + Abs(m_TempPx1(pIndex) - m_TempPx2(pIndex)) / 2) * 15
                    pIndex = 0
                    CentreY = (IIf(m_TempPy1(pIndex) < m_TempPy2(pIndex), m_TempPy1(pIndex), m_TempPy2(pIndex)) + Abs(m_TempPy1(pIndex) - m_TempPy2(pIndex)) / 2) * 15
                    Call CreateRegion
                    Call CreatePath
                End If
            Else
                'pPoints = m_Points

            End If
            Erase m_TempPx1
            Erase m_TempPy1
            Erase m_TempPx2
            Erase m_TempPy2
            m_TempUbound = 0
            'DrawLine = pPoints
            Exit Function
        End If

        If m_TempUbound <= pLineIndex Then
            ReDim Preserve m_TempPx1(pLineIndex)
            ReDim Preserve m_TempPy1(pLineIndex)
            ReDim Preserve m_TempPx2(pLineIndex)
            ReDim Preserve m_TempPy2(pLineIndex)
            m_TempUbound = UBound(m_TempPx1)
        End If

        m_ConvasPicBox.DrawMode = DrawModeConstants.vbInvert
        If ErasePresious Then m_ConvasPicBox.Line (m_TempPx1(pLineIndex) * 15, m_TempPy1(pLineIndex) * 15)-(m_TempPx2(pLineIndex) * 15, m_TempPy2(pLineIndex) * 15)
        m_TempPx1(pLineIndex) = X1
        m_TempPy1(pLineIndex) = Y1
        m_TempPx2(pLineIndex) = X2
        m_TempPy2(pLineIndex) = Y2
        m_ConvasPicBox.Line (X1 * 15, Y1 * 15)-(X2 * 15, Y2 * 15)
        pLineIndex = pLineIndex + 1

End Function


Private Sub DrawTempLine(ByVal X As Double, ByVal Y As Double)
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double
    Dim pShiftX As Double
    Dim pShiftY As Double
    Dim pLineIndex As Single
    pShiftX = m_CentreX / 15 - m_MouseX
    pShiftY = m_CentreY / 15 - m_MouseY

    Select Case m_PointerLocation
        Case PointerLocation.LeftLine:
            Call DrawLine(0, X + pShiftX - m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX - m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(1, X + pShiftX - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, True)
            Call DrawLine(2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(3, X + pShiftX - m_Width / 2, m_CentreY / 15 + m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
        Case PointerLocation.RightLine:
            Call DrawLine(0, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(1, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 - m_Height / 2, True)
            Call DrawLine(2, X + pShiftX + m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(3, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
        Case PointerLocation.TopLine:
            Call DrawLine(0, m_CentreX / 15 - m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(1, m_CentreX / 15 - m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY - m_Height / 2, True)
            Call DrawLine(2, m_CentreX / 15 + m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(3, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
        Case PointerLocation.BottomLine:
            Call DrawLine(0, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 - m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(1, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, True)
            Call DrawLine(2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(3, m_CentreX / 15 - m_Width / 2, Y + pShiftY + m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY + m_Height / 2, True)
        Case PointerLocation.TopLeftLine:
            Call DrawLine(0, X + pShiftX - m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX - m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(1, X + pShiftX - m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY - m_Height / 2, True)
            Call DrawLine(2, m_CentreX / 15 + m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(3, X + pShiftX - m_Width / 2, m_CentreY / 15 + m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
        Case PointerLocation.TopRightLine:
            Call DrawLine(0, m_CentreX / 15 - m_Width / 2, Y + pShiftY - m_Height / 2, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(1, m_CentreX / 15 - m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY - m_Height / 2, True)
            Call DrawLine(2, X + pShiftX + m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
            Call DrawLine(3, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 + m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 + m_Height / 2, True)
        Case PointerLocation.BottomLeftLine:
            Call DrawLine(0, X + pShiftX - m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX - m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(1, X + pShiftX - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, True)
            Call DrawLine(2, m_CentreX / 15 + m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(3, X + pShiftX - m_Width / 2, Y + pShiftY + m_Height / 2, m_CentreX / 15 + m_Width / 2, Y + pShiftY + m_Height / 2, True)
        Case PointerLocation.BottomRightLine:
            Call DrawLine(0, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, m_CentreX / 15 - m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(1, m_CentreX / 15 - m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX + m_Width / 2, m_CentreY / 15 - m_Height / 2, True)
            Call DrawLine(2, X + pShiftX + m_Width / 2, m_CentreY / 15 - m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(3, m_CentreX / 15 - m_Width / 2, Y + pShiftY + m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY + m_Height / 2, True)
        Case PointerLocation.None
            Call DrawLine(0, X + pShiftX - m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX - m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(1, X + pShiftX - m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY - m_Height / 2, True)
            Call DrawLine(2, X + pShiftX + m_Width / 2, Y + pShiftY - m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY + m_Height / 2, True)
            Call DrawLine(3, X + pShiftX - m_Width / 2, Y + pShiftY + m_Height / 2, X + pShiftX + m_Width / 2, Y + pShiftY + m_Height / 2, True)
    End Select

End Sub


Private Sub m_ConvasPicBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pShiftX As Single
    Dim pShiftY As Single
    Dim pResult As Long
    If m_Visible Then
        If m_DblClicked Then
            Call GdipIsVisibleRegionPoint(m_Region, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)
            If pResult = 1 Then
                RaiseEvent DoubleClick(m_NodeKey)
            End If
            m_DblClicked = False
        End If
        If m_blnPointMoving Then
            Call DrawLine(0, 0, 0, 0, 0, True)
            m_blnPointMoving = False
            m_ResizingObject = False
            Me.Paint
            RaiseEvent MouseUp(m_NodeKey, Button, Shift, CDbl(X), CDbl(Y))
        ElseIf m_ResizingObject Then
            Call DrawLine(0, 0, 0, 0, 0, False)
            pShiftX = (X / 15 - m_MouseX)
            pShiftY = (Y / 15 - m_MouseY)

            Select Case m_PointerLocation
                Case TopLine:
                    m_Height = m_Height - pShiftY
                    m_Width = m_Width - pShiftX
                    pShiftY = pShiftY / 2
                    m_CentreY = m_CentreY + pShiftY * 15
                Case LeftLine:
                    m_Height = m_Height - pShiftY
                    m_Width = m_Width - pShiftX
                    pShiftX = pShiftX / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                Case TopLeftLine:
                    m_Height = m_Height - pShiftY
                    m_Width = m_Width - pShiftX
                    pShiftX = pShiftX / 2
                    pShiftY = pShiftY / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                    m_CentreY = m_CentreY + pShiftY * 15
                Case RightLine:
                    m_Width = m_Width + pShiftX
                    pShiftX = pShiftX / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                Case BottomLine:
                    m_Height = m_Height + pShiftY
                    pShiftY = pShiftY / 2
                    m_CentreY = m_CentreY + pShiftY * 15
                Case BottomRightLine:
                    m_Height = m_Height + pShiftY
                    m_Width = m_Width + pShiftX
                    pShiftX = pShiftX / 2
                    pShiftY = pShiftY / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                    m_CentreY = m_CentreY + pShiftY * 15
                Case TopRightLine
                    m_Height = m_Height - pShiftY
                    m_Width = m_Width + pShiftX
                    pShiftX = pShiftX / 2
                    pShiftY = pShiftY / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                    m_CentreY = m_CentreY + pShiftY * 15
                Case BottomLeftLine:
                    m_Height = m_Height + pShiftY
                    m_Width = m_Width - pShiftX
                    pShiftX = pShiftX / 2
                    pShiftY = pShiftY / 2
                    m_CentreX = m_CentreX + pShiftX * 15
                    m_CentreY = m_CentreY + pShiftY * 15
                Case Else
                    m_Height = m_Height + pShiftY
                    m_Width = m_Width + pShiftX
            End Select


            m_blnPointMoving = False
            m_ResizingObject = False
            m_WidthChanged = True
            m_HeightChanged = True
            m_ChangesAccepted = True
            Call CreatePath
            Call CreateRegion
            Me.Paint
            RaiseEvent MouseUp(m_NodeKey, Button, Shift, CDbl(X), CDbl(Y))
        End If
    End If
End Sub


Public Sub ReSizeConvas()
        Call InitilizeEnviornment
        Call CreatePath
        Call CreateRegion
End Sub

'--------------------------Control Handling---------------------------------
Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Print UserControl.Name
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim pHeight As Single
    Dim pWidth As Single
    Call InitilizeEnviornment
    m_Caption = PropBag.ReadProperty("Caption", "<Annotation>")
    Set m_Font = PropBag.ReadProperty("Font", UserControl.Font)
    m_CentreX = PropBag.ReadProperty("CentreX", m_def_CentreX)
    m_CentreY = PropBag.ReadProperty("CentreY", m_def_CentreY)
    m_Visible = PropBag.ReadProperty("Visible", False)
    pHeight = PropBag.ReadProperty("tHeight", 0)
    pWidth = PropBag.ReadProperty("tWidth", 0)
    If pHeight <> 0 Then
        Me.tHeight = pHeight
    End If
    If pWidth <> 0 Then
        Me.tWidth = pWidth
    End If
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_IsTransparent = PropBag.ReadProperty("IsTransparent", m_def_IsTransparent)
    m_Transparency = PropBag.ReadProperty("Transparency", m_def_Transparency)
    m_IsOutlined = PropBag.ReadProperty("IsOutlined", m_def_IsOutlined)
    m_Selected = PropBag.ReadProperty("oSelected", False)
    m_NodeKey = PropBag.ReadProperty("NodeKey", "")
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", "")
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 500
    UserControl.Width = 500
End Sub

Private Sub UserControl_Terminate()
    Call BeforeTerminate
End Sub

Friend Sub BeforeTerminate()
    If m_Path <> 0 Then
        Call GdipDeletePath(m_Path)
        m_Path = 0
    End If
    
    If m_Region <> 0 Then
        Call GdipDeleteRegion(m_Region)
        m_Region = 0
    End If
    
    If m_RegionOuter <> 0 Then
        Call GdipDeleteRegion(m_RegionOuter)
        m_RegionOuter = 0
    End If
    
    If m_PathBack <> 0 Then
        Call GdipDeletePath(m_PathBack)
        m_PathBack = 0
    End If
    
    If m_PathText <> 0 Then
        Call GdipDeletePath(m_PathText)
        m_PathText = 0
    End If
    
    Set m_Font = Nothing
    Set m_IGraphics = Nothing
    Set m_ConvasPicBox = Nothing
 End Sub
 
'--------------------------Control Handling End---------------------------------
'--------------------------Standard Functions-----------------------------------
Friend Sub GetCordinates(ByRef CenX As Double, ByRef CenY As Double, ByRef oWidth As Double, ByRef oHeight As Double, Optional ByVal ResetValues As Boolean)
    If ResetValues Then
        m_TempX = 0
        m_TempY = 0
        m_MouseX = 0
        m_MouseY = 0
    End If
    CenX = m_TempX + (m_CentreX - m_MouseX * 15)
    CenY = m_TempY + (m_CentreY - m_MouseY * 15)
    oWidth = m_Width * 15
    oHeight = m_Height * 15
End Sub


Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "TextProp"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
End Property
'
Public Sub Activate()
    Call CreatePath
    Call CreateRegion
End Sub

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    m_HeightChanged = False
    m_WidthChanged = False
    Call CreatePath
    Call CreateRegion
    Me.Paint
    PropertyChanged "Font"
End Property


Public Property Get CentreX() As Double
    CentreX = m_CentreX
End Property

Public Property Let CentreX(ByVal New_CentreX As Double)
    m_CentreX = New_CentreX
    PropertyChanged "CentreX"
End Property

Public Property Get CentreY() As Double
    CentreY = m_CentreY
End Property

Public Property Let CentreY(ByVal New_CentreY As Double)
    m_CentreY = New_CentreY
    PropertyChanged "CentreY"
End Property

Private Sub UserControl_InitProperties()
    m_CentreX = m_def_CentreX
    m_CentreY = m_def_CentreY
    m_BackColor = m_def_BackColor
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_IsTransparent = m_def_IsTransparent
    m_Transparency = m_def_Transparency
    m_IsOutlined = m_def_IsOutlined
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, "<Annotation>")
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("CentreX", m_CentreX, m_def_CentreX)
    Call PropBag.WriteProperty("CentreY", m_CentreY, m_def_CentreY)
    Call PropBag.WriteProperty("tHeight", m_Height, 0)
    Call PropBag.WriteProperty("tWidth", m_Width, 0)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("IsTransparent", m_IsTransparent, m_def_IsTransparent)
    Call PropBag.WriteProperty("Transparency", m_Transparency, m_def_Transparency)
    Call PropBag.WriteProperty("IsOutlined", m_IsOutlined, m_def_IsOutlined)
    Call PropBag.WriteProperty("NodeKey", m_NodeKey, "")
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, "")
    Call PropBag.WriteProperty("oSelected", m_Selected, False)
    Call PropBag.WriteProperty("Visible", m_Visible, False)
    
End Sub

Private Sub InitilizeEnviornment()
    Dim pConvas As Convas
    On Error Resume Next
    If (UserControl.Ambient.UserMode) Then
        Set pConvas = UserControl.Parent
        Set m_IGraphics = pConvas
        Set m_ConvasPicBox = m_IGraphics.GetConvasPicBox
        Set pConvas = Nothing
    End If
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get IsTransparent() As Boolean
Attribute IsTransparent.VB_ProcData.VB_Invoke_Property = "TextProp"
    IsTransparent = m_IsTransparent
End Property

Public Property Let IsTransparent(ByVal New_IsTransparent As Boolean)
    m_IsTransparent = New_IsTransparent
    PropertyChanged "IsTransparent"
End Property

Public Property Get Transparency() As Byte
Attribute Transparency.VB_ProcData.VB_Invoke_Property = "TextProp"
    Transparency = m_Transparency
End Property

Public Property Let Transparency(ByVal New_Transparency As Byte)
    m_Transparency = New_Transparency
    PropertyChanged "Transparency"
End Property

Public Property Get IsOutlined() As Boolean
Attribute IsOutlined.VB_ProcData.VB_Invoke_Property = "TextProp"
    IsOutlined = m_IsOutlined
End Property

Public Property Let IsOutlined(ByVal New_IsOutlined As Boolean)
    m_IsOutlined = New_IsOutlined
    PropertyChanged "IsOutlined"
End Property
