VERSION 5.00
Begin VB.UserControl Line 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   840
   ToolboxBitmap   =   "Line.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------Object Declaration----------------------------------
Private WithEvents m_ConvasPicBox As PictureBox
Attribute m_ConvasPicBox.VB_VarHelpID = -1

Private WithEvents m_ConnectedFrom As Graph.Picture
Attribute m_ConnectedFrom.VB_VarHelpID = -1

Private WithEvents m_ConnectedTo As Graph.Picture
Attribute m_ConnectedTo.VB_VarHelpID = -1

Private Type LineBand
    BandCur() As POINTF
    BandPre() As POINTF
End Type

Private m_IGraphics As IGraphics
'--------------------------Private Variables-----------------------------------
Private m_Selected As Boolean
Private m_NodeKey As String
Private m_ToolTipText As String
Private m_ConstraintIndex As Long
Private m_StepName As String
Private m_Visible As Boolean
Private m_LayerLineType As LineType
Private m_Points() As LINE_POINT
Private m_Path As Long
Private m_Pen As Long
Private m_PointCount As Long
Private m_Region As Long
Private m_LineWidth As Single
Private m_MovingPointIndex As Single
Private m_blnPointMoving As Boolean
Private m_TempPoints() As LINE_POINT
Private m_TempPointCount As Single
Private m_X1 As Double
Private m_X2 As Double
Private m_Y1 As Double
Private m_Y2 As Double
Private m_XShift As Double
Private m_YShift As Double
Private m_MovingLinePoints As Boolean
Private m_AllLinesForThreeStateDeleted As Boolean
Private m_NeedToReversePoints As Boolean
Private m_RealinedOnce As Boolean
Private m_PointRemoved As Boolean
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
Private m_PointerArrowDrawn As Boolean
Dim m_FromX As Double
Dim m_FromY As Double
Private m_XStart As Double
Private m_YStart As Double
Private m_XEnd As Double
Private m_YEnd As Double
Private m_TempToPx1() As Double
Private m_TempToPy1() As Double
Private m_TempToPx2() As Double
Private m_TempToPy2() As Double
Private m_TempToUbound  As Single

Private m_TempFromPx1() As Double
Private m_TempFromPy1() As Double
Private m_TempFromPx2() As Double
Private m_TempFromPy2() As Double
Private m_TempFromUbound  As Single

Private m_AllocatedSideKeyFrom As String
Private m_AllocatedSideKeyTo As String
Private m_RemovedPoint As LINE_POINT
Private m_LastToMouseX As Double
Private m_LastToMouseY As Double
Private m_LastFromMouseX As Double
Private m_LastFromMouseY As Double
Private m_RaiseFromInside As Boolean
Private m_DblClicked As Boolean
Private m_ToKey As String
Private m_FromKey As String
Private m_Use3D As Boolean
Private m_PathCircle As Long
Private m_PathArrow As Long
Private m_BrushCircle As Long
Private m_BrushArrow As Long
Private m_BrushVerticle As Long
Private m_BrushHorizontal As Long
Private m_3DSegmentPath() As Long

'--------------------------Events End---------------------------------

'--------------------------Properties---------------------------------
Public Property Let oSelected(ByVal pSelected As Boolean)
    m_Selected = pSelected
    If m_RaiseFromInside Then
        RaiseEvent Selected(m_NodeKey)
        m_RaiseFromInside = False
        Me.Paint
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

Public Property Let Use3D(ByVal pUse3D As Boolean)
    m_Use3D = pUse3D
End Property

Public Property Get Use3D() As Boolean
    Use3D = m_Use3D
End Property

Public Property Let ToolTipText(ByVal pToolTipText As String)
    m_ToolTipText = pToolTipText
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Get StepName() As String
    StepName = m_StepName
End Property

Public Property Let StepName(ByVal pStepName As String)
    m_StepName = pStepName
End Property

Public Property Get ConstraintIndex() As Long
    ConstraintIndex = m_ConstraintIndex
End Property

Public Property Let ConstraintIndex(ByVal pConstraintIndex As Long)
    m_ConstraintIndex = pConstraintIndex
End Property

Public Property Get ObjectType() As AMSObjectType
    ObjectType = oLine
End Property

Public Property Get ConnectedFrom() As Graph.Picture
    Set ConnectedFrom = m_ConnectedFrom
End Property

Public Property Set ConnectedFrom(ByRef pLayedObject As Graph.Picture)
    Set m_ConnectedFrom = pLayedObject
    If Not pLayedObject Is Nothing Then
        m_FromKey = pLayedObject.NodeKey
    Else
        m_FromKey = ""
    End If
End Property

Public Property Get ConnectedTo() As Graph.Picture
    Set ConnectedTo = m_ConnectedTo
End Property

Public Property Set ConnectedTo(ByRef pLayedObject As Graph.Picture)
    Set m_ConnectedTo = pLayedObject
    If Not pLayedObject Is Nothing Then
        m_ToKey = pLayedObject.NodeKey
    Else
        m_ToKey = ""
    End If
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

Public Property Get LayerLineType() As LineType
    LayerLineType = m_LayerLineType
End Property

Public Property Let LayerLineType(ByVal New_LayerLineType As LineType)
    m_LayerLineType = New_LayerLineType
    PropertyChanged "LayerLineType"
End Property
'--------------------------Properties End-------------------------------

Private Sub m_ConnectedFrom_Activate(ByVal XCentre As Double, ByVal YCentre As Double)
    m_X1 = XCentre
    m_Y1 = YCentre
    Call StartLine
End Sub

Private Sub m_ConnectedFrom_MouseDown(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    
    If Button = vbLeftButton Then
        If Shift = vbCtrlMask And m_Selected Then
            m_MovingLinePoints = False
            Exit Sub
        End If
        m_NeedToReversePoints = False
        m_AllLinesForThreeStateDeleted = False
        m_XShift = m_ConnectedFrom.CentreX - X
        m_YShift = m_ConnectedFrom.CentreY - Y
        m_MovingPointIndex = 0
        m_MovingLinePoints = True
    End If
End Sub

Private Sub m_ConnectedFrom_MouseMove(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    
    Dim pCenX As Double
    Dim pCenY As Double
    Dim pSide As Single
    Dim pX As Double, pY As Double
   
    If Shift = vbCtrlMask And m_Selected Then
       m_MovingLinePoints = False
       Exit Sub
    End If
    If m_MovingLinePoints And Button = vbLeftButton Then
        If m_TempFromUbound > 0 Then
            pX = m_TempFromPx2(1)
            pY = m_TempFromPy2(1)
        Else
            pX = m_Points(1).X
            pY = m_Points(1).Y
        End If
        
        pSide = GetAppropriateSide(True, pX, pY)
        pX = m_Points(m_PointCount - 1).X
        pY = m_Points(m_PointCount - 1).Y
        
        Call m_ConnectedFrom.GetExactCordinates(m_NodeKey, pSide, m_LineWidth, m_AllocatedSideKeyFrom, pX, pY)
     
        X = pX
        Y = pY
        
        
        Beep
        m_NeedToReversePoints = True
        Call DrawTempLinesFromFromPoint(X / 15, Y / 15)
        m_LastFromMouseX = X
        m_LastFromMouseY = Y
    End If
End Sub

Private Sub m_ConnectedFrom_MouseUp(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim pPoints() As LINE_POINT
    Dim pRPoints() As LINE_POINT
    Dim pExactPoints() As LINE_POINT
    
    If m_MovingLinePoints Then
        pPoints = DrawLineFrom(0, 0, 0, 0, 0)
        If m_NeedToReversePoints Then
            pRPoints = ReversePoints(pPoints)
            Me.SetPoints pRPoints
        Else
            Me.SetPoints pPoints
        End If
        Me.Paint
    End If
    m_NeedToReversePoints = False
    m_MovingLinePoints = False
End Sub

Private Function GetAppropriateSide(ByVal IsFromPoint As Boolean, ByVal X As Double, ByVal Y As Double) As Single
    Dim pCentX As Double
    Dim pCentY As Double
    Dim pWidth As Double
    Dim pHeight As Double
    If IsFromPoint Then
        Call m_ConnectedFrom.GetCordinates(pCentX, pCentY, pWidth, pHeight)
    Else
        Call m_ConnectedTo.GetCordinates(pCentX, pCentY, pWidth, pHeight)
    End If
    If X * 15 < (pCentX + pWidth) And X * 15 > (pCentX - pWidth) Then '2 or 4
        If Y * 15 < pCentY Then
            GetAppropriateSide = 2
        Else
            GetAppropriateSide = 4
        End If
    Else    ' 1 or 3
        If X * 15 < pCentX Then
            GetAppropriateSide = 1
        Else
            GetAppropriateSide = 3
        End If
    End If
End Function


Private Sub m_ConnectedFrom_ReAlign(ByVal ShiftX As Double, ByVal ShiftY As Double)
    Dim pIndex As Single
    If m_Selected Then
        If Not m_RealinedOnce Then
            Debug.Print "Align From " & m_ConnectedFrom.NodeKey & " : " & m_NodeKey
            If Not (ShiftX = ShiftY And ShiftX = 0) Then
                For pIndex = 0 To m_PointCount - 1
                    m_Points(pIndex).X = m_Points(pIndex).X + ShiftX / 15
                    m_Points(pIndex).Y = m_Points(pIndex).Y + ShiftY / 15
                Next
                Call CalculatePointCount
                Call CreatePath
                
                m_RealinedOnce = True
            End If
        Else
            m_RealinedOnce = False
        End If
    End If
End Sub

Private Sub m_ConnectedFrom_Selected(ByVal Key As String)
    
    If m_ConnectedTo.oSelected And m_ConnectedFrom.oSelected Then
        Me.oSelected = True
    Else
        If m_Selected Then m_Selected = False
    End If
End Sub

Private Sub m_ConnectedTo_Activate(ByVal XCentre As Double, ByVal YCentre As Double)
     m_X2 = XCentre
     m_Y2 = YCentre
     Call StartLine
End Sub

Private Function ReversePoints(ByRef pPoints() As LINE_POINT) As LINE_POINT()
    Dim pRPoints() As LINE_POINT
    Dim pUbound As Single
    Dim pIndex As Single
    pUbound = UBound(pPoints)
    ReDim pRPoints(pUbound)
    For pIndex = 0 To pUbound
        Let pRPoints(pIndex) = pPoints(pUbound - pIndex)
    Next
    ReversePoints = pRPoints
End Function

Public Sub Paint()
    Dim pBrush As Long
    Dim pIndex As Single
    Dim pDirectionCode As DirectionCode
    Dim pPointCount As Single
    Dim pLinePoints() As LINE_POINT
    Dim ptPoint As LINE_POINT
    Dim pPreviousDirCode As DirectionCode
    
    
    If Not m_Use3D Then
        If m_Path = 0 Then Call CreatePath
    Else
        If m_PathArrow = 0 Or m_PathCircle = 0 Then Call CreatePath
        If m_PointCount > 0 Then
            ReDim pLinePoints(m_PointCount - 1)
        Else
            ReDim pLinePoints(m_PointCount)
        End If
        For pIndex = 0 To m_PointCount - 1
            LSet pLinePoints(pIndex) = m_Points(pIndex)
        Next
        
        If m_PointRemoved Then
            ReDim Preserve pLinePoints(m_PointCount)
            LSet pLinePoints(m_PointCount) = m_RemovedPoint
        End If
        pPointCount = UBound(pLinePoints)
    End If
    If m_Visible Then
        If m_IGraphics Is Nothing Then Call InitilizeEnviornment
        If Not m_Use3D Then
            Call CreatePen(m_LineWidth)
            Call GdipDrawPath(m_IGraphics.GetGraphicsHandle, m_Pen, m_Path)
        Else
            For pIndex = 0 To pPointCount '- 1
                If pIndex > 0 Then
                    If pLinePoints(pIndex).X = pLinePoints(pIndex - 1).X Then 'Verticle
                        If pLinePoints(pIndex).Y > pLinePoints(pIndex - 1).Y Then
                            pDirectionCode = U2D
                        Else
                            pDirectionCode = D2U
                        End If
                    Else    'Horizontal
                        If pLinePoints(pIndex).X > pLinePoints(pIndex - 1).X Then
                            pDirectionCode = L2R
                        Else
                            pDirectionCode = R2L
                        End If
                    End If
                End If
                
                Select Case pIndex:
                    Case 0:
                    Case pPointCount:
                        If pDirectionCode = L2R Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.X = ptPoint.X - m_LineWidth * 2.5
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, True, pPreviousDirCode)
                        ElseIf pDirectionCode = R2L Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.X = ptPoint.X + m_LineWidth * 2.5
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, True, pPreviousDirCode)
                        ElseIf pDirectionCode = U2D Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.Y = ptPoint.Y - m_LineWidth * 2.5
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, True, pPreviousDirCode)
                        ElseIf pDirectionCode = D2U Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.Y = ptPoint.Y + m_LineWidth * 2.5
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, True, pPreviousDirCode)
                        End If
                        Call DrawArrow(False, pDirectionCode, pLinePoints(pIndex))
                    Case Else
                        If pIndex = 1 Then
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), pLinePoints(pIndex), False, pPreviousDirCode)
                        Else
                            Call DrawSegment(pIndex, False, pDirectionCode, pLinePoints(pIndex - 1), pLinePoints(pIndex), True, pPreviousDirCode)
                        End If
                End Select
                pPreviousDirCode = pDirectionCode
            Next
            Call DrawCircle(False, pLinePoints(0))
        End If
        m_IGraphics.RePaintConvas
    End If
    
End Sub

Private Sub DrawCircle(ByVal pOnlyCreate As Boolean, ByRef pPoint As LINE_POINT)
    Dim pPenWidth As Long
    Dim phDC As Long
    Dim pStartColor As Long
    Dim pEndColor As Long
    Dim pRect As RECTF
    Dim rct2 As RECTF
    Dim pGradientMode As LinearGradientMode
    Dim pPoints(2) As POINTL
    Dim pRegion As Long
    Dim pPen As Long
    Dim pXShift As Single
    Dim pYShift As Single
    Dim pColour As SegColour
    Dim pGraphics As Long
    
    
    If m_Selected Then
        pColour = sSelected
    Else
        Select Case m_LayerLineType
            Case OnCompletion:
                pColour = sBlue
            Case OnFail:
                pColour = sRed
            Case Else:
                pColour = sGreen
        End Select
    End If
        
    Call GetGradColours(pColour, pStartColor, pEndColor)
    
    If pOnlyCreate Then
        pPenWidth = 1
        With pRect
            .Width = m_LineWidth * 1.4
            .Height = m_LineWidth * 1.4
            .Left = pPoint.X - .Width / 2
            .Top = pPoint.Y - .Height / 2
        End With
        
        pXShift = -pPenWidth
        pYShift = -pPenWidth
        InflateRectF pRect, pPenWidth, pPenWidth
        If m_PathCircle <> 0 Then
            Call GdipDeletePath(m_PathCircle)
            m_PathCircle = 0
        End If
        
        
        GdipCreatePath FillModeWinding, m_PathCircle
        GdipAddPathEllipse m_PathCircle, pRect.Left, pRect.Top, pRect.Width, pRect.Height
        InflateRectF pRect, pXShift, pYShift
    Else
        pGraphics = m_IGraphics.GetGraphicsHandle
        If m_BrushCircle <> 0 Then
            Call GdipDeleteBrush(m_BrushCircle)
            m_BrushCircle = 0
        End If
        GdipCreatePathGradientFromPath m_PathCircle, m_BrushCircle
        GdipSetPathGradientCenterColor m_BrushCircle, pEndColor
        GdipSetPathGradientSurroundColorsWithCount m_BrushCircle, pStartColor, 1
        GdipSetLineGammaCorrection m_BrushCircle, True
        
        If m_BrushCircle Then
            GdipFillPath pGraphics, m_BrushCircle, m_PathCircle
        End If
    
        GdipCreatePen1 pStartColor, pPenWidth, UnitPixel, pPen
        If pPen Then
            GdipDrawPath pGraphics, pPen, m_PathCircle
            GdipDeletePen pPen
        End If
        'GdipDeletePath pPath
    End If
    
    
    
End Sub

Private Sub DrawSegment(ByVal Index As Single, ByVal pOnlyCreate As Boolean, ByVal pDirectonCode As DirectionCode, ByRef pPtStart As LINE_POINT, ByRef pPtEnd As LINE_POINT, ByVal DrawJoint As Boolean, ByVal pPreviousDirCode As DirectionCode)
    Dim pPenWidth As Long
    Dim phDC As Long, pGraphics As Long
    Dim pPath As Long
    Dim pPathII As Long
    Dim pStartColor As Long
    Dim pEndColor As Long
    Dim pRect As RECTF
    Dim rct2 As RECTF
    Dim pGradientMode As LinearGradientMode
    Dim pBrush As Long
    Dim pPen As Long
    Dim pRegion As Long
    Dim pColour As SegColour
    Dim TriGradientMode As LinearGradientMode
    Dim pOrientation As SegOrientation
    Dim pTriPoints As LineBand
    Dim pblnPointDrawn As Boolean
    Dim pDebugGUI As Boolean
    Dim pPolyPoints() As POINTF
    
    If m_Selected Then
        pColour = sSelected
    Else
        Select Case m_LayerLineType
            Case OnCompletion:
                pColour = sBlue
            Case OnFail:
                pColour = sRed
            Case Else:
                pColour = sGreen
        End Select
    End If
        
    Call GetGradColours(pColour, pStartColor, pEndColor)
    
    If pDirectonCode = D2U Or pDirectonCode = U2D Then
        pOrientation = Verticle
    Else
        pOrientation = Horizontal
    End If
    
    pPenWidth = 1
    pRect = GetBoundedRectangle(pPtStart.X, pPtStart.Y, pPtEnd.X, pPtEnd.Y, m_LineWidth)
    If pOrientation = Horizontal Then
        Call InflateRectF(pRect, m_LineWidth / 2, 0)
    Else
        Call InflateRectF(pRect, 0, m_LineWidth / 2)
    End If
    LSet rct2 = pRect
    pGraphics = m_IGraphics.GetGraphicsHandle
    
    
    If pOnlyCreate Then
        If m_3DSegmentPath(Index) <> 0 Then
            Call GdipDeletePath(m_3DSegmentPath(Index))
        End If
        GdipCreatePath FillModeWinding, m_3DSegmentPath(Index)
        GdipAddPathRectangles m_3DSegmentPath(Index), pRect, 1
        If m_BrushHorizontal <> 0 Then
            Call GdipDeleteBrush(m_BrushHorizontal)
        End If
            
        If m_BrushVerticle <> 0 Then
            Call GdipDeleteBrush(m_BrushVerticle)
        End If
    Else
        With pRect
            .Left = pPtStart.X - m_LineWidth / 2
            .Top = pPtStart.Y - m_LineWidth / 2
            .Height = m_LineWidth
            .Width = m_LineWidth
        End With
        
        GdipCreatePen1 pStartColor, pPenWidth, UnitPixel, pPen
        If pOrientation = Horizontal Then
            rct2 = SetGradientRectF(pRect, LinearGradientModeVertical, 2, True)
            GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, LinearGradientModeVertical, WrapModeTileFlipX, pBrush
            GdipSetLineGammaCorrection pBrush, True
        Else
            rct2 = SetGradientRectF(pRect, LinearGradientModeHorizontal, 2, True)
            GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, LinearGradientModeHorizontal, WrapModeTileFlipX, pBrush
            GdipSetLineGammaCorrection pBrush, True
        End If
        
        If pBrush Then
        
            GdipFillPath pGraphics, pBrush, m_3DSegmentPath(Index)
            GdipDeleteBrush pBrush
        End If
            
        If pPen Then
            GdipDrawPath pGraphics, pPen, m_3DSegmentPath(Index)
        End If
        pblnPointDrawn = False
        If DrawJoint Then
            
            If pPreviousDirCode = D2U Or pPreviousDirCode = U2D Then
                TriGradientMode = LinearGradientModeHorizontal
            Else
                TriGradientMode = LinearGradientModeVertical
            End If
        
            Select Case pDirectonCode
                Case DirectionCode.D2U:
                     Select Case pPreviousDirCode
                        Case DirectionCode.D2U:
                            pblnPointDrawn = False
                        Case DirectionCode.U2D:
                            pblnPointDrawn = False
                        Case DirectionCode.L2R:
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 3, pRect)
                        Case DirectionCode.R2L:
                            
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 2, pRect)
                    End Select
                Case DirectionCode.U2D:
                    Select Case pPreviousDirCode
                        Case DirectionCode.D2U:
                            pblnPointDrawn = False
                            
                        Case DirectionCode.U2D:
                            
                            pblnPointDrawn = False
                        Case DirectionCode.L2R:
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 4, pRect)
                            
                        Case DirectionCode.R2L:
                            pblnPointDrawn = True
                            
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 1, pRect)
                    End Select
                Case DirectionCode.L2R:
                    
                    Select Case pPreviousDirCode
                        Case DirectionCode.D2U:
                            pblnPointDrawn = True
                            
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 3, pRect)
                        Case DirectionCode.U2D:
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 4, pRect)
                        Case DirectionCode.L2R:
                            pblnPointDrawn = False
                        Case DirectionCode.R2L:
                            pblnPointDrawn = False
                    End Select
                Case DirectionCode.R2L:
                    
                    Select Case pPreviousDirCode
                        Case DirectionCode.D2U:
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 2, pRect)
                        Case DirectionCode.U2D:
                            pblnPointDrawn = True
                            pTriPoints = GetTriangle(pDirectonCode, pPreviousDirCode, 1, pRect)
                        Case DirectionCode.L2R:
                            pblnPointDrawn = False
                        Case DirectionCode.R2L:
                            pblnPointDrawn = False
                    End Select
            End Select
            If pblnPointDrawn Then
                'Draw Cur Band Part
                
                If pDirectonCode = D2U Or pDirectonCode = U2D Then
                    TriGradientMode = LinearGradientModeVertical
                Else
                    TriGradientMode = LinearGradientModeHorizontal
                End If
                
                GdipCreatePath FillModeWinding, pPath
                GdipAddPathPolygon pPath, pTriPoints.BandCur(0), 4
                
                rct2 = SetGradientRectF(pRect, TriGradientMode, 2, True)
                GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, TriGradientMode, WrapModeTileFlipX, pBrush
                GdipSetLineGammaCorrection pBrush, True
                If pBrush Then
                    GdipFillPath pGraphics, pBrush, pPath
                    GdipDeleteBrush pBrush
                End If
                
            '    GdipDrawPath pGraphics, pPen, pPath
                GdipDeletePath pPath
                
                If pPreviousDirCode = D2U Or pPreviousDirCode = U2D Then
                    TriGradientMode = LinearGradientModeVertical
                Else
                    TriGradientMode = LinearGradientModeHorizontal
                End If
                
                GdipCreatePath FillModeWinding, pPath
                GdipAddPathPolygon pPath, pTriPoints.BandPre(0), 4
                rct2 = SetGradientRectF(pRect, TriGradientMode, 2, True)
                GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, TriGradientMode, WrapModeTileFlipX, pBrush
                GdipSetLineGammaCorrection pBrush, True
                If pBrush Then
                    GdipFillPath pGraphics, pBrush, pPath
                    GdipDeleteBrush pBrush
                End If
                
             '   GdipDrawPath pGraphics, pPen, pPath
                GdipDeletePath pPath


            Else
                If pDirectonCode = D2U Or pDirectonCode = U2D Then
                    TriGradientMode = LinearGradientModeHorizontal
                    Call InflateRectF(pRect, 0, 2)
                Else
                    Call InflateRectF(pRect, 2, 0)
                    TriGradientMode = LinearGradientModeVertical
                End If
                
                rct2 = SetGradientRectF(pRect, TriGradientMode, 2, True)
                GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, TriGradientMode, WrapModeTileFlipX, pBrush
                GdipSetLineGammaCorrection pBrush, True
                
                If pBrush Then
                    Call GdipFillRectangle(pGraphics, pBrush, pRect.Left, pRect.Top, pRect.Width, pRect.Height)
                    GdipDeleteBrush pBrush
                End If
            End If
    End If
    
        GdipDeletePen pPen
    End If
End Sub

'
'
'   |-------------------|
'   | 4 \ 1          / 2 |
'   |         \/        |
'   |   /       3   \   |
'   |-------------------|
'
Private Function GetTriangle(ByVal pCurrentDirCode As DirectionCode, ByVal pPreviousDirCode As DirectionCode, ByVal pTriangleIndex As Single, ByRef pRect As RECTF) As LineBand
        Dim pPoints(3) As POINTF
        Dim pBandWidth As Single
        Dim pLineBand As LineBand
        Dim pCurCord(3) As POINTF
        Dim pPreCord(3) As POINTF
        
        pBandWidth = m_LineWidth / 2
        Select Case pTriangleIndex
            Case 1:
                If pCurrentDirCode = R2L Then
                    pCurCord(0).X = pRect.Left - pBandWidth
                    pCurCord(0).Y = pRect.Top
                    
                    pCurCord(1).X = pRect.Left
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left + pRect.Width
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left - pBandWidth
                    pCurCord(3).Y = pRect.Top + pRect.Height
                    
                    pPreCord(0).X = pRect.Left
                    pPreCord(0).Y = pRect.Top - pBandWidth
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left + pRect.Width
                    pPreCord(3).Y = pRect.Top - pBandWidth
                ElseIf pCurrentDirCode = U2D Then
                    pCurCord(0).X = pRect.Left
                    pCurCord(0).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pCurCord(1).X = pRect.Left + pRect.Width
                    pCurCord(1).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pCurCord(2).X = pRect.Left + pRect.Width
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left
                    pCurCord(3).Y = pRect.Top
                    
                    pPreCord(0).X = pRect.Left + pRect.Width + pBandWidth
                    pPreCord(0).Y = pRect.Top
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left + pRect.Width + pBandWidth
                    pPreCord(3).Y = pRect.Top + pRect.Height
                Else
                    Debug.Print "Get Triangle code Error"
                End If
                pLineBand.BandCur = pPreCord
                pLineBand.BandPre = pCurCord
                    
            Case 2:
                If pCurrentDirCode = D2U Then
                    pCurCord(0).X = pRect.Left
                    pCurCord(0).Y = pRect.Top - pBandWidth
                    
                    pCurCord(1).X = pRect.Left
                    pCurCord(1).Y = pRect.Top + pRect.Height
                    
                    pCurCord(2).X = pRect.Left + pRect.Width
                    pCurCord(2).Y = pRect.Top
                    
                    pCurCord(3).X = pRect.Left + pRect.Width
                    pCurCord(3).Y = pRect.Top - pBandWidth
                    
                    pPreCord(0).X = pRect.Left + pRect.Width + pBandWidth
                    pPreCord(0).Y = pRect.Top
                    
                    pPreCord(1).X = pRect.Left + pRect.Width
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left + pRect.Width + pBandWidth
                    pPreCord(3).Y = pRect.Top + pRect.Height
                ElseIf pCurrentDirCode = R2L Then
                    pCurCord(0).X = pRect.Left - pBandWidth
                    pCurCord(0).Y = pRect.Top
                    
                    pCurCord(1).X = pRect.Left + pRect.Width
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left - pBandWidth
                    pCurCord(3).Y = pRect.Top + pRect.Height
                    
                    pPreCord(0).X = pRect.Left
                    pPreCord(0).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top + pRect.Height
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top
                    
                    pPreCord(3).X = pRect.Left + pRect.Width
                    pPreCord(3).Y = pRect.Top + pRect.Height + pBandWidth
                Else
                    Debug.Print "Get Triangle code Error"
                End If
                
                pLineBand.BandCur = pPreCord
                pLineBand.BandPre = pCurCord
            Case 3:
                If pCurrentDirCode = D2U Then
                    pCurCord(0).X = pRect.Left
                    pCurCord(0).Y = pRect.Top - pBandWidth
                    
                    pCurCord(1).X = pRect.Left
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left + pRect.Width
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left + pRect.Width
                    pCurCord(3).Y = pRect.Top - pBandWidth
                    
                    pPreCord(0).X = pRect.Left - pBandWidth
                    pPreCord(0).Y = pRect.Top
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left - pBandWidth
                    pPreCord(3).Y = pRect.Top + pRect.Height
                ElseIf pCurrentDirCode = L2R Then
                    pCurCord(0).X = pRect.Left + pRect.Width + pBandWidth
                    pCurCord(0).Y = pRect.Top
                    
                    pCurCord(1).X = pRect.Left
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left + pRect.Width
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left + pRect.Width + pBandWidth
                    pCurCord(3).Y = pRect.Top + pRect.Height
                    
                    pPreCord(0).X = pRect.Left
                    pPreCord(0).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left + pRect.Width
                    pPreCord(3).Y = pRect.Top + pRect.Height + pBandWidth
                Else
                    Debug.Print "Get Triangle code Error"
                End If
                
                pLineBand.BandCur = pPreCord
                pLineBand.BandPre = pCurCord
            Case 4:
                If pCurrentDirCode = U2D Then
                    pCurCord(0).X = pRect.Left + pRect.Width
                    pCurCord(0).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pCurCord(1).X = pRect.Left + pRect.Width
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left
                    pCurCord(3).Y = pRect.Top + pRect.Height + pBandWidth
                    
                    pPreCord(0).X = pRect.Left - pBandWidth
                    pPreCord(0).Y = pRect.Top
                    
                    pPreCord(1).X = pRect.Left + pRect.Width
                    pPreCord(1).Y = pRect.Top
                    
                    pPreCord(2).X = pRect.Left
                    pPreCord(2).Y = pRect.Top + pRect.Height
                    
                    pPreCord(3).X = pRect.Left - pBandWidth
                    pPreCord(3).Y = pRect.Top + pRect.Height
                ElseIf pCurrentDirCode = L2R Then
                    pCurCord(0).X = pRect.Left + pRect.Width + pBandWidth
                    pCurCord(0).Y = pRect.Top
                    
                    pCurCord(1).X = pRect.Left + pRect.Width
                    pCurCord(1).Y = pRect.Top
                    
                    pCurCord(2).X = pRect.Left
                    pCurCord(2).Y = pRect.Top + pRect.Height
                    
                    pCurCord(3).X = pRect.Left + pRect.Width + pBandWidth
                    pCurCord(3).Y = pRect.Top + pRect.Height
                    
                    pPreCord(0).X = pRect.Left
                    pPreCord(0).Y = pRect.Top - pBandWidth
                    
                    pPreCord(1).X = pRect.Left
                    pPreCord(1).Y = pRect.Top + pRect.Height
                    
                    pPreCord(2).X = pRect.Left + pRect.Width
                    pPreCord(2).Y = pRect.Top
                    
                    pPreCord(3).X = pRect.Left + pRect.Width
                    pPreCord(3).Y = pRect.Top - pBandWidth
                Else
                    Debug.Print "Get Triangle code Error"
                End If
                
                pLineBand.BandCur = pPreCord
                pLineBand.BandPre = pCurCord
        End Select
        
        
        GetTriangle = pLineBand
End Function


Private Sub DrawArrow(ByVal pOnlyCreate As Boolean, ByVal pDirectonCode As DirectionCode, ByRef pPoint As LINE_POINT)
    Dim pPenWidth As Long
    Dim phDC As Long, pGraphics As Long
    Dim pStartColor As Long
    Dim pEndColor As Long
    Dim pRect As RECTF
    Dim rct2 As RECTF
    Dim pGradientMode As LinearGradientMode
    Dim pPoints(2) As POINTF
    Dim pRegion As Long
    Dim pPen As Long
    Dim pXShift As Single
    Dim pYShift As Single
    Dim pColour As SegColour
    
    
    If m_Selected Then
        pColour = sSelected
    Else
        Select Case m_LayerLineType
            Case OnCompletion:
                pColour = sBlue
            Case OnFail:
                pColour = sRed
            Case Else:
                pColour = sGreen
        End Select
    End If
    
    Call GetGradColours(pColour, pStartColor, pEndColor)
    If pDirectonCode = D2U Or pDirectonCode = U2D Then
        pGradientMode = LinearGradientModeHorizontal
    Else
        pGradientMode = LinearGradientModeVertical
    End If
    
    
    pPenWidth = 1
    
    pGraphics = m_IGraphics.GetGraphicsHandle
    pOnlyCreate = True
    If pOnlyCreate Then
         If m_PathArrow <> 0 Then
            Call GdipDeletePath(m_PathArrow)
            m_PathArrow = 0
         End If
         GdipCreatePath FillModeWinding, m_PathArrow
    End If
  '  If pPoint.x = 0 Then Stop
    Select Case pDirectonCode
        Case DirectionCode.L2R:
            With pRect
                .Width = m_LineWidth * 2
                .Height = m_LineWidth * 2
                .Left = pPoint.X - .Width
                .Top = pPoint.Y - .Height / 2
            End With
            pPoints(0).X = pRect.Left
            pPoints(0).Y = pRect.Top
            pPoints(1).X = pRect.Left + pRect.Width
            pPoints(1).Y = pRect.Top + pRect.Height / 2
            pPoints(2).X = pRect.Left
            pPoints(2).Y = pRect.Top + pRect.Height
            If pOnlyCreate Then GdipAddPathPolygon m_PathArrow, pPoints(0), 3
        Case DirectionCode.R2L:
            With pRect
                .Width = m_LineWidth * 2
                .Height = m_LineWidth * 2
                .Left = pPoint.X
                .Top = pPoint.Y - .Height / 2
            End With
            pPoints(0).X = pRect.Left
            pPoints(0).Y = pRect.Top + pRect.Height / 2
            pPoints(1).X = pRect.Left + pRect.Width
            pPoints(1).Y = pRect.Top
            pPoints(2).X = pRect.Left + pRect.Width
            pPoints(2).Y = pRect.Top + pRect.Height
            If pOnlyCreate Then GdipAddPathPolygon m_PathArrow, pPoints(0), 3
            
        Case DirectionCode.U2D:
            With pRect
                .Width = m_LineWidth * 2
                .Height = m_LineWidth * 2
                .Left = pPoint.X - .Width / 2
                .Top = pPoint.Y - .Height
            End With
            pPoints(0).X = pRect.Left
            pPoints(0).Y = pRect.Top
            pPoints(1).X = pRect.Left + pRect.Width
            pPoints(1).Y = pRect.Top
            pPoints(2).X = pRect.Left + pRect.Width / 2
            pPoints(2).Y = pRect.Top + pRect.Height
            If pOnlyCreate Then GdipAddPathPolygon m_PathArrow, pPoints(0), 3
        Case DirectionCode.D2U:
            With pRect
                .Width = m_LineWidth * 2
                .Height = m_LineWidth * 2
                .Left = pPoint.X - .Width / 2
                .Top = pPoint.Y
            End With
            pPoints(0).X = pRect.Left + pRect.Width / 2
            pPoints(0).Y = pRect.Top
            pPoints(1).X = pRect.Left + pRect.Width
            pPoints(1).Y = pRect.Top + pRect.Height
            pPoints(2).X = pRect.Left
            pPoints(2).Y = pRect.Top + pRect.Height
            If pOnlyCreate Then GdipAddPathPolygon m_PathArrow, pPoints(0), 3
    End Select
    
        
        If m_BrushArrow <> 0 Then
            Call GdipDeletePath(m_BrushArrow)
            m_BrushArrow = 0
         End If
         
        rct2 = SetGradientRectF(pRect, pGradientMode, 2, True)
        GdipCreateLineBrushFromRect rct2, pStartColor, pEndColor, pGradientMode, WrapModeTileFlipX, m_BrushArrow
        GdipSetLineGammaCorrection m_BrushArrow, True
         
        If m_BrushArrow Then
             GdipFillPath pGraphics, m_BrushArrow, m_PathArrow
             GdipDeleteBrush m_BrushArrow
         End If
        
        GdipCreatePen1 pStartColor, pPenWidth, UnitPixel, pPen
        If pPen Then
            GdipDrawPath pGraphics, pPen, m_PathArrow
            GdipDeletePen pPen
        End If
          
        'GdipDeletePath m_PathArrow
    
End Sub

Private Sub m_ConnectedTo_MouseDown(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim pPoints() As LINE_POINT
    If Button = vbLeftButton Then
        
        If Shift = vbCtrlMask And m_Selected Then
            m_MovingLinePoints = False
            Exit Sub
        End If
        m_AllLinesForThreeStateDeleted = False
        m_XShift = m_ConnectedTo.CentreX - X
        m_YShift = m_ConnectedTo.CentreY - Y
        If m_PointerArrowDrawn Then
            Let m_RemovedPoint = m_Points(m_PointCount - 1)
            m_PointRemoved = True
            ReDim Preserve m_Points(m_PointCount - 2)
            Call CalculatePointCount
            m_PointerArrowDrawn = False
        End If
        m_MovingPointIndex = m_PointCount - 2
        m_MovingLinePoints = True
    End If
End Sub

Private Sub m_ConnectedTo_MouseMove(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim pCenX As Double
    Dim pCenY As Double
    Dim pSide As Single
    Dim pX As Double, pY As Double
    Dim pXAct As Double
    Dim pYAct As Double
    If Shift = vbCtrlMask And m_Selected Then
       m_MovingLinePoints = False
       Exit Sub
    End If
    
    If m_MovingLinePoints And Button = vbLeftButton Then
        If m_TempToUbound > 0 Then
            pX = m_TempToPx1(m_TempToUbound - 1)
            pY = m_TempToPy1(m_TempToUbound - 1)
        Else
            If m_PointCount = 2 Then
                pX = m_Points(m_PointCount - 2).X
                pY = m_Points(m_PointCount - 2).Y
            Else
                pX = m_Points(m_PointCount - 2).X
                pY = m_Points(m_PointCount - 2).Y
            End If
            
        End If
        pSide = GetAppropriateSide(False, pX, pY)
        pX = m_Points(0).X
        pY = m_Points(0).Y
        Call m_ConnectedTo.GetExactCordinates(m_NodeKey, pSide, m_LineWidth, m_AllocatedSideKeyTo, pX, pY, pXAct, pYAct)
     
        m_NeedToReversePoints = False
        Call DrawTempLinesFromToPoint(pX / 15, pY / 15, pXAct / 15, pYAct / 15)
        DoEvents
        m_LastToMouseX = X
        m_LastToMouseY = Y
    End If
    
    
End Sub

Private Sub DrawTempLinesFromToPoint(ByVal X As Double, ByVal Y As Double, ByVal XAct As Double, ByVal YAct As Double)
    Dim pIndex As Single
    Dim pLineIndex As Single
    Dim px1  As Double, py1  As Double, px2  As Double, py2 As Double
    Dim pFound As Boolean
    Dim pSegOrient As Single '1 Verticle 0 hori
    Dim pSegmentTolook As Single
    Dim pSegOrientlook As Single
    Dim pNewIndex As Single
    Dim pSegmentsDrawn As Boolean
    Dim pYCordToCheck As Double
    Dim pYCurrent As Double
    Dim pIsU2D As Boolean
    Dim pIsD2U As Boolean
    Dim pIsL2R As Boolean
    Dim pIsR2L As Boolean
    
    pSegmentTolook = m_PointCount - 2
    
    If m_Points(pSegmentTolook).X = m_Points(pSegmentTolook + 1).X Then
        pSegOrientlook = 1
        If m_Points(pSegmentTolook).Y > m_Points(pSegmentTolook + 1).Y Then
            pIsU2D = False
            pIsD2U = True
            pYCordToCheck = Y
            pYCurrent = m_Points(pSegmentTolook).Y
        Else
            pIsU2D = True
            pIsD2U = False
            pYCordToCheck = m_Points(pSegmentTolook).Y
            pYCurrent = Y
        End If
    Else
        pSegOrientlook = 0
        If m_Points(pSegmentTolook).X > m_Points(pSegmentTolook + 1).X Then
            pIsL2R = False
            pIsR2L = True
            pYCordToCheck = X
            pYCurrent = m_Points(pSegmentTolook).X
        Else
            pIsL2R = True
            pIsR2L = False
            pYCordToCheck = m_Points(pSegmentTolook).X
            pYCurrent = X
        End If
    End If
    
    If (m_PointCount - 2) = 0 Then
        Call ThreeSegmentLine(X, Y, False, XAct, YAct)
        Exit Sub
    End If
    
    For pIndex = 0 To m_PointCount - 2
        px1 = m_Points(pIndex).X
        py1 = m_Points(pIndex).Y
        px2 = m_Points(pIndex + 1).X
        py2 = m_Points(pIndex + 1).Y
        If px1 = px2 Then
            pSegOrient = 1
        Else
            pSegOrient = 0
        End If
        If pSegOrientlook = 1 Then
                If pYCurrent >= pYCordToCheck Then
                    If Not pSegmentsDrawn Then
                        Call DrawLineTo(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    Select Case pIndex
                        Case pSegmentTolook - 1
                            Call DrawLineTo(pLineIndex, px1, py1, X, py2, True)
                        Case pSegmentTolook
                                Call DrawLineTo(pLineIndex, X, py1, X, Y, True)
                        Case Else
                            Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                Else
                    If Not pSegmentsDrawn Then
                        Call DrawLineTo(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    If pIsU2D Then
                        If pSegOrient = 0 Then
                            If py2 >= Y Then
                                Call DrawLineTo(pLineIndex, px1, py1, X, py1, True)
                                Call DrawLineTo(pLineIndex, X, py1, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If px2 >= X Then
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineTo(pLineIndex, px2, py2, X, py2, True)
                                Call DrawLineTo(pLineIndex, X, py2, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    ElseIf pIsD2U Then
                        If pSegOrient = 0 Then
                            If py2 <= Y Then
                                Call DrawLineTo(pLineIndex, px1, py1, X, py1, True)
                                Call DrawLineTo(pLineIndex, X, py1, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If px2 <= X Then
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineTo(pLineIndex, px2, py2, X, py2, True)
                                Call DrawLineTo(pLineIndex, X, py2, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    End If
                End If
        Else    'Horizontal
            If pYCurrent >= pYCordToCheck Then
                    If Not pSegmentsDrawn Then
                        Call DrawLineTo(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    Select Case pIndex
                        Case pSegmentTolook - 1
                            Call DrawLineTo(pLineIndex, px1, py1, px2, Y, True)
                        Case pSegmentTolook
                            Call DrawLineTo(pLineIndex, px1, Y, X, Y, True)
                        Case Else
                            Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                Else
                    If Not pSegmentsDrawn Then
                        Call DrawLineTo(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    If pIsL2R Then
                        If pSegOrient = 0 Then
                            If px2 >= X Then
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineTo(pLineIndex, px2, py2, px2, Y, True)
                                Call DrawLineTo(pLineIndex, px2, Y, X, Y, True)
                                Exit For
                                
                            Else
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If py2 >= Y Then
                                Call DrawLineTo(pLineIndex, px1, py1, px1, Y, True)
                                Call DrawLineTo(pLineIndex, px1, Y, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    ElseIf pIsR2L Then
                        If pSegOrient = 0 Then
                            If px2 <= X Then
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineTo(pLineIndex, px2, py2, px2, Y, True)
                                Call DrawLineTo(pLineIndex, px2, Y, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If py2 <= Y Then
                                Call DrawLineTo(pLineIndex, px1, py1, px1, Y, True)
                                Call DrawLineTo(pLineIndex, px1, Y, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineTo(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    End If
                End If
            
        End If
    Next
    
    Call DrawFinalArrow(X, Y, XAct, YAct)
End Sub


Private Function DrawFinalArrow(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double) As LINE_POINT
    Static pArrowPoint(1) As LINE_POINT
    Dim pPoint As LINE_POINT
    If X1 = X2 And Y1 = Y2 And X1 = 0 And Y1 = 0 Then
        m_ConvasPicBox.Line (pArrowPoint(0).X * 15, pArrowPoint(0).Y * 15)-(pArrowPoint(1).X * 15, pArrowPoint(1).Y * 15)
        Let pPoint = pArrowPoint(1)
        DrawFinalArrow = pPoint
        pArrowPoint(0).X = 0
        pArrowPoint(0).Y = 0
        pArrowPoint(1).X = 0
        pArrowPoint(1).Y = 0
    Else
        If pArrowPoint(0).X = pArrowPoint(0).Y And pArrowPoint(1).X = pArrowPoint(1).Y And pArrowPoint(0).X = 0 And pArrowPoint(1).X = 0 Then
            pArrowPoint(0).X = X1
            pArrowPoint(0).Y = Y1
            pArrowPoint(1).X = X2
            pArrowPoint(1).Y = Y2
            m_ConvasPicBox.Line (pArrowPoint(0).X * 15, pArrowPoint(0).Y * 15)-(pArrowPoint(1).X * 15, pArrowPoint(1).Y * 15)
        Else
            m_ConvasPicBox.Line (pArrowPoint(0).X * 15, pArrowPoint(0).Y * 15)-(pArrowPoint(1).X * 15, pArrowPoint(1).Y * 15)
            pArrowPoint(0).X = X1
            pArrowPoint(0).Y = Y1
            pArrowPoint(1).X = X2
            pArrowPoint(1).Y = Y2
            m_ConvasPicBox.Line (pArrowPoint(0).X * 15, pArrowPoint(0).Y * 15)-(pArrowPoint(1).X * 15, pArrowPoint(1).Y * 15)
        End If
    End If
End Function


Private Sub DrawTempLinesFromFromPoint(ByVal X As Double, ByVal Y As Double)
    Dim pIndex As Single
    Dim pLineIndex As Single
    Dim px1  As Double, py1  As Double, px2  As Double, py2 As Double
    Dim pFound As Boolean
    Dim pSegOrient As Single '1 Verticle 0 hori
    Dim pSegmentTolook As Single
    Dim pSegOrientlook As Single
    Dim pNewIndex As Single
    Dim pSegmentsDrawn As Boolean
    Dim pYCordToCheck As Double
    Dim pYCurrent As Double
    Dim pIsU2D As Boolean
    Dim pIsD2U As Boolean
    Dim pIsL2R As Boolean
    Dim pIsR2L As Boolean
    
    pSegmentTolook = 0
    
   
    'If m_Points(pSegmentTolook).X = m_Points(pSegmentTolook + 1).X Then
    If m_Points(pSegmentTolook).X = m_Points(pSegmentTolook + 1).X Then
        pSegOrientlook = 1
        'If m_Points(pSegmentTolook).Y > m_Points(pSegmentTolook + 1).Y Then
        If m_Points(pSegmentTolook).Y > m_Points(pSegmentTolook + 1).Y Then
            pIsU2D = False
            pIsD2U = True
            pYCordToCheck = Y
            pYCurrent = m_Points(pSegmentTolook + 1).Y
        Else
            pIsU2D = True
            pIsD2U = False
            pYCordToCheck = m_Points(pSegmentTolook + 1).Y
            pYCurrent = Y
        End If
    Else
        pSegOrientlook = 0
        If m_Points(pSegmentTolook).X > m_Points(pSegmentTolook + 1).X Then
            pIsL2R = False
            pIsR2L = True
            pYCordToCheck = X
            pYCurrent = m_Points(pSegmentTolook + 1).X
        Else
            pIsL2R = True
            pIsR2L = False
            pYCordToCheck = m_Points(pSegmentTolook + 1).X
            pYCurrent = X
        End If
    End If
    
    If (m_PointCount - 2) = 0 Then
        Call ThreeSegmentLine(X, Y, True)
        Exit Sub
    End If
    
    For pIndex = m_PointCount - 1 To 1 Step -1
        px1 = m_Points(pIndex).X
        py1 = m_Points(pIndex).Y
        px2 = m_Points(pIndex - 1).X
        py2 = m_Points(pIndex - 1).Y
        If pIndex = 1 Then
            'px1 = m_XStart
            'py1 = m_YStart
        End If
        If px1 = px2 Then
            pSegOrient = 1
        Else
            pSegOrient = 0
        End If
        If pSegOrientlook = 1 Then
                If pYCurrent >= pYCordToCheck Then
                    If Not pSegmentsDrawn Then
                        Call DrawLineFrom(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    Select Case pIndex
                        Case pSegmentTolook + 2
                            Call DrawLineFrom(pLineIndex, px1, py1, X, py2, True)
                        Case pSegmentTolook + 1
                            Call DrawLineFrom(pLineIndex, X, py1, X, Y, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                Else
                    If Not pSegmentsDrawn Then
                        Call DrawLineFrom(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    If pIsU2D Then
                        If pSegOrient = 0 Then
                            If py2 >= Y Then
                                Call DrawLineFrom(pLineIndex, px1, py1, X, py1, True)
                                Call DrawLineFrom(pLineIndex, X, py1, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If px2 >= X Then
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineFrom(pLineIndex, px2, py2, X, py2, True)
                                Call DrawLineFrom(pLineIndex, X, py2, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    ElseIf pIsD2U Then
                        If pSegOrient = 0 Then
                            If py2 <= Y Then
                            
                                Call DrawLineFrom(pLineIndex, px1, py1, X, py1, True)
                                Call DrawLineFrom(pLineIndex, X, py1, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If px2 <= X Then
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            
                                Call DrawLineFrom(pLineIndex, px2, py2, X, py2, True)
                                Call DrawLineFrom(pLineIndex, X, py2, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    End If
                End If
        Else    'Horizontal
            If pYCurrent >= pYCordToCheck Then
                    If Not pSegmentsDrawn Then
                        Call DrawLineFrom(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    Select Case pIndex
                        Case pSegmentTolook + 2
                            
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, Y, True)
                        Case pSegmentTolook + 1
                            Call DrawLineFrom(pLineIndex, px1, Y, X, Y, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                Else
                    If Not pSegmentsDrawn Then
                        Call DrawLineFrom(-1, 0, 0, 0, 0, True)
                        pSegmentsDrawn = True
                    End If
                    If pIsL2R Then
                        If pSegOrient = 0 Then
                            If px2 >= X Then
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                                
                                Call DrawLineFrom(pLineIndex, px2, py2, px2, Y, True)
                                Call DrawLineFrom(pLineIndex, px2, Y, X, Y, True)
                                Exit For
                                
                            Else
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If py2 >= Y Then
                                
                                Call DrawLineFrom(pLineIndex, px1, py1, px1, Y, True)
                                Call DrawLineFrom(pLineIndex, px1, Y, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    ElseIf pIsR2L Then
                        If pSegOrient = 0 Then
                            If px2 <= X Then
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                                Call DrawLineFrom(pLineIndex, px2, py2, px2, Y, True)
                                Call DrawLineFrom(pLineIndex, px2, Y, X, Y, True)
                                Exit For
                            Else
                                Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        Else
                            If py2 <= Y Then
                                Call DrawLineFrom(pLineIndex, px1, py1, px1, Y, True)
                                Call DrawLineFrom(pLineIndex, px1, Y, X, Y, True)
                                Exit For
                            Else
                               Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                            End If
                        End If
                    End If
                End If
            
        End If
    Next
    
End Sub


Private Sub m_ConnectedTo_MouseUp(ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    Dim pPoints() As LINE_POINT
    Dim pArrowPoint As LINE_POINT
    Dim pPointCount As Single
    

    If m_MovingLinePoints Then
        pPoints = DrawLineTo(0, 0, 0, 0, 0)
        pArrowPoint = DrawFinalArrow(0, 0, 0, 0)
        If pArrowPoint.X = pArrowPoint.Y And pArrowPoint.Y = 0 Then
            pPointCount = UBound(pPoints)
            ReDim Preserve pPoints(pPointCount + 1)
            Let pPoints(pPointCount + 1) = m_RemovedPoint
        Else
            pPointCount = UBound(pPoints)
            ReDim Preserve pPoints(pPointCount + 1)
            Let pPoints(pPointCount + 1) = pArrowPoint
        End If
        m_PointerArrowDrawn = True
        m_PointRemoved = False
        Me.SetPoints pPoints
        Me.Paint
    End If
    m_MovingLinePoints = False
End Sub

Private Sub m_ConnectedTo_ReAlign(ByVal ShiftX As Double, ByVal ShiftY As Double)
    Dim pIndex As Single
    If m_Selected Then
     
        If Not m_RealinedOnce Then
            Debug.Print "Align to " & m_ConnectedTo.NodeKey & " : " & m_NodeKey
            If Not (ShiftX = ShiftY And ShiftX = 0) Then
                For pIndex = 0 To m_PointCount - 1
                    m_Points(pIndex).X = m_Points(pIndex).X + ShiftX / 15
                    m_Points(pIndex).Y = m_Points(pIndex).Y + ShiftY / 15
                Next
                Call CalculatePointCount
                Call CreatePath
                m_RealinedOnce = True
            End If
        Else
            m_RealinedOnce = False
        End If
    End If
End Sub

Private Sub m_ConnectedTo_Selected(ByVal Key As String)
    If m_ConnectedTo.oSelected And m_ConnectedFrom.oSelected Then
        Me.oSelected = True
     Else
        If m_Selected Then m_Selected = False
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
    If m_Visible Then
        If m_IGraphics.IsConvasLocked Then Exit Sub
        If m_Region <> 0 And Not m_MovingLinePoints Then
            Call GdipIsVisibleRegionPoint(m_Region, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)
            If pResult = 1 Or (m_IGraphics.IsSelected(m_NodeKey) And Shift = vbCtrlMask) Then
                
                If Shift <> vbCtrlMask Then
                    m_MovingPointIndex = GetPointIndex(X / 15, Y / 15)
                    m_blnPointMoving = True
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
                'If m_Selected Then
                    
                    If Shift <> vbCtrlMask Then
                        m_RaiseFromInside = True
                        Me.oSelected = False
                    End If
                'End If
            End If
        End If
    End If
End Sub

Private Function GetPointIndex(ByVal X As Double, ByVal Y As Double) As Single
    Dim pIndex As Single
    Dim pRectF As RECTF
    Dim pRect As RECT
    Dim pResult As Long
    Dim pPointIndex As Single
    GetPointIndex = -1
    For pIndex = 0 To m_PointCount - 2
        pRectF = GetBoundedRectangle(m_Points(pIndex).X, m_Points(pIndex).Y, m_Points(pIndex + 1).X, m_Points(pIndex + 1).Y, m_LineWidth)
        With pRect
            .Left = pRectF.Left
            .Top = pRectF.Top
            .Right = pRectF.Left + pRectF.Width
            .Bottom = pRectF.Top + pRectF.Height
        End With

        pResult = PtInRect(pRect, X, Y)
        If pResult Then
            pPointIndex = pIndex
            Exit For
        End If
    Next
    GetPointIndex = pPointIndex
End Function


Friend Sub CreatePath()
    Dim pIndex As Single
    Dim pRectF As RECTF
    Dim pRegion As Long
    Dim pDirectionCode As DirectionCode
    Dim pPointCount As Single
    Dim pLinePoints() As LINE_POINT
    Dim ptPoint As LINE_POINT
    
    If m_Region <> 0 Then
        GdipDeleteRegion m_Region
    End If
    
    
    
    If Not m_Use3D Then
        If m_Path <> 0 Then
            GdipDeletePath m_Path
            m_Path = 0
        End If
        Call GdipCreatePath(FillModeWinding, m_Path)
    Else
        If m_PointCount > 0 Then
            ReDim pLinePoints(m_PointCount - 1)
        Else
            ReDim pLinePoints(m_PointCount)
        End If
        For pIndex = 0 To m_PointCount - 1
            LSet pLinePoints(pIndex) = m_Points(pIndex)
        Next
        
        If m_PointRemoved Then
            ReDim Preserve pLinePoints(m_PointCount)
            LSet pLinePoints(m_PointCount) = m_RemovedPoint
        End If
        pPointCount = UBound(pLinePoints)
    End If
    
    For pIndex = 0 To m_PointCount - 2
        
        If Not m_Use3D Then Call GdipAddPathLine(m_Path, m_Points(pIndex).X, m_Points(pIndex).Y, m_Points(pIndex + 1).X, m_Points(pIndex + 1).Y)
        pRectF = GetBoundedRectangle(m_Points(pIndex).X, m_Points(pIndex).Y, m_Points(pIndex + 1).X, m_Points(pIndex + 1).Y, m_LineWidth)
        If pIndex = 0 Then
             Call GdipCreateRegionRect(pRectF, m_Region)
        Else
            Call GdipCombineRegionRect(m_Region, pRectF, CombineModeUnion)
        End If
    Next
    
    If m_Use3D Then
        For pIndex = 0 To pPointCount '- 1
                If pIndex > 0 Then
                    If pLinePoints(pIndex).X = pLinePoints(pIndex - 1).X Then 'Verticle
                        If pLinePoints(pIndex).Y > pLinePoints(pIndex - 1).Y Then
                            pDirectionCode = U2D
                        Else
                            pDirectionCode = D2U
                        End If
                    Else    'Horizontal
                        If pLinePoints(pIndex).X > pLinePoints(pIndex - 1).X Then
                            pDirectionCode = L2R
                        Else
                            pDirectionCode = R2L
                        End If
                    End If
                End If
                Select Case pIndex:
                    Case 0:
                    Case pPointCount:
                        If pDirectionCode = L2R Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.X = ptPoint.X - m_LineWidth * 2.5
                            Call DrawSegment(pIndex, True, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, False, 0)
                        ElseIf pDirectionCode = R2L Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.X = ptPoint.X + m_LineWidth * 2.5
                            Call DrawSegment(pIndex, True, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, False, 0)
                        ElseIf pDirectionCode = U2D Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.Y = ptPoint.Y - m_LineWidth * 2.5
                            Call DrawSegment(pIndex, True, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, False, 0)
                        ElseIf pDirectionCode = D2U Then
                            LSet ptPoint = pLinePoints(pIndex)
                            ptPoint.Y = ptPoint.Y + m_LineWidth * 2.5
                            Call DrawSegment(pIndex, True, pDirectionCode, pLinePoints(pIndex - 1), ptPoint, False, 0)
                        End If
                        Call DrawArrow(True, pDirectionCode, pLinePoints(pIndex))
                    Case Else
                        Call DrawSegment(pIndex, True, pDirectionCode, pLinePoints(pIndex - 1), pLinePoints(pIndex), False, 0)
                        If pIndex = 1 Then
                            Call DrawCircle(True, pLinePoints(0))
                        End If
                End Select
        Next

    End If
    
End Sub


Private Sub CreatePen(ByVal pPenWidth As Single)
    Dim pBrush As Long
    Dim pColour As Colors
    Dim pBackColour As Colors
    Dim pWidth As Double
    Dim pHatchStyle As HatchStyle
    
    pHatchStyle = HatchStyle50Percent
    Select Case m_LayerLineType
        Case LineType.OnCompletion:
            pColour = Black
            pBackColour = Blue
        Case LineType.OnFail:
            pColour = Black
            pBackColour = Red
        Case LineType.OnSuccess:
            pColour = Black
            pBackColour = Green
    End Select
    If m_Selected Then
        pPenWidth = pPenWidth '+ 1
        pColour = Colors.LightGray
        pHatchStyle = HatchStyle50Percent
        
    End If
    Call GdipCreateHatchBrush(pHatchStyle, pColour, pBackColour, pBrush)
    Call GdipCreatePen2(pBrush, pPenWidth, UnitPixel, m_Pen)
    Call GdipSetPenStartCap(m_Pen, LineCapRoundAnchor)
    Call GdipSetPenEndCap(m_Pen, LineCapArrowAnchor)
    Call GdipSetPenLineJoin(m_Pen, LineJoinRound)
    Call GdipSetPenBrushFill(m_Pen, pBrush)
    Call GdipDeleteBrush(pBrush)
End Sub


Private Sub CreateDeletePath()
    GdipDeletePath m_Path
    m_Path = 0
End Sub

Public Property Get Points() As LINE_POINT()
     Points = m_Points
End Property

Public Sub SetPoints(ByRef pPoints() As LINE_POINT)
    m_Points = pPoints
    Call CalculatePointCount
    ReDim Preserve m_3DSegmentPath(m_PointCount)
    Call CreatePath
End Sub

Friend Property Get PointCount() As Single
    PointCount = m_PointCount
End Property


Private Sub CalculatePointCount()
    On Error GoTo ErrorTrap
    m_PointCount = UBound(m_Points) + 1

    Exit Sub
ErrorTrap:
    m_PointCount = 0
End Sub

Private Function GetBoundedRectangle(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal LineWidth As Single) As RECTF
    Dim pRect As RECTF
    If Abs(X2 - X1) > Abs(Y2 - Y1) Then 'Horizontal
        With pRect
            .Left = IIf(X2 > X1, X1, X2)
            .Top = IIf(Y2 > Y1, Y1, Y2) - LineWidth / 2
            .Width = Abs(X2 - X1)
            .Height = LineWidth
        End With
    Else    'Verticle
        With pRect
            .Left = IIf(X2 > X1, X1, X2) - LineWidth / 2
            .Top = IIf(Y2 > Y1, Y1, Y2)
            .Width = LineWidth
            .Height = Abs(Y2 - Y1)
        End With
    End If
    GetBoundedRectangle = pRect
End Function


Private Function GetPointer(ByVal X As Double, ByVal Y As Double) As MousePointerConstants
    Dim pIndex As Single
    Dim pRectF As RECTF
    Dim pRect As RECT
    Dim pResult As Long
    GetPointer = vbDefault
    For pIndex = 0 To m_PointCount - 2
        pRectF = GetBoundedRectangle(m_Points(pIndex).X, m_Points(pIndex).Y, m_Points(pIndex + 1).X, m_Points(pIndex + 1).Y, m_LineWidth)
        With pRect
            .Left = pRectF.Left
            .Top = pRectF.Top
            .Right = pRectF.Left + pRectF.Width
            .Bottom = pRectF.Top + pRectF.Height
        End With

        pResult = PtInRect(pRect, X, Y)
        If pResult Then
            If pRectF.Height = m_LineWidth Then
                GetPointer = vbSizeNS
            Else
                GetPointer = vbSizeWE
            End If
            Exit For
        End If
    Next
End Function


Private Sub m_ConvasPicBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pResult  As Long
    If m_Visible Then
        
        If m_Region <> 0 And Not m_blnPointMoving And Not m_MovingLinePoints And Button = 0 Then
            Call GdipIsVisibleRegionPoint(m_Region, X / 15, Y / 15, m_IGraphics.GetGraphicsHandle, pResult)
            If pResult = 1 Then
               m_ConvasPicBox.Tag = m_NodeKey
               If Shift <> vbCtrlMask Then
                    m_ConvasPicBox.ToolTipText = m_ToolTipText
                    m_ConvasPicBox.MousePointer = GetPointer(X / 15, Y / 15)
               Else
                    m_ConvasPicBox.ToolTipText = ""
                    m_ConvasPicBox.MousePointer = MousePointerConstants.vbCustom
               End If
            Else
               If Shift <> vbCtrlMask Then
                     If m_ConvasPicBox.Tag = m_NodeKey Then
                        m_ConvasPicBox.Tag = ""
                        m_ConvasPicBox.ToolTipText = ""
                        m_ConvasPicBox.MousePointer = vbDefault
                    End If
               Else
                    m_ConvasPicBox.ToolTipText = ""
                    m_ConvasPicBox.MousePointer = MousePointerConstants.vbCustom
               End If
            End If
        ElseIf m_blnPointMoving And Not m_MovingLinePoints And Button = vbLeftButton Then
           ' If pLastX <> X And pLastY <> Y Then
           
           Call DrawTempLine(X / 15, Y / 15)
        End If
    End If
End Sub


Private Function DrawLineTo(ByRef pLineIndex As Single, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, Optional ByVal ErasePresious As Boolean = True, Optional ByVal LastSegment As Boolean) As LINE_POINT()

        Dim pPoints() As LINE_POINT
        Dim pIndex As Single
        Dim pSearchX As Double
        Dim pSearchY As Double
        Dim pFound As Boolean

        If X1 = X2 And Y1 = Y2 And X1 = Y1 And X1 = 0 Then
                
            If pLineIndex = -1 And m_TempToUbound > 0 Then
                For pIndex = 0 To m_TempToUbound
                    m_ConvasPicBox.Line (m_TempToPx1(pIndex) * 15, m_TempToPy1(pIndex) * 15)-(m_TempToPx2(pIndex) * 15, m_TempToPy2(pIndex) * 15)
                Next
                Erase m_TempToPx1
                Erase m_TempToPy1
                Erase m_TempToPx2
                Erase m_TempToPy2
                m_TempToUbound = 0
                Exit Function
            End If
                            
            If LastSegment Then
                For pIndex = pLineIndex To m_TempToUbound
                    m_ConvasPicBox.Line (m_TempToPx1(pLineIndex) * 15, m_TempToPy1(pLineIndex) * 15)-(m_TempToPx2(pLineIndex) * 15, m_TempToPy2(pLineIndex) * 15)
                Next
                ReDim Preserve m_TempToPx1(pLineIndex + 1)
                ReDim Preserve m_TempToPy1(pLineIndex + 1)
                ReDim Preserve m_TempToPx2(pLineIndex + 1)
                ReDim Preserve m_TempToPy2(pLineIndex + 1)
                m_TempToUbound = UBound(m_TempToPx1)
                Exit Function
            End If
                            
            ReDim pPoints(m_TempToUbound + 1)
            pIndex = 0
            pLineIndex = 0
            pFound = False
            m_ConvasPicBox.DrawMode = DrawModeConstants.vbCopyPen
            If m_TempToUbound > 0 Then
                For pLineIndex = 0 To m_TempToUbound
                        If pLineIndex > 0 Then
                            If pPoints(pLineIndex - 1).X = m_TempToPx1(pLineIndex) And pPoints(pLineIndex - 1).Y = m_TempToPy1(pLineIndex) Then
                                If pIndex > 1 Then
                                    pIndex = pIndex - 1
                                End If
                            Else

                                pPoints(pIndex).X = m_TempToPx1(pLineIndex)
                                pPoints(pIndex).Y = m_TempToPy1(pLineIndex)
                                pIndex = pIndex + 1
                            End If
                        Else
                                pPoints(pIndex).X = m_TempToPx1(pLineIndex)
                                pPoints(pIndex).Y = m_TempToPy1(pLineIndex)
                                pIndex = pIndex + 1
                        End If
                    m_ConvasPicBox.Line (m_TempToPx1(pLineIndex) * 15, m_TempToPy1(pLineIndex) * 15)-(m_TempToPx2(pLineIndex) * 15, m_TempToPy2(pLineIndex) * 15)
                Next
                
                ReDim Preserve pPoints(pIndex)
                pPoints(pIndex).X = m_TempToPx2(pLineIndex - 1)
                pPoints(pIndex).Y = m_TempToPy2(pLineIndex - 1)
            Else
                pPoints = m_Points

            End If
            Erase m_TempToPx1
            Erase m_TempToPy1
            Erase m_TempToPx2
            Erase m_TempToPy2
            m_TempToUbound = 0
            DrawLineTo = pPoints
            Exit Function
        End If

        If m_TempToUbound <= pLineIndex Then
            ReDim Preserve m_TempToPx1(pLineIndex)
            ReDim Preserve m_TempToPy1(pLineIndex)
            ReDim Preserve m_TempToPx2(pLineIndex)
            ReDim Preserve m_TempToPy2(pLineIndex)
            m_TempToUbound = UBound(m_TempToPx1)
        End If

        m_ConvasPicBox.DrawMode = DrawModeConstants.vbInvert
        If ErasePresious Then m_ConvasPicBox.Line (m_TempToPx1(pLineIndex) * 15, m_TempToPy1(pLineIndex) * 15)-(m_TempToPx2(pLineIndex) * 15, m_TempToPy2(pLineIndex) * 15)
        m_TempToPx1(pLineIndex) = X1
        m_TempToPy1(pLineIndex) = Y1
        m_TempToPx2(pLineIndex) = X2
        m_TempToPy2(pLineIndex) = Y2
        m_ConvasPicBox.Line (X1 * 15, Y1 * 15)-(X2 * 15, Y2 * 15)
        pLineIndex = pLineIndex + 1
        
End Function

Private Function DrawLineFrom(ByRef pLineIndex As Single, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, Optional ByVal ErasePresious As Boolean = True, Optional ByVal LastSegment As Boolean) As LINE_POINT()

        Dim pPoints() As LINE_POINT
        Dim pIndex As Single
        Dim pSearchX As Double
        Dim pSearchY As Double
        Dim pFound As Boolean
        
        If X1 = X2 And Y1 = Y2 And X1 = Y1 And X1 = 0 Then
                
            If pLineIndex = -1 And m_TempFromUbound > 0 Then
                For pIndex = 0 To m_TempFromUbound
                    m_ConvasPicBox.Line (m_TempFromPx1(pIndex) * 15, m_TempFromPy1(pIndex) * 15)-(m_TempFromPx2(pIndex) * 15, m_TempFromPy2(pIndex) * 15)
                Next
                Erase m_TempFromPx1
                Erase m_TempFromPy1
                Erase m_TempFromPx2
                Erase m_TempFromPy2
                m_TempFromUbound = 0
                Exit Function
            End If
                            
            If LastSegment Then
                For pIndex = pLineIndex To m_TempFromUbound
                    m_ConvasPicBox.Line (m_TempFromPx1(pLineIndex) * 15, m_TempFromPy1(pLineIndex) * 15)-(m_TempFromPx2(pLineIndex) * 15, m_TempFromPy2(pLineIndex) * 15)
                Next
                ReDim Preserve m_TempFromPx1(pLineIndex + 1)
                ReDim Preserve m_TempFromPy1(pLineIndex + 1)
                ReDim Preserve m_TempFromPx2(pLineIndex + 1)
                ReDim Preserve m_TempFromPy2(pLineIndex + 1)
                m_TempFromUbound = UBound(m_TempFromPx1)
                Exit Function
            End If
                            
            ReDim pPoints(m_TempFromUbound + 1)
            pIndex = 0
            pLineIndex = 0
            pFound = False
            m_ConvasPicBox.DrawMode = DrawModeConstants.vbCopyPen
            If m_TempFromUbound > 0 Then
                For pLineIndex = 0 To m_TempFromUbound
                        If pLineIndex > 0 Then
                            If pPoints(pLineIndex - 1).X = m_TempFromPx1(pLineIndex) And pPoints(pLineIndex - 1).Y = m_TempFromPy1(pLineIndex) Then
                                If pIndex > 1 Then
                                    pIndex = pIndex - 1
                                End If
                            Else

                                pPoints(pIndex).X = m_TempFromPx1(pLineIndex)
                                pPoints(pIndex).Y = m_TempFromPy1(pLineIndex)
                                pIndex = pIndex + 1
                            End If
                        Else
                                pPoints(pIndex).X = m_TempFromPx1(pLineIndex)
                                pPoints(pIndex).Y = m_TempFromPy1(pLineIndex)
                                pIndex = pIndex + 1
                        End If
                    m_ConvasPicBox.Line (m_TempFromPx1(pLineIndex) * 15, m_TempFromPy1(pLineIndex) * 15)-(m_TempFromPx2(pLineIndex) * 15, m_TempFromPy2(pLineIndex) * 15)
                Next
                
                ReDim Preserve pPoints(pIndex)
                pPoints(pIndex).X = m_TempFromPx2(pLineIndex - 1)
                pPoints(pIndex).Y = m_TempFromPy2(pLineIndex - 1)
            Else
                pPoints = m_Points

            End If
            Erase m_TempFromPx1
            Erase m_TempFromPy1
            Erase m_TempFromPx2
            Erase m_TempFromPy2
            m_TempFromUbound = 0
            DrawLineFrom = pPoints
            Exit Function
        End If

        If m_TempFromUbound <= pLineIndex Then
            ReDim Preserve m_TempFromPx1(pLineIndex)
            ReDim Preserve m_TempFromPy1(pLineIndex)
            ReDim Preserve m_TempFromPx2(pLineIndex)
            ReDim Preserve m_TempFromPy2(pLineIndex)
            m_TempFromUbound = UBound(m_TempFromPx1)
        End If

        m_ConvasPicBox.DrawMode = DrawModeConstants.vbInvert
        If ErasePresious Then m_ConvasPicBox.Line (m_TempFromPx1(pLineIndex) * 15, m_TempFromPy1(pLineIndex) * 15)-(m_TempFromPx2(pLineIndex) * 15, m_TempFromPy2(pLineIndex) * 15)
        m_TempFromPx1(pLineIndex) = X1
        m_TempFromPy1(pLineIndex) = Y1
        m_TempFromPx2(pLineIndex) = X2
        m_TempFromPy2(pLineIndex) = Y2
        m_ConvasPicBox.Line (X1 * 15, Y1 * 15)-(X2 * 15, Y2 * 15)
        pLineIndex = pLineIndex + 1
        
End Function

Private Sub DrawTempLine(ByVal X As Double, ByVal Y As Double)
    Dim pIndex As Single
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double
    Dim pOrient As Single '0 is ori 1 is ver
    Dim pLineIndex As Single
    If m_MovingPointIndex < m_PointCount Then
        pOrient = IIf(Abs(m_Points(m_MovingPointIndex).X - m_Points(m_MovingPointIndex + 1).X) = 0, 1, 0)
    ElseIf m_MovingPointIndex > 0 Then
        pOrient = IIf(Abs(m_Points(m_MovingPointIndex - 1).X - m_Points(m_MovingPointIndex).X) = 0, 1, 0)
    Else
        ' single Line
        'Stop
    End If
    pLineIndex = 0
    For pIndex = 0 To m_PointCount - 2
        px1 = m_Points(pIndex).X
        py1 = m_Points(pIndex).Y
        px2 = m_Points(pIndex + 1).X
        py2 = m_Points(pIndex + 1).Y


            If pOrient = 1 Then ' ver
                If m_MovingPointIndex > 0 And m_MovingPointIndex < m_PointCount - 2 Then
                    Select Case pIndex
                        Case m_MovingPointIndex - 1:
                            Call DrawLineFrom(pLineIndex, px1, py1, X, py2, True)
                        Case m_MovingPointIndex:
                            Call DrawLineFrom(pLineIndex, X, py1, X, py2, True)
                        Case m_MovingPointIndex + 1:
                            Call DrawLineFrom(pLineIndex, X, py1, px2, py2, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                ElseIf m_MovingPointIndex = 0 Then
                    Select Case pIndex
                        Case 0:
                            Call DrawLineFrom(pLineIndex, px1, py1, X, py1, True)
                            Call DrawLineFrom(pLineIndex, X, py1, X, py2, True)
                            If m_PointCount = 2 Then
                                Call DrawLineFrom(pLineIndex, X, py2, px2, py2, True)
                            End If
                        Case 1:
                            Call DrawLineFrom(pLineIndex, X, py2, px2, py2, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                ElseIf m_MovingPointIndex = m_PointCount - 2 Then
                    Select Case pIndex
                        Case m_MovingPointIndex:
                            Call DrawLineFrom(pLineIndex, X, py1, X, py2, True)
                            Call DrawLineFrom(pLineIndex, X, py2, px2, py2, True)
                        Case m_MovingPointIndex - 1:
                            Call DrawLineFrom(pLineIndex, px1, py1, X, py2, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                End If

            Else    'hori
                If m_MovingPointIndex > 0 And m_MovingPointIndex < m_PointCount - 2 Then
                    Select Case pIndex
                        Case m_MovingPointIndex - 1:
                            Call DrawLineFrom(pLineIndex, px1, py1, px1, Y, True)
                        Case m_MovingPointIndex:
                            Call DrawLineFrom(pLineIndex, px1, Y, px2, Y, True)
                        Case m_MovingPointIndex + 1:
                            Call DrawLineFrom(pLineIndex, px2, Y, px2, py2, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                ElseIf m_MovingPointIndex = 0 Then
                    Select Case pIndex
                        Case 0:
                            Call DrawLineFrom(pLineIndex, px1, py1, px1, Y, True)
                            Call DrawLineFrom(pLineIndex, px1, Y, px2, Y, True)
                            If m_PointCount = 2 Then
                                Call DrawLineFrom(pLineIndex, px2, Y, px2, py2, True)
                            End If
                        Case 1:
                            Call DrawLineFrom(pLineIndex, px1, Y, px2, py2, True)
                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                ElseIf m_MovingPointIndex = m_PointCount - 2 Then
                    Select Case pIndex
                        Case m_MovingPointIndex: 'm_PointCount - 2:
                            Call DrawLineFrom(pLineIndex, px1, Y, px2, Y, True)
                            Call DrawLineFrom(pLineIndex, px2, Y, px2, py2, True)
                        Case m_MovingPointIndex - 1 'm_PointCount - 3:
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, Y, True)

                        Case Else
                            Call DrawLineFrom(pLineIndex, px1, py1, px2, py2, True)
                    End Select
                End If
            End If
    Next


End Sub


Private Sub ThreeSegmentLine(ByVal X As Double, ByVal Y As Double, Optional FromPoint As Boolean = False, Optional ByVal XAct As Double, Optional ByVal YAct As Double)
    Dim pPoints() As LINE_POINT
    Dim pX As Double
    Dim pY As Double
    Dim pIndex As Single
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double
    Dim pUbound As Single
    Dim pLineIndex As Single
        
        
        If FromPoint Then
            m_NeedToReversePoints = False
            px1 = X
            py1 = Y
            px2 = m_Points(m_PointCount - 1).X
            py2 = m_Points(m_PointCount - 1).Y
        Else
            m_NeedToReversePoints = True
            px1 = m_Points(0).X
            py1 = m_Points(0).Y
            px2 = X
            py2 = Y
        End If
        
        ReDim Preserve pPoints(pIndex)
        pPoints(pIndex).X = px1 / 15
        pPoints(pIndex).Y = py1 / 15
        pIndex = pIndex + 1
        If Abs(px1 - px2) = 0 Or Abs(py1 - py2) = 0 Then

        ElseIf Abs(px1 - px2) > Abs(py1 - py2) Then
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = ((px1 + px2) / 2) / 15
            pPoints(pIndex).Y = py1 / 15
            pIndex = pIndex + 1
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = ((px1 + px2) / 2) / 15
            pPoints(pIndex).Y = py2 / 15
            pIndex = pIndex + 1

        Else
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = px1 / 15
            pPoints(pIndex).Y = ((py1 + py2) / 2) / 15
            pIndex = pIndex + 1
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = px2 / 15
            pPoints(pIndex).Y = ((py1 + py2) / 2) / 15
            pIndex = pIndex + 1

        End If
        ReDim Preserve pPoints(pIndex)
        pPoints(pIndex).X = px2 / 15
        pPoints(pIndex).Y = py2 / 15
            
        pUbound = UBound(pPoints)
        pLineIndex = 0
        For pIndex = 0 To pUbound - 1
            If FromPoint Then
                Call DrawLineFrom(pLineIndex, pPoints(pIndex).X * 15, pPoints(pIndex).Y * 15, pPoints(pIndex + 1).X * 15, pPoints(pIndex + 1).Y * 15, True)
            Else
                Call DrawLineTo(pLineIndex, pPoints(pIndex).X * 15, pPoints(pIndex).Y * 15, pPoints(pIndex + 1).X * 15, pPoints(pIndex + 1).Y * 15, True)
            End If
        Next
        
        If Not FromPoint Then
            Call DrawFinalArrow(pPoints(pIndex).X * 15, pPoints(pIndex).Y * 15, XAct, YAct)
        End If
End Sub



Private Sub m_ConvasPicBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pPoints() As LINE_POINT
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
            Let pPoints = DrawLineFrom(0, 0, 0, 0, 0)
            Me.SetPoints pPoints
            Me.Paint
            m_blnPointMoving = False
        End If
    End If
End Sub




Private Sub StartLine()
    Dim pPoints() As LINE_POINT
    Dim pX As Double
    Dim pY As Double
    Dim pIndex As Single
    Dim px1 As Double
    Dim py1 As Double
    Dim px2 As Double
    Dim py2 As Double

        If m_X1 = -1 Or m_Y1 = -1 Or m_X2 = -1 Or m_Y2 = -1 Then
            Exit Sub
        End If

        px1 = m_X1
        py1 = m_Y1
        px2 = m_X2
        py2 = m_Y2
        m_X1 = -1
        m_Y1 = -1
        m_X2 = -1
        m_Y2 = -1

        ReDim Preserve pPoints(pIndex)
        pPoints(pIndex).X = px1 / 15
        pPoints(pIndex).Y = py1 / 15
        pIndex = pIndex + 1
        If Abs(px1 - px2) = 0 Or Abs(py1 - py2) = 0 Then

        ElseIf Abs(px1 - px2) > Abs(py1 - py2) Then
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = ((px1 + px2) / 2) / 15
            pPoints(pIndex).Y = py1 / 15
            pIndex = pIndex + 1
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = ((px1 + px2) / 2) / 15
            pPoints(pIndex).Y = py2 / 15
            pIndex = pIndex + 1

        Else
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = px1 / 15
            pPoints(pIndex).Y = ((py1 + py2) / 2) / 15
            pIndex = pIndex + 1
            ReDim Preserve pPoints(pIndex)
            pPoints(pIndex).X = px2 / 15
            pPoints(pIndex).Y = ((py1 + py2) / 2) / 15
            pIndex = pIndex + 1

        End If
        ReDim Preserve pPoints(pIndex)
        pPoints(pIndex).X = px2 / 15
        pPoints(pIndex).Y = py2 / 15
        Call Me.SetPoints(pPoints)
        
        Call m_ConnectedFrom_MouseDown(m_ConnectedFrom.NodeKey, vbLeftButton, 0, m_ConnectedFrom.CentreX, m_ConnectedFrom.CentreY)
        Call m_ConnectedFrom_MouseMove(m_ConnectedFrom.NodeKey, vbLeftButton, 0, m_ConnectedFrom.CentreX + 1, m_ConnectedFrom.CentreY + 1)
        Call m_ConnectedFrom_MouseMove(m_ConnectedFrom.NodeKey, vbLeftButton, 0, m_ConnectedFrom.CentreX - 1, m_ConnectedFrom.CentreY - 1)
        Call m_ConnectedFrom_MouseUp(m_ConnectedFrom.NodeKey, vbLeftButton, 0, m_ConnectedFrom.CentreX, m_ConnectedFrom.CentreY)
       
        Call m_ConnectedTo_MouseDown(m_ConnectedTo.NodeKey, vbLeftButton, 0, m_ConnectedTo.CentreX, m_ConnectedTo.CentreY)
        Call m_ConnectedTo_MouseMove(m_ConnectedTo.NodeKey, vbLeftButton, 0, m_ConnectedTo.CentreX + 1, m_ConnectedTo.CentreY + 1)
        Call m_ConnectedTo_MouseMove(m_ConnectedTo.NodeKey, vbLeftButton, 0, m_ConnectedTo.CentreX - 1, m_ConnectedTo.CentreY - 1)
        Call m_ConnectedTo_MouseUp(m_ConnectedTo.NodeKey, vbLeftButton, 0, m_ConnectedTo.CentreX, m_ConnectedTo.CentreY)
        
        Me.Paint
End Sub


Public Sub ReSizeConvas()
    Call InitilizeEnviornment
    Call CreatePath
End Sub

'--------------------------Control Handling---------------------------------

Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Print UserControl.Name
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
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim pIndex As Single
    Call InitilizeEnviornment
    m_X1 = -1
    m_Y1 = -1
    m_X2 = -1
    m_Y2 = -1
    m_LineWidth = 5
    m_Use3D = PropBag.ReadProperty("Use3D", False)
    m_NodeKey = PropBag.ReadProperty("NodeKey", "")
    m_FromKey = PropBag.ReadProperty("FromKey", "")
    m_ToKey = PropBag.ReadProperty("ToKey", "")
    m_ConstraintIndex = PropBag.ReadProperty("ConstraintIndex", 0)
    m_LayerLineType = PropBag.ReadProperty("LayerLineType", LineType.OnSuccess)
    m_StepName = PropBag.ReadProperty("StepName", "")
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Visible = PropBag.ReadProperty("Visible", False)
    m_Selected = PropBag.ReadProperty("oSelected", False)
    m_PointCount = PropBag.ReadProperty("PointCount", 0)
    If m_PointCount > 0 Then
        ReDim m_Points(m_PointCount - 1)
        For pIndex = 0 To m_PointCount - 1
            m_Points(pIndex).X = PropBag.ReadProperty("Px" & pIndex, 0)
            m_Points(pIndex).Y = PropBag.ReadProperty("Py" & pIndex, 0)
        Next
    End If
    'm_Use3D = True
End Sub

Public Property Get FromKey() As String
    FromKey = m_FromKey
End Property

Public Property Get ToKey() As String
    ToKey = m_ToKey
End Property
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
    If m_Pen <> 0 Then
        Call GdipDeletePen(m_Pen)
        m_Pen = 0
    End If
    
    If m_Region <> 0 Then
        Call GdipDeleteRegion(m_Region)
        m_Region = 0
    End If
    
    
    If m_PathArrow Then
        Call GdipDeletePath(m_PathArrow)
        m_PathArrow = 0
    End If
    
    If m_PathCircle Then
        Call GdipDeletePath(m_PathCircle)
        m_PathCircle = 0
    End If
    
    If m_BrushArrow Then
        Call GdipDeleteBrush(m_BrushArrow)
        m_BrushArrow = 0
    End If
    
    If m_BrushCircle Then
        Call GdipDeleteBrush(m_BrushCircle)
        m_BrushCircle = 0
    End If
    
    Set m_IGraphics = Nothing
    Set m_ConvasPicBox = Nothing
 End Sub
'--------------------------Control Handling End---------------------------------

'--------------------------Standard Functions-----------------------------------
Friend Sub GetCordinates(ByRef CenX As Double, ByRef CenY As Double, ByRef oWidth As Double, ByRef oHight As Double, Optional ByVal ResetValues As Boolean)
    CenX = ((m_Points(0).X + m_Points(m_PointCount - 1).X) / 2) * 15
    CenY = ((m_Points(0).Y + m_Points(m_PointCount - 1).Y) / 2) * 15
    oWidth = Abs(m_Points(0).X - m_Points(m_PointCount - 1).X) * 15
    oHight = Abs(m_Points(0).Y - m_Points(m_PointCount - 1).Y) * 15
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim pIndex As Single
    Call PropBag.WriteProperty("NodeKey", m_NodeKey, "")
    Call PropBag.WriteProperty("FromKey", m_FromKey, Nothing)
    Call PropBag.WriteProperty("ToKey", m_ToKey, Nothing)
    Call PropBag.WriteProperty("ConstraintIndex", m_ConstraintIndex, -1)
    Call PropBag.WriteProperty("LayerLineType", m_LayerLineType, LineType.OnSuccess)
    Call PropBag.WriteProperty("StepName", m_StepName, "")
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, "")
    Call PropBag.WriteProperty("Visible", m_Visible, False)
    Call PropBag.WriteProperty("oSelected", m_Selected, False)
    Call PropBag.WriteProperty("PointCount", m_PointCount, 0)
    Call PropBag.WriteProperty("Use3D", m_Use3D, False)
    If m_PointCount > 0 Then
        For pIndex = 0 To m_PointCount - 1
            Call PropBag.WriteProperty("Px" & pIndex, m_Points(pIndex).X, 0)
            Call PropBag.WriteProperty("Py" & pIndex, m_Points(pIndex).Y, 0)
        Next
    End If
    'Save Points and Point Count
    
End Sub
