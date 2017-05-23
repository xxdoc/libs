VERSION 5.00
Begin VB.UserControl oConvas 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   7050
   ToolboxBitmap   =   "Convas.ctx":0000
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox picoConvas 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   10000
         Left            =   0
         MouseIcon       =   "Convas.ctx":0312
         MousePointer    =   1  'Arrow
         ScaleHeight     =   10005
         ScaleWidth      =   10005
         TabIndex        =   4
         Top             =   0
         Width           =   10000
      End
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   -480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar BarVs 
      Enabled         =   0   'False
      Height          =   2415
      LargeChange     =   10
      Left            =   3240
      Max             =   100
      SmallChange     =   5
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar BarHs 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      SmallChange     =   5
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Timer SelectionTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   840
      Top             =   2520
   End
   Begin VB.Timer RealignTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   180
      Top             =   2550
   End
   Begin VB.Timer PaintTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3900
      Top             =   2190
   End
   Begin oGraph.oLine LLine 
      Index           =   0
      Left            =   4980
      Top             =   3030
      _ExtentX        =   873
      _ExtentY        =   873
      FromKey         =   ""
      ToKey           =   ""
      ConstraintIndex =   0
      Use3D           =   -1  'True
   End
   Begin oGraph.oPicture LPicture 
      Index           =   0
      Left            =   3420
      Top             =   3270
      _ExtentX        =   873
      _ExtentY        =   873
      LSCount         =   0
   End
   Begin oGraph.oText LText 
      Index           =   0
      Left            =   2400
      Top             =   3450
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "oConvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IXGraphics
Private WithEvents m_picoConvas As PictureBox
Attribute m_picoConvas.VB_VarHelpID = -1
Private m_oGraphics As Long
Private m_SelectedTextKeys As Collection
Private m_SelectedPicKeys  As Collection
Private m_SelectedLineKeys As Collection
Private m_ObjectCollections As Collection
Private m_LineIndex As Single
Private m_PictureIndex As Single
Private m_TextIndex As Single
Private m_LockLayerEdit As Boolean
Private m_X1MouseDown As Double
Private m_Y1MouseDown As Double
Private m_Drawing As Boolean
Private m_X1Old As Double
Private m_X2Old As Double
Private m_Y1Old As Double
Private m_Y2Old As Double
Private m_MemoryDC As cDibSection 'cMemDC

Private Declare Function GetCursorPos Lib "USER32.DLL" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "USER32.DLL" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private m_MemorycDIB As cDibSection

Public Event DoubleClick(ByVal ObjectType As eObjectType, ByVal ObjectKey As String)
Public Event RightClick(ByVal ObjectType As eObjectType, ByVal ObjectKey As String, ByVal X As Double, ByVal Y As Double)
Public Event MouseDown(ByVal ObjectType As eObjectType, ByVal ObjectKey As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event MouseMove(ByVal ObjectType As eObjectType, ByVal ObjectKey As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event MouseUp(ByVal ObjectType As eObjectType, ByVal ObjectKey As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
Public Event Click(ByVal ObjectType As eObjectType, ByVal ObjectKey As String)
Public Event Deleted(ByVal ObjectType As eObjectType, ByVal ObjectKey As String)

Private m_AlreadyPainting As Boolean
Private m_IsSelected As Boolean
Private m_SelectionByRect As Boolean
Private m_LeftMargin As Long, m_MinLeftMargin As Long, m_RightMargin As Long, m_MinRightMargin As Long, m_TopMargin As Long, m_MinTopMargin As Long, m_BottomMargin As Long, m_MinBottomMargin As Long
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_LoadingFromBinary As Boolean
Private m_ClickCount As Single
Private m_Zoom As Single
Private m_OrgMatrix As Long
Private m_MaxHeight As Single
Private m_MaxWidth As Single
Private m_MaxExtendByType As eObjectType
Private m_MaxExtendIndex As Single
Private m_ReadyForPaint As Boolean
Private Sub BarHs_Change()
    Call lft
End Sub

Private Sub BarHs_Scroll()
    Call lft
End Sub

Private Sub BarVs_Change()
    Call tP
End Sub

Private Sub BarVs_Scroll()
    Call tP
End Sub

Private Sub tP()
   Dim l As Double
   Dim A As Double
   Dim X As Double
   
   X = BarVs.value
   A = picoConvas.Height - picControl.Height
   l = (A * X) / 100
   picoConvas.Top = -l

End Sub

Private Sub lft()
   Dim l As Double
   Dim A As Double
   Dim X As Double
   
   X = BarHs.value
   A = picoConvas.Width - picControl.Width
   l = (A * X) / 100
   picoConvas.Left = -l

End Sub

Private Function IXGraphics_GetoConvasPicBox() As Object
    If m_picoConvas Is Nothing Then Set m_picoConvas = picoConvas
    Set IXGraphics_GetoConvasPicBox = m_picoConvas
End Function

Private Function IXGraphics_GetoGraphicsHandle() As Long
    IXGraphics_GetoGraphicsHandle = m_oGraphics
End Function

Private Function IXGraphics_IsoConvasLocked() As Boolean
    IXGraphics_IsoConvasLocked = m_LockLayerEdit
End Function

Private Function IXGraphics_IsSelected(ByVal NodeKey As String) As Boolean
    On Error GoTo Errortrap
    Dim pObject As ControlWrapper
    Set pObject = m_ObjectCollections.Item(NodeKey)
    IXGraphics_IsSelected = pObject.ControlObject.oSelected
    Exit Function
Errortrap:
    IXGraphics_IsSelected = False
End Function

Private Sub IXGraphics_RePaintoConvas()
    If Not m_AlreadyPainting Then PaintTimer.Enabled = True
End Sub

Private Sub LLine_Click(Index As Integer, ByVal Key As String)
    On Error Resume Next
    If (LLine(Index).oSelected) Then
        m_SelectedLineKeys.Add Key, Key
    Else
        m_SelectedLineKeys.Remove Key
    End If
    RaiseEvent Click(oTLine, Key)
End Sub

Private Sub LLine_Deleted(Index As Integer, ByVal Key As String)
    RaiseEvent Deleted(oTLine, Key)
End Sub

Private Sub LLine_DoubleClick(Index As Integer, ByVal Key As String)
    RaiseEvent DoubleClick(oTLine, Key)
End Sub

Private Sub LLine_MaxExtend(Index As Integer, ByVal Width As Single, ByVal Height As Single)
    Dim pChanged As Boolean
    If m_MaxExtendByType = oTLine And m_MaxExtendIndex = Index Then
        m_MaxHeight = Height
        pChanged = True
    Else
        If m_MaxHeight < Height Then
            m_MaxHeight = Height
            pChanged = True
        End If
    End If
    
    If m_MaxExtendByType = oTLine And m_MaxExtendIndex = Index Then
        m_MaxWidth = Width
        pChanged = True
    Else
        If m_MaxWidth < Width Then
            m_MaxWidth = Width
            pChanged = True
        End If
    End If
    
    If pChanged Then
        Call MaxExtendChanged
    End If
    m_MaxExtendByType = oTLine
    m_MaxExtendIndex = Index
End Sub

Private Sub LLine_MouseDown(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseDown(oTLine, Key, Button, Shift, X, Y)
End Sub

Private Sub LLine_MouseMove(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseMove(oTLine, Key, Button, Shift, X, Y)
End Sub

Private Sub LLine_MouseUp(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseUp(oTLine, Key, Button, Shift, X, Y)
End Sub

Private Sub LLine_RightClick(Index As Integer, ByVal Key As String, ByVal X As Double, ByVal Y As Double)
   On Error Resume Next
    If (LLine(Index).oSelected) Then
        m_SelectedLineKeys.Add Key, Key
    Else
        m_SelectedLineKeys.Remove Key
    End If
    RaiseEvent RightClick(oTLine, Key, X, Y)
End Sub

Private Sub LLine_Selected(Index As Integer, ByVal Key As String)
    Call SelectoGraphObject(oTLine, Index, Key)
    m_ClickCount = m_ClickCount + 1
    SelectionTimer.Enabled = True
End Sub

Private Sub SelectoGraphObject(ByVal pObjectType As eObjectType, ByVal pIndex As Single, ByVal pKey As String)
    Dim pObject As ControlWrapper
    On Error Resume Next
    Set pObject = m_ObjectCollections.Item(pKey)
    If pObject.ControlObject.oSelected Then
        m_ObjectCollections.Remove (pKey)
        Call m_ObjectCollections.Add(pObject, pKey)
    End If
    Set pObject = Nothing
    
    Select Case pObjectType
        Case eObjectType.oTLine:
            If m_SelectedLineKeys Is Nothing Then Set m_SelectedLineKeys = New Collection
            m_SelectedLineKeys.Add pKey, pKey
        Case eObjectType.oTPicture:
            If m_SelectedPicKeys Is Nothing Then Set m_SelectedPicKeys = New Collection
            m_SelectedPicKeys.Add pKey, pKey
        Case eObjectType.oTText:
            If m_SelectedTextKeys Is Nothing Then Set m_SelectedTextKeys = New Collection
            m_SelectedTextKeys.Add pKey, pKey
    End Select
    Set m_picoConvas = picoConvas
End Sub


Private Sub LLine_UnSelectAll(Index As Integer, ByVal pExceptNodeKey As String)
    If IsSelected Then Call UnSelectAllObjects
End Sub

Private Sub LPicture_Click(Index As Integer, ByVal Key As String)
    On Error Resume Next
    If (LPicture(Index).oSelected) Then
        m_SelectedPicKeys.Add Key, Key
    Else
        m_SelectedPicKeys.Remove Key
    End If
    RaiseEvent Click(oTPicture, Key)
End Sub

Private Sub LPicture_Deleted(Index As Integer, ByVal Key As String)
    RaiseEvent Deleted(oTPicture, Key)
End Sub

Private Sub LPicture_DoubleClick(Index As Integer, ByVal Key As String)
    RaiseEvent DoubleClick(oTPicture, Key)
End Sub

Private Sub LPicture_MaxExtend(Index As Integer, ByVal Width As Single, ByVal Height As Single)
    Dim pChanged As Boolean
       
    
    If m_MaxExtendByType = oTPicture And m_MaxExtendIndex = Index Then
        m_MaxHeight = Height
        pChanged = True
    Else
        If m_MaxHeight < Height Then
            m_MaxHeight = Height
            pChanged = True
        End If
    End If
    
    If m_MaxExtendByType = oTPicture And m_MaxExtendIndex = Index Then
        m_MaxWidth = Width
        pChanged = True
    Else
        If m_MaxWidth < Width Then
            m_MaxWidth = Width
            pChanged = True
        End If
    End If
    
    If pChanged Then
        Call MaxExtendChanged
    End If
    m_MaxExtendByType = oTPicture
    m_MaxExtendIndex = Index
End Sub

Private Sub LPicture_MouseDown(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseDown(oTPicture, Key, Button, Shift, X, Y)
End Sub

Private Sub LPicture_MouseMove(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseMove(oTPicture, Key, Button, Shift, X, Y)
End Sub

Private Sub LPicture_MouseUp(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseUp(oTPicture, Key, Button, Shift, X, Y)
End Sub

Private Sub LPicture_RightClick(Index As Integer, ByVal Key As String, ByVal X As Double, ByVal Y As Double)
    On Error Resume Next
    If (LPicture(Index).oSelected) Then
        m_SelectedPicKeys.Add Key, Key
    Else
        m_SelectedPicKeys.Remove Key
    End If
    RaiseEvent RightClick(oTPicture, Key, X, Y)
End Sub

Private Sub LPicture_Selected(Index As Integer, ByVal Key As String)
    Call SelectoGraphObject(oTPicture, Index, Key)
    m_ClickCount = m_ClickCount + 1
    SelectionTimer.Enabled = True
End Sub

Private Sub LPicture_UnSelectAll(Index As Integer, ByVal pExceptNodeKey As String)
    If IsSelected Then Call UnSelectAllObjects
End Sub

Private Sub LText_Click(Index As Integer, ByVal Key As String)
    On Error Resume Next
    If (LText(Index).oSelected) Then
        m_SelectedTextKeys.Add Key, Key
    Else
        m_SelectedTextKeys.Remove Key
    End If
    RaiseEvent Click(oTText, Key)
End Sub

Private Sub LText_Deleted(Index As Integer, ByVal Key As String)
    RaiseEvent Deleted(oTText, Key)
End Sub

Private Sub LText_DoubleClick(Index As Integer, ByVal Key As String)
     RaiseEvent DoubleClick(oTText, Key)
     LText(Index).oSelected = False
     LText(Index).Paint
     Call ShowPropertyPages(LText(Index).object, "Text", UserControl.hWnd)
     LText(Index).oSelected = True
     LText(Index).Activate
     LText(Index).Paint
     
End Sub

Private Sub ShowPropertyPages(comObject As Object, name As String, hWnd As Long)
    Dim specifyPages As ISpecifyPropertyPages
    Set specifyPages = comObject
    If Not specifyPages Is Nothing Then
        Dim pages As CAUUID
        specifyPages.GetPages pages
        OleCreatePropertyFrame hWnd, 0, 0, name, 1, comObject, pages.cElems, ByVal pages.pElems, 0, 0, 0
        CoTaskMemFree pages.pElems
        
    End If
End Sub

Private Sub LText_MaxExtend(Index As Integer, ByVal Width As Single, ByVal Height As Single)
    Dim pChanged As Boolean
    If m_MaxExtendByType = oTText And m_MaxExtendIndex = Index Then
        m_MaxHeight = Height
        pChanged = True
    Else
        If m_MaxHeight < Height Then
            m_MaxHeight = Height
            pChanged = True
        End If
    End If
    
    If m_MaxExtendByType = oTText And m_MaxExtendIndex = Index Then
        m_MaxWidth = Width
        pChanged = True
    Else
        If m_MaxWidth < Width Then
            m_MaxWidth = Width
            pChanged = True
        End If
    End If
    
    If pChanged Then
        Call MaxExtendChanged
    End If
    m_MaxExtendByType = oTText
    m_MaxExtendIndex = Index
End Sub

Private Sub LText_MouseDown(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseDown(oTText, Key, Button, Shift, X, Y)
End Sub

Private Sub LText_MouseMove(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseMove(oTText, Key, Button, Shift, X, Y)
End Sub

Private Sub LText_MouseUp(Index As Integer, ByVal Key As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
    RaiseEvent MouseUp(oTText, Key, Button, Shift, X, Y)
End Sub

Private Sub LText_RightClick(Index As Integer, ByVal Key As String, ByVal X As Double, ByVal Y As Double)
   On Error Resume Next
    If (LText(Index).oSelected) Then
        m_SelectedTextKeys.Add Key, Key
    Else
        m_SelectedTextKeys.Remove Key
    End If
    RaiseEvent RightClick(oTText, Key, X, Y)
End Sub

Private Sub LText_Selected(Index As Integer, ByVal Key As String)
    Call SelectoGraphObject(oTText, Index, Key)
    m_ClickCount = m_ClickCount + 1
    SelectionTimer.Enabled = True
End Sub

Private Sub LText_UnSelectAll(Index As Integer, ByVal pExceptNodeKey As String)
    If IsSelected Then Call UnSelectAllObjects
End Sub


Public Sub ShowTextProperties(ByVal TextKey As String)
    Dim pControlWrap As ControlWrapper
    Set pControlWrap = m_ObjectCollections.Item(TextKey)
    
    Call ShowPropertyPages(LText(pControlWrap.Index).object, "Text Properties", UserControl.hWnd)
    
End Sub


Public Sub GetCursorPosition(ByRef X As Single, ByRef Y As Single)
    Dim pCurPoint As POINTAPI
    Call GetCursorPos(pCurPoint)
    Call ScreenToClient(UserControl.hWnd, pCurPoint)
    X = pCurPoint.X * Screen.TwipsPerPixelX
    Y = pCurPoint.Y * Screen.TwipsPerPixelY
End Sub


Public Function AddNode(Optional ByRef NodeKey As String, Optional ByVal Caption As String, Optional ByVal CentreX As Double, Optional ByVal CentreY As Double) As Object
Attribute AddNode.VB_Description = "Add picture as node to graph"
    Dim pControlObject As New ControlWrapper
    If m_ObjectCollections Is Nothing Then UserControl_InitProperties
    m_PictureIndex = m_PictureIndex + 1
    Load LPicture(m_PictureIndex)
    
    If (Len(NodeKey) <= 0) Then NodeKey = CreateGUID()
    
    Set pControlObject.ControlObject = LPicture(m_PictureIndex - 1).object
    pControlObject.CtrlType = oTPicture
    pControlObject.Index = m_PictureIndex - 1
    m_ObjectCollections.Add pControlObject, NodeKey
    LPicture(m_PictureIndex - 1).NodeKey = NodeKey
    LPicture(m_PictureIndex - 1).Caption = Caption
    LPicture(m_PictureIndex - 1).CentreX = CentreX
    LPicture(m_PictureIndex - 1).CentreY = CentreY
    Set AddNode = LPicture(m_PictureIndex - 1)
    Set pControlObject = Nothing
End Function

Public Function AddText(Optional ByRef TextKey As String, Optional ByVal Caption As String, Optional ByVal CentreX As Double, Optional ByVal CentreY As Double) As Object
Attribute AddText.VB_Description = "Add text objet to graph"
    Dim pControlObject As New ControlWrapper
    
    If (m_ObjectCollections Is Nothing) Then UserControl_InitProperties
    m_TextIndex = m_TextIndex + 1
    Load LText(m_TextIndex)
    If (Len(TextKey) <= 0) Then TextKey = CreateGUID()
    
    Set pControlObject.ControlObject = LText(m_TextIndex - 1).object
    pControlObject.CtrlType = oTText
    pControlObject.Index = m_TextIndex - 1
    
    m_ObjectCollections.Add pControlObject, TextKey
    LText(m_TextIndex - 1).NodeKey = TextKey
    LText(m_TextIndex - 1).Caption = Caption
    Call LText(m_TextIndex - 1).ScaleRegion(m_Zoom / 100)
    LText(m_TextIndex - 1).CentreX = CentreX
    LText(m_TextIndex - 1).CentreY = CentreY
    
    'Call ResizeObjects(LText(m_TextIndex - 1))
    Set AddText = LText(m_TextIndex - 1)
    Set pControlObject = Nothing
End Function

Public Function AddStep(ByVal StepType As eLineType, ByVal FromNodeKey As String, ByVal ToNodeKey As String, Optional ByRef StepKey As String) As Object
Attribute AddStep.VB_Description = "Join two nodes (pictures) with line"
    Dim pControlObject As New ControlWrapper
    Dim pTaskObject As oGraph.oPicture
    
    If (m_ObjectCollections Is Nothing) Then UserControl_InitProperties
    
    m_LineIndex = m_LineIndex + 1
    Load LLine(m_LineIndex)
    
    If (Len(StepKey) <= 0) Then StepKey = CreateGUID()
    
    Set pControlObject.ControlObject = LLine(m_LineIndex - 1).object
    pControlObject.CtrlType = oTLine
    pControlObject.Index = m_LineIndex - 1
    m_ObjectCollections.Add pControlObject, StepKey
    LLine(m_LineIndex - 1).LayereLineType = StepType
    If (StepType = OnTCompletion) Then
        LLine(m_LineIndex - 1).ToolTipText = "On Completion"
    ElseIf (StepType = OnTSuccess) Then
        LLine(m_LineIndex - 1).ToolTipText = "On Success"
    ElseIf (StepType = OnTFail) Then
        LLine(m_LineIndex - 1).ToolTipText = "On Failure"
    Else
        LLine(m_LineIndex - 1).ToolTipText = "Unknown"
    End If
    LLine(m_LineIndex - 1).NodeKey = StepKey
    
    Set LLine(m_LineIndex - 1).ConnectedFrom = m_ObjectCollections.Item(FromNodeKey).ControlObject
    Set LLine(m_LineIndex - 1).ConnectedTo = m_ObjectCollections.Item(ToNodeKey).ControlObject
    Set pTaskObject = LLine(m_LineIndex - 1).ConnectedFrom
    pTaskObject.LinkedStepKey.Add StepKey
    
    Set pTaskObject = LLine(m_LineIndex - 1).ConnectedTo
    pTaskObject.LinkedStepKey.Add StepKey
    Call LLine(m_LineIndex - 1).ScaleRegion(m_Zoom / 100)
    'Call ResizeObjects(LLine(m_LineIndex - 1))
    Set AddStep = LLine(m_LineIndex - 1)
    Set pControlObject = Nothing
    Set pTaskObject = Nothing
End Function

Private Sub m_picoConvas_KeyDown(KeyCode As Integer, Shift As Integer)
     If Shift = vbCtrlMask And (Chr$(KeyCode) = "A" Or Chr$(KeyCode) = "a") Then
        Call SelectAll
    End If
End Sub

Private Sub SelectAll()
    Dim pCtrlObject As ControlWrapper
    If Not m_ObjectCollections Is Nothing Then
        For Each pCtrlObject In m_ObjectCollections
            pCtrlObject.ControlObject.oSelected = True
        Next
    End If
    PaintTimer.Enabled = True
End Sub

Public Sub Paint()
Attribute Paint.VB_Description = "Paint the entire graph"
    PaintTimer.Enabled = True
End Sub

Private Sub m_picoConvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_X1MouseDown = X
    m_Y1MouseDown = Y
    If Not (m_LockLayerEdit) Then
        If Not IsSelected And Shift <> vbCtrlMask And Button = vbLeftButton Then
            m_Drawing = True
            Call DrawTempLine(X, Y, X, Y)
        End If
    End If
    RaiseEvent MouseDown(oTConvas, "", Button, Shift, CDbl(X), CDbl(Y))
End Sub

Private Sub DrawTempLine(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
    m_picoConvas.DrawStyle = DrawStyleConstants.vbDot
    m_picoConvas.DrawMode = DrawModeConstants.vbInvert
    m_picoConvas.Line (m_X1Old, m_Y1Old)-(m_X2Old, m_Y2Old), , B
    If (m_Drawing = True) Then
        m_picoConvas.Line (X1, Y1)-(X2, Y2), , B
        m_X1Old = X1
        m_Y1Old = Y1
        m_X2Old = X2
        m_Y2Old = Y2
    Else
        m_X1Old = 0
        m_Y1Old = 0
        m_X2Old = 0
        m_Y2Old = 0
    End If
   
End Sub


Private Sub UnSelectAllObjects()
    Dim pObject As ControlWrapper
    Set m_SelectedLineKeys = Nothing
    Set m_SelectedPicKeys = Nothing
    Set m_SelectedTextKeys = Nothing
    
    If Not (m_ObjectCollections Is Nothing) Then
        If (IsSelected) Then
            For Each pObject In m_ObjectCollections
                pObject.ControlObject.oSelected = False
            Next
        End If
    End If
    Set m_SelectedLineKeys = New Collection
    Set m_SelectedPicKeys = New Collection
    Set m_SelectedTextKeys = New Collection
    m_IsSelected = False
End Sub


Private Function IsSelected() As Boolean
   Dim pObjectCtrl As ControlWrapper
   If Not (m_ObjectCollections Is Nothing) Then
        For Each pObjectCtrl In m_ObjectCollections
         If Not pObjectCtrl Is Nothing Then
            If (pObjectCtrl.ControlObject.oSelected = True) Then
                 IsSelected = True
            End If
        End If
        Next
   End If
End Function

Private Sub m_picoConvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        If Shift <> vbCtrlMask Then
            If m_picoConvas.MousePointer = MousePointerConstants.vbCustom Then
                m_picoConvas.MousePointer = MousePointerConstants.vbDefault
            End If
        Else
            If m_picoConvas.MousePointer <> MousePointerConstants.vbCustom Then
                m_picoConvas.MousePointer = MousePointerConstants.vbCustom
            End If
        End If
        If IsSelected And Shift <> vbCtrlMask And Button = vbLeftButton Then
            m_Drawing = False
            Call DrawTempLine(m_X1Old, m_Y1Old, X, Y)
            Exit Sub
        End If
        
        If Not (m_LockLayerEdit) Then
            If m_Drawing = True Then
                Call DrawTempLine(m_X1Old, m_Y1Old, X, Y)
            End If
        End If
        RaiseEvent MouseMove(oTConvas, "", Button, Shift, CDbl(X), CDbl(Y))
        
End Sub

Private Sub m_picoConvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pKey As Variant
    
        If Not (m_LockLayerEdit) Then
            If m_Drawing Then
                m_Drawing = False
                Call DrawTempLine(m_X1Old, m_Y1Old, X, Y)
                If (Abs(m_X1MouseDown - X) > 5 And Abs(m_Y1MouseDown - Y) > 5) Then
                    Call SelectInsideRect(m_X1MouseDown, m_Y1MouseDown, X, Y)
                End If
            End If
        End If
        
        If Button = vbLeftButton And Shift = vbCtrlMask Then
               If Not (Abs(m_X1MouseDown - X) = 0 And Abs(m_Y1MouseDown - Y) = 0) Then
                   RealignTimer.Enabled = True
                End If
        End If
        
        RaiseEvent MouseUp(oTConvas, "", Button, Shift, CDbl(X), CDbl(Y))
        If Button = MouseButtonConstants.vbRightButton Then
            RaiseEvent RightClick(oTConvas, "", X, Y)
        End If
End Sub

Private Sub ReAlign()
    Dim pKey As Variant
    Dim pObject As ControlWrapper
    Dim pPicture As oGraph.oPicture
    If Not m_SelectedPicKeys Is Nothing Then
    For Each pKey In m_SelectedPicKeys
        Set pObject = m_ObjectCollections.Item(pKey)
        If pObject.ControlObject.oSelected Then
            Set pPicture = pObject.ControlObject
            pPicture.ReAlign
        End If
    Next
    End If
    Me.Paint
End Sub

Private Sub SelectInsideRect(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double)
    Dim pCtrlObject As ControlWrapper
    Dim CentreX As Double, CentreY As Double, oHeight As Double, oWidth As Double
    Dim p1 As Double
    Dim p2 As Double
    On Error Resume Next
    p1 = IIf(X1 > X2, X2, X1)
    p2 = IIf(X2 > X1, X2, X1)
    X1 = p1
    X2 = p2
    
    p1 = IIf(Y1 > Y2, Y2, Y1)
    p2 = IIf(Y2 > Y1, Y2, Y1)
    Y1 = p1
    Y2 = p2
    Call UnSelectAllObjects
    If Not m_ObjectCollections Is Nothing Then
        For Each pCtrlObject In m_ObjectCollections
           Call pCtrlObject.ControlObject.GetCordinates(CentreX, CentreY, oWidth, oHeight, True)
           If (X1 <= CentreX - oWidth / 2 And _
           X2 >= CentreX + oWidth / 2 And _
           Y1 <= CentreY - oHeight / 2 And _
           Y2 >= CentreY + oHeight / 2) Then
               pCtrlObject.ControlObject.oSelected = True
               If pCtrlObject.CtrlType = oTLine Then
                    m_SelectedLineKeys.Add LLine(pCtrlObject.Index).NodeKey, LLine(pCtrlObject.Index).NodeKey
               ElseIf pCtrlObject.CtrlType = oTPicture Then
                    m_SelectedPicKeys.Add LPicture(pCtrlObject.Index).NodeKey, LPicture(pCtrlObject.Index).NodeKey
               ElseIf pCtrlObject.CtrlType = oTText Then
                    m_SelectedTextKeys.Add LText(pCtrlObject.Index).NodeKey, LText(pCtrlObject.Index).NodeKey
               End If
               m_SelectionByRect = True
           Else
               pCtrlObject.ControlObject.oSelected = False
           End If
        Next
        PaintTimer.Enabled = True
    End If
    
End Sub


Public Sub ClearWorkSheet()
Attribute ClearWorkSheet.VB_Description = "Clear the worksheet"
    Dim pCtrlWrapper As ControlWrapper
    Dim pControl As oGraph.oPicture

    For Each pCtrlWrapper In m_ObjectCollections
        If pCtrlWrapper.CtrlType = oTLine Then
            If pCtrlWrapper.Index > 0 Then
                Unload LLine(pCtrlWrapper.Index)
            Else
                LLine(pCtrlWrapper.Index).Visible = False
            End If
        ElseIf pCtrlWrapper.CtrlType = oTPicture Then
            If pCtrlWrapper.Index > 0 Then
                Unload LPicture(pCtrlWrapper.Index)
            Else
                LPicture(pCtrlWrapper.Index).Visible = False
            End If
        ElseIf pCtrlWrapper.CtrlType = oTText Then
            If pCtrlWrapper.Index > 0 Then
                Unload LText(pCtrlWrapper.Index)
            Else
                LText(pCtrlWrapper.Index).Visible = False
            End If
        End If
    Next
    Do While LText.Count > 1
        Unload LText(m_TextIndex)
        m_TextIndex = m_TextIndex - 1
    Loop
    
    Do While LPicture.Count > 1
        Unload LPicture(m_PictureIndex)
        m_PictureIndex = m_PictureIndex - 1
    Loop
    
    Do While LLine.Count > 1
        Unload LLine(m_LineIndex)
        m_LineIndex = m_LineIndex - 1
    Loop
    
    m_TextIndex = 0
    m_PictureIndex = 0
    m_LineIndex = 0
    Set m_SelectedLineKeys = Nothing
    Set m_SelectedPicKeys = Nothing
    Set m_SelectedTextKeys = Nothing
    Set m_ObjectCollections = Nothing
    
    Set m_ObjectCollections = New Collection
    Set m_SelectedLineKeys = New Collection
    Set m_SelectedPicKeys = New Collection
    Set m_SelectedTextKeys = New Collection
    Me.Paint
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Support For Each keyword"
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_ObjectCollections.[_NewEnum]
End Property

Public Property Get Item(vntIndexKey As Variant) As ControlWrapper
Attribute Item.VB_Description = "Returns Graph object (Line/Picture/Text) by node key"
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = m_ObjectCollections(vntIndexKey)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Number of graph objects (Line/Picture/Text)"
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = m_ObjectCollections.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "Remove object from graph"
    Dim pCtrlWrapper As ControlWrapper
    Dim pTaskObject As oGraph.oPicture
    Dim pStepKey As StepKey
    Dim pStepKeys As StepKeys
    Dim pCtrlIndex As Integer
    Dim pType As eObjectType
    Dim pKey As String
    
    
    Set pCtrlWrapper = m_ObjectCollections.Item(vntIndexKey)
    pCtrlIndex = pCtrlWrapper.Index
    pType = pCtrlWrapper.CtrlType
    pKey = pCtrlWrapper.ControlObject.NodeKey
    Set pCtrlWrapper = Nothing
    
    
    
    If (pType = oTPicture) Then
        Set pStepKeys = LPicture(pCtrlIndex).LinkedStepKey
        If pStepKeys.Count > 0 Then
            For Each pStepKey In pStepKeys
            
                Set pTaskObject = LLine(m_ObjectCollections.Item(pStepKey.Key).Index).ConnectedFrom
                If Not (pTaskObject Is Nothing) Then
                    pTaskObject.LinkedStepKey.Remove pStepKey.Key
                    Set LLine(m_ObjectCollections.Item(pStepKey.Key).Index).ConnectedFrom = Nothing
                    Set pTaskObject = Nothing
                End If
                
                Set pTaskObject = LLine(m_ObjectCollections.Item(pStepKey.Key).Index).ConnectedTo
                If Not (pTaskObject Is Nothing) Then
                    pTaskObject.LinkedStepKey.Remove pStepKey.Key
                    Set LLine(m_ObjectCollections.Item(pStepKey.Key).Index).ConnectedTo = Nothing
                    Set pTaskObject = Nothing
                End If
                If (m_ObjectCollections.Item(pStepKey.Key).Index > 0) Then
                    Unload LLine(m_ObjectCollections.Item(pStepKey.Key).Index)
                Else
                    LLine(m_ObjectCollections.Item(pStepKey.Key).Index).Visible = False
                End If
                m_ObjectCollections.Remove pStepKey.Key
            Next
        End If
        Set pStepKeys = Nothing
        If (pCtrlIndex > 0) Then
            Unload LPicture(pCtrlIndex)
        Else
            LPicture(pCtrlIndex).Visible = False
        End If
        m_ObjectCollections.Remove vntIndexKey
        m_SelectedPicKeys.Remove vntIndexKey
    ElseIf pType = oTLine Then
        
        pKey = LLine(pCtrlIndex).NodeKey
        Set pTaskObject = LLine(pCtrlIndex).ConnectedFrom
        pTaskObject.LinkedStepKey.Remove pKey
        
        Set pTaskObject = LLine(pCtrlIndex).ConnectedTo
        pTaskObject.LinkedStepKey.Remove pKey
        
        m_ObjectCollections.Remove pKey
        m_SelectedLineKeys.Remove vntIndexKey
        If pCtrlIndex > 0 Then
            Unload LLine(pCtrlIndex)
            'm_LineIndex = m_LineIndex - 1
        Else
            LLine(pCtrlIndex).Visible = False
        End If
    ElseIf pType = oTText Then
        m_ObjectCollections.Remove LText(pCtrlIndex).NodeKey
        m_SelectedTextKeys.Remove vntIndexKey
        If pCtrlIndex > 0 Then
            LText(pCtrlIndex).Visible = False
            Unload LText(pCtrlIndex)
            'm_TextIndex = m_TextIndex - 1
        Else
            LText(pCtrlIndex).Visible = False
        End If
    End If
    
    Set pTaskObject = Nothing
    Set pCtrlWrapper = Nothing
    UserControl.MousePointer = vbDefault
End Sub

Public Property Get LockLayerEdit() As Boolean
Attribute LockLayerEdit.VB_Description = "Lock graph layer so that user can't draw selection rect"
    LockLayerEdit = m_LockLayerEdit
End Property

Public Property Let LockLayerEdit(ByVal New_LockLayerEdit As Boolean)
    m_LockLayerEdit = New_LockLayerEdit
End Property


Private Sub PaintTimer_Timer()
    Dim pRect As RECT
    Dim poGraphObject As ControlWrapper
    Dim pMemDc As cMemDC
    Dim pRegion As Long
    
        If m_LoadingFromBinary Then
            PaintTimer.Enabled = False
            Exit Sub
        End If
        If m_AlreadyPainting Then Exit Sub
        m_AlreadyPainting = True
        PaintTimer.Enabled = False
        pRect.Top = 0
        pRect.Left = 0
        pRect.Right = UserControl.Width \ 15
        pRect.Bottom = UserControl.Height \ 15
        If m_MemoryDC Is Nothing Then
            PaintTimer.Enabled = False
            Exit Sub
        End If
        If m_oGraphics = 0 Then
            Call UserControl_Resize
        End If
        Call GdipGraphicsClear(m_oGraphics, GetGDIColorFromOLE(m_BackColor))
        If Not m_ObjectCollections Is Nothing Then
            For Each poGraphObject In m_ObjectCollections
                If Not poGraphObject Is Nothing Then poGraphObject.ControlObject.Paint
            Next
        End If
        
        Call m_MemoryDC.PaintPicture(m_picoConvas.hdc)
        m_AlreadyPainting = False
        m_ReadyForPaint = True
End Sub

Public Sub AutoLayout()
    Dim pcount As Single
    Dim pGridSize As POINTL
    Dim pRow As Single
    Dim pCol As Single
    Dim pIndex As Single
    Dim pInnerIndex As Single
    Dim pBlockHeight As Single
    Dim pBlockWidth As Single
    Dim pGridPoints As Collection
    Dim pPoint As SidePoint
    Dim pControlWrapper As ControlWrapper
    Dim pObject As Object
    Dim pCentX As Double, pCentY As Double, pWidth As Double, pHeight As Double
    Dim pCalRow As Single
    Dim pCalCol As Single
    
    Set pGridPoints = New Collection
    If m_ObjectCollections Is Nothing Then Set m_ObjectCollections = New Collection
    pcount = 10 * m_ObjectCollections.Count
    LSet pGridSize = GetFactors(pcount)
    If UserControl.Height > UserControl.Width Then
        pRow = IIf(pGridSize.X > pGridSize.Y, pGridSize.X, pGridSize.Y)
        pCol = IIf(pGridSize.X > pGridSize.Y, pGridSize.Y, pGridSize.Y)
    Else
        pCol = IIf(pGridSize.X > pGridSize.Y, pGridSize.X, pGridSize.Y)
        pRow = IIf(pGridSize.X > pGridSize.Y, pGridSize.Y, pGridSize.Y)
    End If
    pBlockHeight = picControl.Height / (pRow)
    pBlockWidth = picControl.Width / (pCol)
    For pIndex = 1 To pRow
        For pInnerIndex = 1 To pCol
            Set pPoint = New SidePoint
            pPoint.X = pInnerIndex * pBlockWidth - pBlockWidth / 2
            pPoint.Y = pIndex * pBlockHeight - pBlockHeight / 2
            pPoint.IsVacant = True
            pGridPoints.Add pPoint, "K" & pIndex - 1 & "," & pInnerIndex - 1
        Next
    Next
    On Error Resume Next
    If Not m_ObjectCollections Is Nothing Then
        For Each pControlWrapper In m_ObjectCollections
            If pControlWrapper.CtrlType = oTPicture Or pControlWrapper.CtrlType = oTText Then
                Set pObject = pControlWrapper.ControlObject
                Call pObject.GetCordinates(pCentX, pCentY, pWidth, pHeight, True)
                pCalRow = Fix(pCentY / pBlockHeight)
                pCalCol = Fix(pCentX / pBlockWidth)
                Set pPoint = pGridPoints.Item("K" & pCalRow & "," & pCalCol)
                If pPoint.IsVacant = True Then
                    pPoint.IsVacant = False
                    pObject.CentreX = pPoint.X
                    pObject.CentreY = pPoint.Y
                    If pControlWrapper.CtrlType = oTPicture Then
                        pObject.CordsChanged
                    End If
                Else
                    Do Until Err.Number <> 0
                        For pIndex = pCalRow To pRow - 1
                            For pInnerIndex = pCalCol To pCol - 1
                                Set pPoint = pGridPoints.Item("K" & pIndex & "," & pInnerIndex)
                                If Err.Number = 0 Then
                                    If pPoint.IsVacant = True Then
                                        pPoint.IsVacant = False
                                        pObject.CentreX = pPoint.X
                                        pObject.CentreY = pPoint.Y
                                        If pControlWrapper.CtrlType = oTPicture Then
                                            pObject.CordsChanged
                                        End If
                                        Exit Do
                                    End If
                                Else
                                    Err.Clear
                                End If
                            Next
                        Next
                        Exit Do
                    Loop
                End If
            End If
        Next
    End If
    Me.Paint
End Sub


Private Function GetFactors(ByVal pNumber As Single) As POINTL
    Dim pFact1 As Single
    Dim pFact2 As Single
    Dim pNum As Double
    Dim pDiff As Single
    Dim pX As Single
    Dim pY As Single
    pX = 1
    pY = IIf(pNumber > 0, pNumber, 1)
    pDiff = pNumber - 1
    For pFact1 = 1 To pNumber
        pNum = pNumber / pFact1
        'If CInt(pNum) = pNum Then
            pFact2 = IIf(pNum > Fix(pNum), Fix(pNum) + 1, Fix(pNum))
            If Abs(pFact1 - pFact2) <= pDiff Then
                pX = pFact1
                pY = pFact2
                pDiff = Abs(pX - pY)
            End If
        'End If
    Next
    GetFactors.X = pX
    GetFactors.Y = pY
End Function


Private Sub m_picoConvas_Paint()
    If Not m_AlreadyPainting Then
        If Not m_MemoryDC Is Nothing Then
            If m_ReadyForPaint Then
                Call m_MemoryDC.PaintPicture(m_picoConvas.hdc)
            End If
        End If
    End If
End Sub



Private Sub RealignTimer_Timer()
    RealignTimer.Enabled = False
    Call ReAlign
End Sub

Private Sub SelectionTimer_Timer()
    SelectionTimer.Enabled = False
    If m_ClickCount <= 1 Then
        UnSelectAllObjects
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    Set m_ObjectCollections = New Collection
    Set m_SelectedPicKeys = New Collection
    Set m_SelectedTextKeys = New Collection
    Set m_SelectedLineKeys = New Collection
   picoConvas.Width = 2 * picControl.Width
   picoConvas.Left = 0
   picoConvas.Top = 0
   picoConvas.Height = 2 * picControl.Height
   
    Set m_picoConvas = picoConvas
    m_ForeColor = &H0
    m_BackColor = &HFFFFFF
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Zoom = 100
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H0)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
End Sub

Friend Property Get PictureBuffer() As PictureBox
    Set PictureBuffer = picBuffer
End Property

Private Sub MaxExtendChanged()
    If m_MaxHeight > picControl.Height Then
        BarVs.Enabled = True
    Else
        BarVs.Enabled = True
    End If
    If m_MaxWidth > picControl.Width Then
        BarHs.Enabled = True
    Else
        BarHs.Enabled = True
    End If
    
    'BarHs.Min = picControl.Width / 15
    'BarHs.Max = IIf(m_MaxWidth / 15 > BarHs.Min, m_MaxWidth / 15, BarHs.Min)
    'BarHs.value = BarHs.Min
    'BarHs.SmallChange = 15
    'BarHs.LargeChange = 15 * 3
   '
    'BarVs.Min = picControl.Height / 15
    'BarVs.Max = IIf(m_MaxHeight / 15 > BarVs.Min, m_MaxHeight / 15, BarVs.Min)
    'BarVs.SmallChange = 15
    'BarVs.LargeChange = 15 * 3
    'BarVs.value = BarVs.Min
End Sub

Private Sub UserControl_Resize()
   Dim pRect As RECT
   Dim pCtrlObject As ControlWrapper
    With picControl 'picoConvas
        .Top = 0
        .Left = 0
        .Width = UserControl.Width - BarVs.Width - 20
        .Height = UserControl.Height - BarHs.Height - 20
        BarHs.Left = 0
        BarHs.Top = .Height
        BarHs.Width = .Width
        
        BarVs.Height = .Height
        BarVs.Top = 0
        BarVs.Left = .Width
    End With
    
   picoConvas.Width = 2 * picControl.Width
   picoConvas.Left = 0
   picoConvas.Top = 0
   picoConvas.Height = 2 * picControl.Height
   
    If m_MaxHeight < picControl.Height Then
        m_MaxHeight = picControl.Height
    End If
    
    If m_MaxWidth < picControl.Width Then
        m_MaxWidth = picControl.Width
    End If
    Call MaxExtendChanged
    If (UserControl.Ambient.UserMode) Then
        m_ReadyForPaint = False
        Set m_picoConvas = picoConvas
        PaintTimer.Enabled = False
        If m_oGraphics <> 0 Then
            Call GdipDeleteGraphics(m_oGraphics)
            m_oGraphics = 0
        End If
        'Call UserControl_Terminate
        
        Set m_MemoryDC = New cDibSection
        With m_MemoryDC
            Call .Create((picoConvas.Width \ Screen.TwipsPerPixelY), (picoConvas.Height \ Screen.TwipsPerPixelX))
        End With
        Call modGDIPlus.GDIPlusCreate
        Call GdipCreateFromHDC(m_MemoryDC.hdc, m_oGraphics)
        Call GdipSetSmoothingMode(m_oGraphics, SmoothingModeAntiAlias)
        Call GdipSetTextRenderingHint(m_oGraphics, TextRenderingHintAntiAliasGridFit)
        If Not m_ObjectCollections Is Nothing Then
            For Each pCtrlObject In m_ObjectCollections
                pCtrlObject.ControlObject.ReSizeoConvas
            Next
        End If
   Else
        Call UserControl_Terminate
   End If
   PaintTimer.Enabled = True
End Sub

Public Property Get BackColor() As OLE_COLOR
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

Private Sub UserControl_Terminate()
    PaintTimer.Enabled = False
    Set m_picoConvas = Nothing
    If m_OrgMatrix <> 0 Then
        Call GdipDeleteMatrix(m_OrgMatrix)
    End If
    If m_oGraphics <> 0 Then
        Call GdipDeleteGraphics(m_oGraphics)
        Call modGDIPlus.GDIPlusDispose
        m_oGraphics = 0
    End If
    Set m_MemoryDC = Nothing
    Set m_ObjectCollections = Nothing
    Set m_SelectedLineKeys = Nothing
    Set m_SelectedPicKeys = Nothing
    Set m_SelectedTextKeys = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &HF000000)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFFF)
End Sub


Public Property Get BinaryData() As Variant
    Dim pPropBag As New PropertyBag
    Dim lC As Single
    Dim Pc As Single
    Dim Tc As Single
    Dim pcount As Single
    Dim pControlObject As ControlWrapper
    On Error Resume Next
        For pcount = 0 To LText.Count - 1
            
                Set pControlObject = New ControlWrapper
                pControlObject.CtrlType = oTText
                pControlObject.Index = pcount
                Set pControlObject.ControlObject = LText(pcount).object
                If Not pControlObject.ControlObject Is Nothing Then
                    pPropBag.WriteProperty "Td" & pcount, pControlObject.ControlObject
                End If
                
        Next
        
        For pcount = 0 To LPicture.Count - 1
            Set pControlObject = New ControlWrapper
            pControlObject.CtrlType = oTPicture
            pControlObject.Index = pcount
            Set pControlObject.ControlObject = LPicture(pcount).object
            If Not pControlObject.ControlObject Is Nothing Then
                pPropBag.WriteProperty "Pd" & pcount, pControlObject.ControlObject
            End If
        Next
        
        For pcount = 0 To LLine.Count - 1
            Set pControlObject = New ControlWrapper
            pControlObject.CtrlType = oTLine
            pControlObject.Index = pcount
            Set pControlObject.ControlObject = LLine(pcount).object
            If Not pControlObject.ControlObject Is Nothing Then
                pPropBag.WriteProperty "Ld" & pcount, pControlObject.ControlObject
            End If
        Next
        pPropBag.WriteProperty "Lc", LLine.Count - 1
        pPropBag.WriteProperty "Pc", LPicture.Count - 1
        pPropBag.WriteProperty "Tc", LText.Count - 1
        pPropBag.WriteProperty "oConvas", Me, Nothing
        
        BinaryData = pPropBag.Contents
        Set pPropBag = Nothing
End Property

Public Property Let BinaryData(ByVal pData As Variant)
    Dim pPropBag As New PropertyBag
    Dim lC As Single
    Dim Pc As Single
    Dim Tc As Single
    
    Dim Ld As Object
    Dim Pd As Object
    Dim Td As Object
    Dim poConvas As oConvas
    Dim pControlObject As ControlWrapper
    Dim pcount As Single
    
    Dim pLine As Object
    Dim pPicture As Object
    Dim pText As Object
    
    m_LoadingFromBinary = True
    If m_ObjectCollections Is Nothing Then Set m_ObjectCollections = New Collection
    pPropBag.Contents = pData
    
    lC = pPropBag.ReadProperty("Lc", 0)
    Pc = pPropBag.ReadProperty("Pc", 0)
    Tc = pPropBag.ReadProperty("Tc", 0)
    Set poConvas = pPropBag.ReadProperty("oConvas", Nothing)
    
    Dim pKey As String
    
    For pcount = 0 To Tc - 1
        Set Td = pPropBag.ReadProperty("Td" & pcount)
        m_TextIndex = m_TextIndex + 1
        Load LText(m_TextIndex)
        Set pText = LText(m_TextIndex - 1)
        Set pControlObject = New ControlWrapper
        Set pControlObject.ControlObject = LText(m_TextIndex - 1).object
        pControlObject.CtrlType = oTText
        pControlObject.Index = m_TextIndex - 1
        If Td.Visible Then
            m_ObjectCollections.Add pControlObject, Td.NodeKey
        End If
        With pText
            .Caption = Td.Caption
            .CentreX = Td.CentreX
            .CentreY = Td.CentreY
            .BackColor = Td.BackColor
            Set .Font = Td.Font
            .ForeColor = Td.ForeColor
            .IsOutlined = Td.IsOutlined
            .IsTransparent = Td.IsTransparent
            .tHeight = Td.tHeight
            .NodeKey = Td.NodeKey
            .oSelected = Td.oSelected
            .ToolTipText = Td.ToolTipText
            .Transparency = Td.Transparency
            .tWidth = Td.tWidth
            .Visible = Td.Visible
            
        End With
        Set Td = Nothing
        Set pText = Nothing
    Next
    
    For pcount = 0 To Pc - 1
        Set Pd = pPropBag.ReadProperty("Pd" & pcount)
        m_PictureIndex = m_PictureIndex + 1
        Load LPicture(m_PictureIndex)
        Set pPicture = LPicture(m_PictureIndex - 1)
        Set pControlObject = New ControlWrapper
        Set pControlObject.ControlObject = LPicture(m_PictureIndex - 1).object
        pControlObject.CtrlType = oTPicture
        pControlObject.Index = m_PictureIndex - 1
        If Pd.Visible Then
            m_ObjectCollections.Add pControlObject, Pd.NodeKey
        End If
        With pPicture
            .Caption = Pd.Caption
            .CentreX = Pd.CentreX
            .CentreY = Pd.CentreY
            .IsConnection = Pd.IsConnection
             Set .Font = Pd.Font
            .ForeColor = Pd.ForeColor
            Set .Image = Pd.Image
            .NodeKey = Pd.NodeKey
            .oSelected = Pd.oSelected
            .PointPerSide = Pd.PointPerSide
            Set .LinkedStepKey = Pd.LinkedStepKey
            .ToolTipText = Pd.ToolTipText
            .Activate
            .Visible = Pd.Visible
        End With
        Set Pd = Nothing
        Set pPicture = Nothing
    Next
    
    For pcount = 0 To lC - 1
        Set Ld = pPropBag.ReadProperty("Ld" & pcount)
        m_LineIndex = m_LineIndex + 1
        Load LLine(m_LineIndex)
        Set pLine = LLine(m_LineIndex - 1)
        Set pControlObject = New ControlWrapper
        Set pControlObject.ControlObject = LLine(m_LineIndex - 1).object
        pControlObject.CtrlType = oTLine
        pControlObject.Index = m_LineIndex - 1
        If Ld.Visible Then
            m_ObjectCollections.Add pControlObject, Ld.NodeKey
        End If
        With pLine
            Set pControlObject = m_ObjectCollections.Item(Ld.FromKey)
            Set .ConnectedFrom = LPicture(pControlObject.Index).object
            Set pControlObject = m_ObjectCollections.Item(Ld.ToKey)
            Set .ConnectedTo = LPicture(pControlObject.Index).object
            .ConstraintIndex = Ld.ConstraintIndex
            .oSelected = Ld.oSelected
            .LayereLineType = Ld.LayereLineType
            .NodeKey = Ld.NodeKey
            .StepName = Ld.StepName
            .ToolTipText = Ld.ToolTipText
            .SetPoints Ld.Points
            .Visible = Ld.Visible
        End With
        Set Ld = Nothing
        Set pLine = Nothing
    Next
    Me.BackColor = poConvas.BackColor
    Me.ForeColor = poConvas.ForeColor
    m_LoadingFromBinary = False
End Property

Public Function GetSelectedCount(Optional ByVal pObjectType As eObjectType = 0) As Single
    On Error Resume Next
    If pObjectType = 0 Then
        GetSelectedCount = GetSelectedCount + m_SelectedPicKeys.Count
        GetSelectedCount = GetSelectedCount + m_SelectedTextKeys.Count
        GetSelectedCount = GetSelectedCount + m_SelectedLineKeys.Count
    Else
        If pObjectType = oTPicture Then
            GetSelectedCount = m_SelectedPicKeys.Count
        ElseIf pObjectType = oTLine Then
            GetSelectedCount = m_SelectedLineKeys.Count
        Else
            GetSelectedCount = m_SelectedTextKeys.Count
        End If
    End If
End Function

Public Function GetSelectedKeys(Optional pObjectType As eObjectType = 0) As Collection
    Dim pCombinedCol As Collection
    Dim pKey As Variant
    If (pObjectType = 0) Then
        Call GetSelect(oTLine)
        Call GetSelect(oTPicture)
        Call GetSelect(oTText)
        Set pCombinedCol = New Collection
        On Error Resume Next
        If Not m_SelectedLineKeys Is Nothing Then
            For Each pKey In m_SelectedLineKeys
                pCombinedCol.Add pKey, pKey
            Next
        End If
        
        If Not m_SelectedPicKeys Is Nothing Then
            For Each pKey In m_SelectedPicKeys
                pCombinedCol.Add pKey, pKey
            Next
        End If
        
        If Not m_SelectedTextKeys Is Nothing Then
            For Each pKey In m_SelectedTextKeys
                pCombinedCol.Add pKey, pKey
            Next
        End If
        
        Set GetSelectedKeys = pCombinedCol
        Set pCombinedCol = Nothing
    Else
        Call GetSelect(pObjectType)
        If pObjectType = oTPicture Then
            Set GetSelectedKeys = m_SelectedPicKeys
        ElseIf pObjectType = oTLine Then
            Set GetSelectedKeys = m_SelectedLineKeys
        Else
            Set GetSelectedKeys = m_SelectedTextKeys
        End If
    End If
End Function

Private Sub GetSelect(ByVal pSelType As eObjectType)
    Dim pCtrlObject As ControlWrapper
    Dim pKeyCol As Collection
    Set pKeyCol = New Collection
    If Not m_ObjectCollections Is Nothing Then
        For Each pCtrlObject In m_ObjectCollections
            If pCtrlObject.CtrlType = pSelType Then
                If pCtrlObject.ControlObject.oSelected Then
                    pKeyCol.Add pCtrlObject.ControlObject.NodeKey, pCtrlObject.ControlObject.NodeKey
                End If
            End If
        Next
    End If
    Select Case pSelType
        Case eObjectType.oTLine:
            Set m_SelectedLineKeys = Nothing
            Set m_SelectedLineKeys = pKeyCol
        Case eObjectType.oTPicture:
            Set m_SelectedPicKeys = Nothing
            Set m_SelectedPicKeys = pKeyCol
        Case eObjectType.oTText:
            Set m_SelectedTextKeys = Nothing
            Set m_SelectedTextKeys = pKeyCol
    End Select
    
End Sub

Public Property Let Zoom(ByVal pZoom As Single)
        Dim poGraphObject As ControlWrapper
        Dim pZoomFactor As Double
        Dim pReverseZoom As Double
        Static pPreviousZoom As Single
        If pZoom < 60 Or pZoom > 200 Then Exit Property
        If pPreviousZoom = 0 Then pPreviousZoom = 100
        m_Zoom = pZoom
        pZoomFactor = (m_Zoom) / 100
        pReverseZoom = 1 / (pPreviousZoom / 100)
        If Not m_ObjectCollections Is Nothing Then
            For Each poGraphObject In m_ObjectCollections
                Call poGraphObject.ControlObject.ScaleRegion(pReverseZoom)
                Call poGraphObject.ControlObject.ScaleRegion(pZoomFactor)
            Next
        End If
        pPreviousZoom = m_Zoom
End Property

Public Property Get Zoom() As Single
    Zoom = m_Zoom
End Property
