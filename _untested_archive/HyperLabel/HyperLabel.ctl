VERSION 5.00
Begin VB.UserControl HyperLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ForwardFocus    =   -1  'True
   MouseIcon       =   "HyperLabel.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   1725
End
Attribute VB_Name = "HyperLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : HyperLabel
' DateTime  : 04 jan 2006 03:14
' Author    : Joacim Andersson, Brixoft Software, http://www.brixoft.net
' Purpose   : Used as a cool replacement for a regular label.
'             Supports different font formats, colors, and hyperlinks
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Function ShellExecute _
 Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Private Declare Function GetSysColor Lib "user32.dll" ( _
    ByVal nIndex As Long _
) As Long

Public Enum eBorderStyle
    None = 0
    [Fixed Single] = 1
End Enum

Private Type tStyle
    Color As Long
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
End Type

Private m_Caption As String
Private m_sPlainCaption As String
Private m_blnAutoNavigate As Boolean
Private m_blnWordWrap As Boolean
Private m_blnAutoSize As Boolean
Private m_Font As StdFont

Private nMaxWidth As Long
Private nMaxHeight As Long
Private blnInUrl As Boolean
Private blnInList As Boolean
Private nBulletIndent As Long
Private tOrig As tStyle
Private tCurrent As tStyle
Private colHyperlinks As Collection
Private hlCurrent As CHyperlink
Private blnRenderToContainer As Boolean
Private WithEvents oForm As Form
Attribute oForm.VB_VarHelpID = -1
Private WithEvents oPic As PictureBox
Attribute oPic.VB_VarHelpID = -1
Private oRender As Object

Public Event Click()
Public Event DblClick()
Public Event HyperlinkClick(ByVal URL As String)

Public Function ConvertWinColor(ByVal nColor As Long) As String
    'Is this a system color then convert it to the correct color
    If (nColor And &H80000000) = &H80000000 Then
        nColor = GetSysColor(nColor And &HFF&)
    End If
    'Convert the color to a hex string in the format #RRGGBB
    ConvertWinColor = "#" & Right$("0" & Hex$(nColor And &HFF&), 2) & _
     Right$("0" & Hex$((nColor And &HFF00&) \ &H100&), 2) & _
     Right$("0" & Hex$((nColor And &HFF0000) \ &H10000), 2)
End Function

Private Sub SetColor(ByVal sHex As String)
    Dim n As Long, nColor As Long
    If Left$(sHex, 1) = "#" Then
        sHex = Mid$(sHex, 2)
    End If
    If Len(sHex) <> 6 Then
        sHex = Right$("000000" & sHex, 6)
    End If
    For n = 1 To 6
        If InStr(1, "0123456789ABCDEF", Mid$(sHex, n, 1), vbTextCompare) = 0 Then
            'this is not a correct color, ignore it
            Exit Sub
        End If
    Next
    sHex = Right$(sHex, 6)
    UserControl.ForeColor = RGB(CInt("&H" & Left$(sHex, 2)), _
               CInt("&H" & Mid$(sHex, 3, 2)), _
               CInt("&H" & Right$(sHex, 2)))
    On Error Resume Next
    oRender.ForeColor = UserControl.ForeColor
End Sub

Private Sub SetStartTag(ByVal sTag As String)
    Dim s As String
    Dim sVal As String
    Dim oHL As CHyperlink
    Dim sFontName As String
    
    On Error Resume Next
    If blnInUrl Then
        'No new tags are allowed inside an [url] tag
        Exit Sub
    End If
    'remove the [ and ] characters
    s = LCase$(Mid$(sTag, 2, Len(sTag) - 2))
    'if an equal sign exist in the tag then split it up
    If InStr(s, "=") Then
        sTag = Trim$(Split(s, "=")(0))
        sVal = Trim$(Split(s, "=")(1))
    Else
        sTag = s
    End If
    'check the tag
    Select Case sTag
        Case "b"
            UserControl.Font.Bold = True
            oRender.Font.Bold = True
        Case "i"
            UserControl.Font.Italic = True
            oRender.Font.Italic = True
        Case "u"
            UserControl.Font.Underline = True
            oRender.Font.Underline = True
        Case "color"
            Select Case sVal
                Case "red"
                    UserControl.ForeColor = vbRed
                    oRender.ForeColor = vbRed
                Case "green"
                    UserControl.ForeColor = vbGreen
                    oRender.ForeColor = vbGreen
                Case "blue"
                    UserControl.ForeColor = vbBlue
                    oRender.ForeColor = vbBlue
                Case "white"
                    UserControl.ForeColor = vbWhite
                    oRender.ForeColor = vbWhite
                Case "black"
                    UserControl.ForeColor = vbBlack
                    oRender.ForeColor = vbBlack
                Case "yellow"
                    UserControl.ForeColor = vbYellow
                    oRender.ForeColor = vbYellow
                Case "magenta"
                    UserControl.ForeColor = vbMagenta
                    oRender.ForeColor = vbMagenta
                Case "cyan"
                    UserControl.ForeColor = vbCyan
                    oRender.ForeColor = vbCyan
                Case Else
                    Call SetColor(sVal)
            End Select
        Case "url"
            If Len(sVal) Then
                Set oHL = New CHyperlink
                With oHL
                    .HRef = sVal
                    .TextHeight = TextHeight("Xyz")
                    .X1 = CurrentX
                    .Y1 = CurrentY
                End With
                Call colHyperlinks.Add(oHL)
                blnInUrl = True
                With UserControl.Font
                    tCurrent.Bold = .Bold
                    tCurrent.Italic = .Italic
                    tCurrent.Underline = .Underline
                    .Underline = True
                    .Bold = False
                    .Italic = False
                End With
                With oRender.Font
                    .Underline = True
                    .Bold = False
                    .Italic = False
                End With
                tCurrent.Color = UserControl.ForeColor
                UserControl.ForeColor = vbBlue
                oRender.ForeColor = vbBlue
            End If
        Case "list"
            blnInList = True
            nBulletIndent = 0
            If Not (CurrentX = 0 And CurrentY = 0) Then
                Print
                oRender.Print
                oRender.CurrentX = Extender.Left
                m_sPlainCaption = m_sPlainCaption & vbCrLf
            End If
        Case "*"
            If blnInList Then
                If Not (CurrentX = 0 And CurrentY = 0) Then
                    Print
                    oRender.Print
                    oRender.CurrentX = Extender.Left
                    m_sPlainCaption = m_sPlainCaption & vbCrLf & "* "
                End If
                sFontName = UserControl.Font.Name
                UserControl.Font.Name = "Wingdings"
                Print " l";
                oRender.Font.Name = "Wingdings"
                oRender.Print " l";
                UserControl.Font.Name = sFontName
                Print " ";
                oRender.Font.Name = sFontName
                oRender.Print " ";
                nBulletIndent = CurrentX
            End If
        Case ""
            'empty tag = new line
            Print
            m_sPlainCaption = m_sPlainCaption & vbCrLf
    End Select
End Sub

Private Sub SetCloseTag(ByVal sTag As String)
    'Remove the [/ and ] parts
    sTag = LCase$(Mid$(sTag, 3, Len(sTag) - 3))
    On Error Resume Next
    'no other tags are allowed inside an [url] tag
    If sTag <> "url" And blnInUrl = True Then
        Exit Sub
    End If
    Select Case sTag
        Case "b"
            UserControl.Font.Bold = False
            oRender.Font.Bold = False
        Case "i"
            UserControl.Font.Italic = False
            oRender.Font.Italic = False
        Case "u"
            UserControl.Font.Underline = False
            oRender.Font.Underline = False
        Case "color"
            UserControl.ForeColor = tOrig.Color
            oRender.ForeColor = tOrig.Color
        Case "url"
            On Error Resume Next
            With colHyperlinks(colHyperlinks.Count)
                .X2 = CurrentX
                .Y2 = CurrentY
            End With
            With UserControl.Font
                .Bold = tCurrent.Bold
                .Italic = tCurrent.Italic
                .Underline = tCurrent.Underline
            End With
            With oRender.Font
                .Bold = tCurrent.Bold
                .Italic = tCurrent.Italic
                .Underline = tCurrent.Underline
            End With
            UserControl.ForeColor = tCurrent.Color
            oRender.ForeColor = tCurrent.Color
            blnInUrl = False
        Case "list"
            Print: Print
            oRender.Print: oRender.Print
            oRender.CurrentX = Extender.Left
            m_sPlainCaption = m_sPlainCaption & vbCrLf & vbCrLf
            blnInList = False
    End Select
End Sub

Private Sub WrapText(ByVal sText As String)
    Dim n As Long, nCount As Long
    Dim sChar As String, sWord As String
    
    m_sPlainCaption = m_sPlainCaption & sText
    If Not oRender Is Nothing Then
        oRender.CurrentX = Extender.Left + CurrentX
        oRender.CurrentY = Extender.Top + CurrentY
    End If
    If m_blnWordWrap = False Then
        Print sText;
        If blnRenderToContainer Then
            oRender.Print sText;
            'oRender.CurrentX = Extender.Left + CurrentX
        End If
        If CurrentX > nMaxWidth Then
            nMaxWidth = CurrentX
        End If
    Else
        sText = Replace(sText, vbCrLf, vbLf)
        nCount = Len(sText)
        If nCount Then
            For n = 1 To nCount
                sChar = Mid$(sText, n, 1)
                Select Case sChar
                    Case " ", vbLf, vbTab, "-"
                        If CurrentX + TextWidth(sWord) >= ScaleWidth Then
                            Print
                            If blnInList Then
                                CurrentX = nBulletIndent
                            End If
                            If blnRenderToContainer Then
                                oRender.Print
                                oRender.CurrentX = Extender.Left + CurrentX
                            End If
                        End If
                        If sChar = vbLf Then
                            Print sWord
                            If blnInList Then
                                CurrentX = nBulletIndent
                            End If
                            If blnRenderToContainer Then
                                oRender.Print sWord
                                oRender.CurrentX = Extender.Left + CurrentX
                            End If
                        Else
                            Print sWord; sChar;
                            If blnRenderToContainer Then
                                oRender.Print sWord; sChar;
                            End If
                        End If
                        sWord = ""
                    Case Else
                        sWord = sWord & sChar
                End Select
            Next
            If Len(sWord) Then
                If CurrentX + TextWidth(sWord) >= ScaleWidth Then
                    Print
                    If blnInList Then
                        CurrentX = nBulletIndent
                    End If
                    If blnRenderToContainer Then
                        oRender.Print
                        oRender.CurrentX = Extender.Left + CurrentX
                    End If
                End If
                Print sWord;
                If blnRenderToContainer Then
                    oRender.Print sWord;
                End If
            End If
        End If
    End If
End Sub

Private Sub DrawCaption()
    Dim s As String, s1 As String
    Dim sTag As String
    Dim nPos As Long, nPos2 As Long
    Dim nWidth As Long, nHeight As Long
    On Error Resume Next
    nMaxWidth = 0
    s = m_Caption
    m_sPlainCaption = ""
    Set colHyperlinks = New Collection
    Set hlCurrent = Nothing
    MousePointer = vbDefault
    oRender.MousePointer = vbDefault
    UserControl.Cls
    oRender.Cls
    With UserControl.Font
        .Bold = tOrig.Bold
        .Italic = tOrig.Italic
        .Underline = tOrig.Underline
    End With
    With oRender.Font
        .Bold = tOrig.Bold
        .Italic = tOrig.Italic
        .Underline = tOrig.Underline
    End With
    UserControl.ForeColor = tOrig.Color
    oRender.ForeColor = tOrig.Color
    Do
        nPos = InStr(1, s, "[", vbTextCompare)
        If nPos Then
            If nPos > 1 Then
                s1 = Left$(s, nPos - 1)
                Call WrapText(s1)
            End If
            nPos2 = InStr(nPos + 1, s, "]", vbTextCompare)
            If nPos2 Then
                sTag = Mid$(s, nPos, nPos2 - nPos + 1)
                If Mid$(sTag, 2, 1) = "/" Then
                    Call SetCloseTag(sTag)
                Else
                    Call SetStartTag(sTag)
                End If
                s = Mid$(s, nPos2 + 1)
            Else
                s = ""
            End If
        End If
    Loop While nPos
    If Len(s) Then
        Call WrapText(s)
    End If
    If m_blnAutoSize Then
        If m_blnWordWrap Then
            nWidth = UserControl.Width
        Else
            nWidth = nMaxWidth + ((2 + 4 * UserControl.BorderStyle) * _
             Screen.TwipsPerPixelX)
        End If
        nHeight = CurrentY + TextHeight("Xyz") + _
         ((2 + 4 * UserControl.BorderStyle) * Screen.TwipsPerPixelY)
        Call UserControl.Size(nWidth, nHeight)
    End If
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/Sets the text that is shown in the label. The Caption can contain formatting tags."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal sNew As String)
    m_Caption = sNew
    Call DrawCaption
    PropertyChanged "Caption"
End Property

Public Property Get PlainCaption() As String
    PlainCaption = m_sPlainCaption
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal nNew As OLE_COLOR)
    UserControl.BackColor = nNew
    DrawCaption
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = tOrig.Color
End Property

Public Property Let ForeColor(ByVal nNew As OLE_COLOR)
    UserControl.ForeColor = nNew
    tOrig.Color = nNew
    LSet tCurrent = tOrig
    DrawCaption
    PropertyChanged "ForeColor"
End Property

Public Property Let AutoNavigate(ByVal blnNew As Boolean)
    m_blnAutoNavigate = blnNew
    PropertyChanged "AutoNavigate"
End Property

Public Property Get AutoNavigate() As Boolean
    AutoNavigate = m_blnAutoNavigate
End Property

Public Property Get RenderToContainer() As Boolean
Attribute RenderToContainer.VB_MemberFlags = "400"
    RenderToContainer = blnRenderToContainer
End Property

Public Property Let RenderToContainer(ByVal blnNew As Boolean)
    Dim obj
    Dim frm As Form
    Dim pic As PictureBox
    If blnNew <> blnRenderToContainer Then
        If blnNew = True Then
            Set obj = Extender.Container
            If (TypeOf obj Is PictureBox) Then
                Set pic = Extender.Container
                Set oPic = pic
                Set oRender = oPic
                Set oForm = Nothing
            ElseIf (TypeOf obj Is Form) Then
                Set frm = Extender.Container
                Set oForm = frm
                Set oRender = oForm
                Set oPic = Nothing
            Else
                Set oForm = Nothing
                Set oRender = Nothing
                Set oPic = Nothing
                Exit Property
            End If
        Else
            Set oPic = Nothing
            Set oForm = Nothing
        End If
        blnRenderToContainer = blnNew
        PropertyChanged "RenderToContainer"
    End If
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal eNew As eBorderStyle)
    UserControl.BorderStyle = eNew
    PropertyChanged "BorderStyle"
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_blnWordWrap
End Property

Public Property Let WordWrap(ByVal blnNew As Boolean)
    If blnNew <> m_blnWordWrap Then
        m_blnWordWrap = blnNew
        DrawCaption
        PropertyChanged "WordWrap"
    End If
End Property

Property Get Font() As StdFont
    Set Font = m_Font
End Property

Property Set Font(ByVal oNew As StdFont)
    Set m_Font = oNew
    Set UserControl.Font = oNew
    DrawCaption
    PropertyChanged "Font"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_blnAutoSize
End Property

Public Property Let AutoSize(ByVal blnAutoSize As Boolean)
    If blnAutoSize <> m_blnAutoSize Then
        m_blnAutoSize = blnAutoSize
        DrawCaption
        PropertyChanged "AutoSize"
    End If
End Property

Private Sub oForm_Click()
    UserControl_Click
End Sub

Private Sub oForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set oForm.MouseIcon = UserControl.MouseIcon
    X = X - Extender.Left
    Y = Y - Extender.Top
    UserControl_MouseMove Button, Shift, X, Y
    oForm.MousePointer = MousePointer
End Sub

Private Sub oForm_Paint()
    DrawCaption
End Sub

Private Sub oPic_Click()
    UserControl_Click
End Sub

Private Sub oPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set oPic.MouseIcon = UserControl.MouseIcon
    X = X - Extender.Left
    Y = Y - Extender.Top
    UserControl_MouseMove Button, Shift, X, Y
    oPic.MousePointer = MousePointer
End Sub

Private Sub oPic_Paint()
    DrawCaption
End Sub

Private Sub UserControl_Click()
    Dim oHL As CHyperlink
    If Not hlCurrent Is Nothing Then 'the mouse is over a hyperlink
        If m_blnAutoNavigate Then
            Call ShellExecute(0&, "Open", hlCurrent.HRef, _
              vbNullString, vbNullString, vbNormalFocus)
            RaiseEvent Click
        Else
            RaiseEvent HyperlinkClick(hlCurrent.HRef)
        End If
    Else
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    m_blnWordWrap = True
    Set m_Font = Ambient.Font
    With tOrig
        .Bold = Ambient.Font.Bold
        .Italic = Ambient.Font.Italic
        .Underline = Ambient.Font.Underline
        .Color = Ambient.ForeColor
    End With
    LSet tCurrent = tOrig
    m_blnAutoNavigate = True
    Caption = Extender.Name
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oHL As CHyperlink
    On Error Resume Next
    If Not colHyperlinks Is Nothing Then
        For Each oHL In colHyperlinks
            If oHL.PosInside(X, Y) Then
                Set hlCurrent = oHL
                MousePointer = vbCustom
                Exit Sub
            End If
        Next
    End If
    MousePointer = vbDefault
    Set hlCurrent = Nothing
End Sub

Private Sub UserControl_Resize()
    DrawCaption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", eBorderStyle.None)
    m_blnAutoNavigate = PropBag.ReadProperty("AutoNavigate", True)
    m_blnAutoSize = PropBag.ReadProperty("AutoSize", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    tOrig.Color = UserControl.ForeColor
    m_blnWordWrap = PropBag.ReadProperty("WordWrap", True)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = m_Font
    Caption = PropBag.ReadProperty("Caption", "")
    RenderToContainer = PropBag.ReadProperty("RenderToContainer", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, eBorderStyle.None)
    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("AutoNavigate", m_blnAutoNavigate, True)
    Call PropBag.WriteProperty("AutoSize", m_blnAutoSize, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, Ambient.BackColor)
    Call PropBag.WriteProperty("ForeColor", tOrig.Color, Ambient.ForeColor)
    Call PropBag.WriteProperty("WordWrap", m_blnWordWrap, True)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("RenderToContainer", blnRenderToContainer, False)
End Sub

