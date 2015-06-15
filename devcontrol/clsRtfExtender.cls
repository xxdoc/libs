VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRtfExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Private Declare Function GetCaretBlinkTime Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Private Const MM_TWIPS = 6
Private Const WM_USER = &H400

Private Type Rect
    left As Long
    top As Long
    right As Long
    Bottom As Long
End Type

Private Type HighlightedLine
    startChar As Long
    endChar As Long
    lineNo As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Enum tmMsgs
        EM_CHARFROMPOS = &HD7
        EM_UNDO = &HC7
        EM_CANUNDO = &HC6
        EM_SETWORDBREAKPROC = &HD0
        EM_SETTABSTOPS = &HCB
        EM_SETSEL = &HB1
        EM_SETRECTNP = &HB4
        EM_SETRECT = &HB3
        EM_SETREADONLY = &HCF
        EM_SETPASSWORDCHAR = &HCC
        EM_SETMODIFY = &HB9
        EM_SCROLLCARET = &HB7
        EM_SETHANDLE = &HBC
        EM_SETCHARFORMAT = (WM_USER + 68)
        EM_SCROLL = &HB5
        EM_REPLACESEL = &HC2
        EM_LINESCROLL = &HB6
        EM_LINELENGTH = &HC1
        EM_LINEINDEX = &HBB
        EM_LINEFROMCHAR = &HC9
        EM_LIMITTEXT = &HC5
        EM_GETWORDBREAKPROC = &HD1
        EM_GETTHUMB = &HBE
        EM_GETRECT = &HB2
        EM_GETSEL = &HB0
        EM_GETPASSWORDCHAR = &HD2
        EM_GETMODIFY = &HB8
        EM_GETLINECOUNT = &HBA
        EM_GETLINE = &HC4
        EM_GETHANDLE = &HBD
        EM_GETFIRSTVISIBLELINE = &HCE
        EM_FMTLINES = &HC8
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
End Enum

Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As charFormatMasks 'Long
    dwEffects As charFormatEffects 'Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To 32 - 1) As Byte
    wPad2 As Integer
    'Additional stuff supported by RICHEDIT20
    wWeight As Integer     'Font weight (LOGFONT value)
    sSpacing As Integer    'Amount to space between letters
    crBackColor As Long    'Background color
    lLCID As Long          'Locale ID
    dwReserved As Long     'Reserved. Must be 0
    sStyle As Integer      'Style handle
    wKerning As Integer    'Twip size above which to kern char pair
    bUnderlineType As Byte 'Underline type
    bAnimation As Byte     'Animated text like marching ants
    bRevAuthor As Byte     'Revision author index
    bReserved1 As Byte
End Type

Private Const WM_VScroll = &H115
Private Const WM_CHAR = &H102
'Private Const CFE_BOLD = &H1
Private Const EC_LEFTMARGIN = &H1
Private Const EC_RIGHTMARGIN = &H2


Enum charFormatMasks
    CFM_BOLD = &H1&
    CFM_ITALIC = &H2&
    CFM_UNDERLINE = &H4&
    CFM_STRIKEOUT = &H8&
    CFM_PROTECTED = &H10&
    CFM_LINK = &H20&
    CFM_SIZE = &H80000000
    CFM_COLOR = &H40000000
    CFM_FACE = &H20000000
    CFM_OFFSET = &H10000000
    CFM_CHARSET = &H8000000
End Enum

Enum charFormatEffects
    CFE_BOLD = &H1&
    CFE_ITALIC = &H2&
    CFE_UNDERLINE = &H4&
    CFE_STRIKEOUT = &H8&
    CFE_PROTECTED = &H10&
    CFE_LINK = &H20&
    CFE_AUTOCOLOR = &H40000000
End Enum


'Private Const EC_LEFTMARGIN = &H1
'Private Const EC_RIGHTMARGIN = &H2
'Private Const LF_FACESIZE = 32
Private Const CFM_BACKCOLOR = &H4000000
Private Const SCF_SELECTION = &H1
Private Const WM_PASTE = &H302

Private myTopLine As Long
Private TrackingScroll As Boolean
Private OverRidingTabs As Boolean
Private OverrideTabNow As Boolean

Private hLine As HighlightedLine
Private LastMousePos As POINTAPI
Private myCurLine As Long

Private WithEvents SubClassMessage As clsSubClass
Attribute SubClassMessage.VB_VarHelpID = -1

Private WithEvents mRtf As RichTextBox
Attribute mRtf.VB_VarHelpID = -1

Public TabWidth As Integer
Public AutoIndent As Boolean

Public AddAccelerators As Boolean

Event Scrolled()
Event AutoComplete()
Event newLine(lineIndex As Long)
Event ArrowUpLine(prevlineIndex As Long)
Event ArrowDownLine(prevlineIndex As Long)
Event ClickedToNewLine(PrevLine As Long, curLine As Long)

Public isLocked As Boolean
Private LockCount As Integer
Private ww As Boolean

Sub SetRtf(rtf As Object)
    Set mRtf = rtf
End Sub

Function GetRtf() As Object
    Set GetRtf = mRtf
End Function


Public Property Let WordWrap(on_ As Boolean)
    ww = on_
    If on_ Then mRtf.RightMargin = 0 Else Call SetRightMargain
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = ww
End Property

Sub LockUpdate()
    If LockCount = 0 Then LockWindowUpdate mRtf.hwnd
    LockCount = LockCount + 1
    isLocked = True
End Sub

Sub UnlockUpdate()
    LockCount = LockCount - 1
    If LockCount = 0 Then
        LockWindowUpdate 0
        isLocked = False
    End If
End Sub

Function IndexOfFirstCharOnLine(lNo As Long)
    IndexOfFirstCharOnLine = SendMsg(EM_LINEINDEX, lNo)
End Function

Function CharIndexUnderMouse() As Long
    CharIndexUnderMouse = SendMessage(mRtf.hwnd, EM_CHARFROMPOS, 0&, LastMousePos)
End Function

Function CharUnderMouse() As String
    Dim pos As Long, lNo As Long, lStartIndex As Long, myLine As String
    'we go through some extra stuff only getting one line
    'of text for speed
    
    pos = CharIndexUnderMouse
    If pos < 1 Or pos > Len(mRtf.Text) Then Exit Function
    lNo = LineCharIsOn(pos)
    lStartIndex = Me.IndexOfFirstCharOnLine(lNo)
    myLine = GetLine(lNo)
    
    'pos is absolute, we need relative now
    pos = pos - lStartIndex
    If pos < 1 Then pos = 1
    CharUnderMouse = Mid(myLine, pos, 1)
    
End Function

Function WordUnderMouse() As String
     Dim pos As Long, lNo As Long, lStartIndex As Long, myLine As String
    'we go through some extra stuff only getting one line
    'of text for speed
    
    pos = CharIndexUnderMouse
    If pos < 1 Or pos > Len(mRtf.Text) Then Exit Function
    lNo = LineCharIsOn(pos)
    lStartIndex = Me.IndexOfFirstCharOnLine(lNo)
    myLine = GetLine(lNo)
    
    'pos is absolute, we need relative now
    pos = pos - lStartIndex
    If pos < 1 Then pos = 1
    
    Dim fs As Long, ns As Long, curChar As String, l As Long
    
    curChar = Mid(myLine, pos, 1)
    If curChar = " " Then
        fs = pos
        ns = InStr(fs + 1, myLine, " ")
        fs = fs + 1
    Else
        fs = InStrRev(myLine, " ", pos)
        ns = InStr(fs + 1, myLine, " ")
        fs = fs + 1
    End If

    If ns < 1 Then
        ns = Len(myLine) - fs - 1 '(2 = len(vbcrlf)
    Else
        ns = (ns - fs)
        If ns < 1 Then ns = 1
    End If
    
    WordUnderMouse = Mid(myLine, fs, ns)
    
End Function

Function LineCharIsOn(chrPos As Long) As Long
    LineCharIsOn = SendMsg(EM_LINEFROMCHAR, chrPos, 0)
End Function

Sub HighLightSelection(Optional bgColor = vbYellow, Optional fgColor = vbBlack)
    Dim udtCharFormat As CHARFORMAT2
    With udtCharFormat
        .cbSize = Len(udtCharFormat)
        .dwMask = CFM_BACKCOLOR Or CFM_COLOR
        .crBackColor = bgColor
        .crTextColor = fgColor
    End With
    SendMessage mRtf.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, udtCharFormat
End Sub

Sub HighLightLine(lNo As Long, Optional bgColor = vbYellow, Optional fgColor = vbBlack)
    LockUpdate
    SelectLine lNo
    HighLightSelection bgColor, fgColor
    mRtf.selStart = IndexOfFirstCharOnLine(lNo) + Len(GetLine(lNo)) - 2 '2= Len(vbCrLf)
    mRtf.selLength = 0
    mRtf.SelColor = vbBlack
    UnlockUpdate
End Sub

Sub SelectLine(lNo As Long)
    lNo = lNo - 1 'use 1 based index instead of 0 based
    hLine.lineNo = lNo
    hLine.startChar = SendMsg(EM_LINEINDEX, lNo)
    hLine.endChar = SendMsg(EM_LINEINDEX, lNo + 1) - Len(vbCrLf)
    mRtf.selStart = hLine.startChar
    If hLine.endChar < 0 Then hLine.endChar = Len(mRtf.Text)
    mRtf.selLength = hLine.endChar - hLine.startChar
    ScrollToLine CInt(lNo)
End Sub


Friend Function SendMsg(Msg As tmMsgs, Optional wParam As Long = 0, Optional lParam = 0) As Long
    SendMsg = SendMessage(mRtf.hwnd, Msg, wParam, lParam)
End Function

Property Get IsDirty() As Boolean
    IsDirty = CBool(SendMessage(mRtf.hwnd, EM_GETMODIFY, 0, 0))
End Property
Property Let IsDirty(s As Boolean)
        SendMessage mRtf.hwnd, EM_SETMODIFY, IIf(s, 1, 0), 0
End Property


Private Sub Class_Initialize()
    Set SubClassMessage = New clsSubClass
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'SubClassMessage.DetatchMessage mRtf.hwnd, WM_VScroll
    'SubClassMessage.DetatchMessage mRtf.hwnd, WM_CHAR
    If TrackingScroll Then TrackScrolling = False
    If OverRidingTabs Then OverrideTabs = False
End Sub

Property Let TrackScrolling(trackit As Boolean)
    On Error Resume Next ' weird sporatic tear down bug
    
    If trackit Then
        If TrackingScroll Then Exit Property
        SubClassMessage.AttachMessage mRtf.hwnd, WM_VScroll
        TrackingScroll = True
    Else
        If TrackingScroll Then
            SubClassMessage.DetatchMessage mRtf.hwnd, WM_VScroll
            TrackingScroll = False
        End If
    End If
    
End Property

Property Let OverrideTabs(overRide As Boolean)
    On Error Resume Next ' weird sporatic tear down bug
    Dim r As Boolean
    
    If overRide Then
        If OverRidingTabs Then Exit Property
        r = SubClassMessage.AttachMessage(mRtf.hwnd, WM_CHAR)
        If Not r Then MsgBox SubClassMessage.ErrorMessage
        OverRidingTabs = True
    Else
        If OverRidingTabs Then
            SubClassMessage.DetatchMessage mRtf.hwnd, WM_CHAR
        End If
    End If
End Property

Sub SetMargins(left As Integer, Optional right As Integer = -1)
     SendMessage mRtf.hwnd, EM_SETMARGINS, EC_LEFTMARGIN, ByVal MakeLong(0, left)
     If right > -1 Then
        SendMessage mRtf.hwnd, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal MakeLong(right, 0)
     End If
End Sub


Function CurrentColumn() As Long
       CurrentColumn = mRtf.selStart - SendMessage(mRtf.hwnd, EM_LINEINDEX, -1&, ByVal 0&) + 1
End Function

Function CurrrentLineIndex() As Long
    CurrrentLineIndex = SendMessage(mRtf.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
End Function

Function GetLine(lineNo As Long) As String
    'When retrieving a line into a buffer the first word
    'of the buffer specifies the max number of bytes to read,
    'so one has to guess the maximum line length.  When the bytes
    'are read into the buffer the first word gets overwritten.
    'Remember that lines in a TextBox are numbered starting at zero.

    Dim lret As Long
    Dim strBuffer As String
    Dim intLineLen As Integer
    Const MAX_LINE_LEN = 300
    
    strBuffer = Space(MAX_LINE_LEN) 'max width possible for line
    intLineLen = SendMessageStr(mRtf.hwnd, EM_GETLINE, lineNo, strBuffer)
    GetLine = left(strBuffer, intLineLen)


End Function


Property Get lineCount() As Long
    lineCount = SendMessage(mRtf.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
End Property

Property Get TopLineIndex() As Long
    TopLineIndex = SendMessage(mRtf.hwnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Property

Property Get VisibleLines() As Long
    Dim udtRect As Rect, tm As TEXTMETRIC
    Dim hdc As Long, lFont As Long, lOrgFont As Long
    Const WM_GETFONT As Long = &H31
    
    SendMessage mRtf.hwnd, EM_GETRECT, 0, udtRect

    lFont = SendMessage(mRtf.hwnd, WM_GETFONT, 0, 0)
    hdc = GetDC(mRtf.hwnd)

    If lFont <> 0 Then
        lOrgFont = SelectObject(hdc, lFont)
    End If

    GetTextMetrics hdc, tm
    
    If lFont <> 0 Then
        lFont = SelectObject(hdc, lOrgFont)
    End If

    VisibleLines = (udtRect.Bottom - udtRect.top) \ tm.tmHeight

    ReleaseDC mRtf.hwnd, hdc

End Property

Friend Function CaretPos() As POINTAPI
    Dim pos As Long, caret As Long
    Dim p As POINTAPI
    Const EM_POSFROMCHAR = &HD6
    
    caret = mRtf.selStart - 1
    pos = SendMessage(mRtf.hwnd, EM_POSFROMCHAR, caret, 0)
    
    'pos is really a short of the x and a short of the y, packed into the
    'x is the lower 2 bytes, y is the upper 2 bytes.
    
    CaretPos.x = mRtf.left + (LoWord(pos) * 15)
    CaretPos.y = mRtf.top + (HiWord(pos) * 15)
    
End Function

Sub ScrollToLine(x As Long)
     x = x - TopLineIndex
     ScrollIncremental , x
End Sub

Sub ScrollIncremental(Optional horz As Long = 0, Optional vert As Long = 0)
    'lParam&  The low-order 2 bytes specify the number of vertical
    '          lines to scroll. The high-order 2 bytes specify the
    '          number of horizontal columns to scroll. A positive
    '          value for lParam& causes text to scroll upward or to the
    '          left. A negative value causes text to scroll downward or
    '          to the right.
    ' r&       Indicates the number of lines actually scrolled.
    
    Dim r As Long
    r = CLng(&H10000 * horz) + vert
    r = SendMessage(mRtf.hwnd, EM_LINESCROLL, 0, ByVal r)
    RaiseEvent Scrolled

End Sub




Private Sub mRtf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    With LastMousePos
         .x = x \ Screen.TwipsPerPixelX
         .y = y \ Screen.TwipsPerPixelY
    End With
End Sub

Private Sub mRtf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    Dim newCurLine As Long, tLine As Long
    tLine = TopLineIndex
    newCurLine = CurrrentLineIndex - 1
    
    If hLine.endChar > 0 Then
        LockUpdate
        HighLightLine hLine.lineNo + 1, vbWhite, vbBlack
        hLine.endChar = 0: hLine.startChar = 0: hLine.lineNo = 0
        mRtf.selStart = CharIndexUnderMouse
        UnlockUpdate
    End If
  
    'ignore block selections!
    If newCurLine <> myCurLine And mRtf.selLength = 0 Then
        LockUpdate
        RaiseEvent ClickedToNewLine(myCurLine, newCurLine)
        myCurLine = newCurLine
        ScrollToLine tLine
        mRtf.selStart = CharIndexUnderMouse
        UnlockUpdate
    End If
    
End Sub

Private Sub SubClassMessage_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    Select Case wMsg
        Case WM_VScroll
            'now make sure we are not between line even a little cause it syucks and happens!
            'we dont do it when they are dragging the scrollbar through cause it
            'makes it jump real bad
            
            RaiseEvent Scrolled
            If Not LMouseDown Then ScrollToLine TopLineIndex 'has to be like this :0-

        Case WM_CHAR
            If wParam = 9 Then
                Cancel = True
                If TabWidth = 0 Then TabWidth = 5
                If mRtf.selLength > 0 Then
                    If GetAsyncKeyState(vbKeyShift) Then 'shift tab - moves sel text left
                        mRtf.selText = Replace(mRtf.selText, Space(TabWidth), "")
                    Else
                        mRtf.selText = Space(TabWidth) & Replace(mRtf.selText, vbCrLf, vbCrLf & Space(TabWidth))
                    End If
                Else
                    mRtf.selText = Space(TabWidth)
                End If
            End If
            
'            Debug.Print
    End Select
    
    
End Sub
 
Private Sub mRtf_KeyDown(KeyCode As Integer, Shift As Integer)
    If AddAccelerators Then
        If KeyCode = 32 And Shift = 2 Then 'autocomplete ctrl-space
            RaiseEvent AutoComplete
            KeyCode = 0
        End If
    End If
End Sub

Private Sub mRtf_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tli As Long
    Dim lstLine As String, spacer As String
    
    tli = TopLineIndex
    myCurLine = CurrrentLineIndex - 1
    
    If tli <> myTopLine Then  'textbox has auto scrolled on us
        myTopLine = tli
        RaiseEvent Scrolled
        Exit Sub
    End If
   
    If KeyCode = 13 Then
        RaiseEvent newLine(CurrrentLineIndex - 1)
    End If
    
    If KeyCode = 38 Then
        RaiseEvent ArrowUpLine(CurrrentLineIndex + 1)
    End If
    
    If KeyCode = 40 Then
        RaiseEvent ArrowDownLine(CurrrentLineIndex - 1)
    End If
    
    If KeyCode = 13 And AutoIndent Then
        LockUpdate
        tli = CurrrentLineIndex
        lstLine = GetLine(tli - 2)
        While left(lstLine, 1) = " " Or left(lstLine, 1) = vbTab
            spacer = spacer & left(lstLine, 1)
            lstLine = Mid(lstLine, 2, Len(lstLine))
        Wend
        mRtf.selText = spacer
        UnlockUpdate
    End If
    
    If AddAccelerators Then
    
        If KeyCode = 65 And Shift = 2 Then 'select all ctrl-a
            mRtf.selStart = 0
            mRtf.selLength = Len(mRtf.Text)
        End If
        
    End If
        
End Sub

Function charHeight() As Long
    Dim tm As TEXTMETRIC
    Dim hdc As Long, PrevMapMode
    
    Const MM_TWIPS = 6
    
    hdc = GetWindowDC(mRtf.hwnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC mRtf.hwnd, hdc
    End If
    
    charHeight = tm.tmHeight

End Function

Function MaxCharWidth() As Long
    Dim tm As TEXTMETRIC
    Dim hdc As Long, PrevMapMode
    
    Const MM_TWIPS = 6
    
    hdc = GetWindowDC(mRtf.hwnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC mRtf.hwnd, hdc
    End If
    
    MaxCharWidth = tm.tmAveCharWidth

End Function

Function CharBeforeCursor() As Byte
    Dim lastCharByPos As Byte
    Dim curLine As String
    Dim LineCharIndex As Long
    Dim curLineIndex As Long
    Dim charPosOnLine As Long
    
    'On Error Resume Next
    curLineIndex = CurrrentLineIndex - 1
    curLine = GetLine(curLineIndex)
    LineCharIndex = IndexOfFirstCharOnLine(curLineIndex)
    charPosOnLine = mRtf.selStart - LineCharIndex
    lastCharByPos = Asc(Mid(curLine, charPosOnLine, 1))
    CharBeforeCursor = lastCharByPos
End Function

Function WordBeforeCursor(indexFound As Integer, ParamArray dividers()) As String
    Dim lastWordByPos As String
    Dim curLine As String
    Dim LineCharIndex As Long
    Dim curLineIndex As Long
    Dim charPosOnLine As Long
    Dim tmpIndex As Long, BoundryIndex As Long, i As Long
    
    indexFound = -1
    
    curLineIndex = CurrrentLineIndex - 1
    curLine = GetLine(curLineIndex)
    LineCharIndex = IndexOfFirstCharOnLine(curLineIndex)
    charPosOnLine = mRtf.selStart - LineCharIndex
    
    If charPosOnLine = 0 Then Exit Function
    
    For i = 0 To UBound(dividers)
        tmpIndex = InStrRev(curLine, dividers(i), charPosOnLine)
        'find the farthest right divider (closest to cursor)
        If tmpIndex > BoundryIndex Then
            BoundryIndex = tmpIndex
            indexFound = i 'return which divider we use
        End If
    Next
    
    If charPosOnLine < 2 Then Exit Function
    
    lastWordByPos = Mid(curLine, BoundryIndex + 1, charPosOnLine - BoundryIndex)
    WordBeforeCursor = lastWordByPos
    
End Function


Private Sub SetRightMargain()
    Dim tm As TEXTMETRIC
    Dim longestLine As Long, hdc As Long, i As Long
    Dim lineCount As Long, lineLength As Long, lineIndex As Long
    Dim PrevMapMode
    
    lineCount = SendMessage(mRtf.hwnd, EM_GETLINECOUNT, 0&, 0&)
    
    For i = 0 To lineCount - 1
        lineIndex = SendMessage(mRtf.hwnd, EM_LINEINDEX, i, 0&)
        lineLength = SendMessage(mRtf.hwnd, EM_LINELENGTH, lineIndex, 0&)
        If lineLength > longestLine Then longestLine = lineLength
    Next
        
    hdc = GetWindowDC(mRtf.hwnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC mRtf.hwnd, hdc
    End If
    
    mRtf.RightMargin = longestLine * tm.tmMaxCharWidth
    
End Sub


