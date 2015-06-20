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

Implements ICSubclass

'David Zimmer <dzzie@yahoo.com> 2001
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

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
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Const MM_TWIPS = 6
Private Const WM_USER = &H400

Private Type Rect
    left As Long
    top As Long
    right As Long
    Bottom As Long
End Type


Private Type LongWords
    LoWord As String
    HiWord As String
End Type

Private Type HighlightedLine
    startChar As Long
    endChar As Long
    lineNo As Long
End Type

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECHARRANGE
    Min As Long
    Max As Long
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
        EC_LEFTMARGIN = &H1
        EC_RIGHTMARGIN = &H2
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
        EM_FINDWORDBREAK = WM_USER + 76
        EM_EMPTYUNDOBUFFER = &HCD
        EM_SETMARGINS = &HD3
        EM_SETTARGETDEVICE = (WM_USER + 72)
        EM_EXGETSEL = (WM_USER + 52)
        EM_EXSETSEL = (WM_USER + 55)
        EM_GETTEXTLENGTHEX = (WM_USER + 95)
        EM_GETSELTEXT = (WM_USER + 62)
        EM_EXLINEFROMCHAR = (WM_USER + 54)
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
Private Const WM_MBUTTONUP = &H208


Private Const LF_FACESIZE = 32
 
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
    CFM_BACKCOLOR = &H4000000
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

Private Type REGETTEXTLENGTHEX
    Flags As Long
    CodePage As Long
End Type

Public Enum WordBreakOptions
      WB_LEFT = 0 'Finds the beginning of a word to the left of the specified position.</summary>
      WB_RIGHT = 1      'Finds the beginning of a word to the right of the specified position. This is useful in right-aligned edit controls.</summary>
      WB_ISDELIMITER = 2     'Returns TRUE if the character at the specified position is a delimiter, or FALSE otherwise.</summary>
      WB_CLASSIFY = 3      'Retrieves the character class and word break flags of the character at the specified position. This value is for use with rich edit controls.</summary>
      WB_MOVEWORDLEFT = 4      'Finds the beginning of a word to the left of the specified position. This value is used during CTRL+LEFT key processing. This value is for use with rich edit controls.</summary>
      WB_MOVEWORDRIGHT = 5     'Finds the beginning of a word to the right of the specified position. This value is used during CTRL+RIGHT key processing. This value is for use with rich edit controls.</summary>
      WB_LEFTBREAK = 6   'Finds the end-of-word delimiter to the left of the specified position. This value is for use with rich edit controls.</summary>
      WB_RIGHTBREAK = 7 'Finds the beginning of a word to the right of the specified position. This is useful in right-aligned edit controls.</summary>
End Enum
   

'Private Const EC_LEFTMARGIN = &H1
'Private Const EC_RIGHTMARGIN = &H2
'Private Const LF_FACESIZE = 32
'Private Const CFM_BACKCOLOR = &H4000000
Private Const SCF_SELECTION = &H1
Private Const WM_PASTE = &H302
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const GTL_DEFAULT As Long = 0
Private Const GTL_USECRLF As Long = 1
Private Const GTL_PRECISE As Long = 2
Private Const GTL_CLOSE As Long = 4
Private Const GTL_NUMCHARS As Long = 8
Private Const GTL_NUMBYTES As Long = 16
Private Const CP_UNICODE As Long = 1200

Private myTopLine As Long
Private TrackingScroll As Boolean
Private OverRidingTabs As Boolean
Private OverrideTabNow As Boolean

Private hLine As HighlightedLine
Private LastMousePos As POINTAPI
Private myCurLine As Long

Private SubClass As cSubclass
Attribute SubClass.VB_VarHelpID = -1


Public TabWidth As Integer
Public AutoIndent As Boolean


Event Scrolled()
'Event AutoComplete()
Event newLine(lineIndex As Long)
Event ArrowUpLine(prevlineIndex As Long)
Event ArrowDownLine(prevlineIndex As Long)
'Event ClickedToNewLine(PrevLine As Long, curLine As Long)

Public isLocked As Boolean
Private LockCount As Integer
Private ww As Boolean
Private hwnd As Long

Sub InitRtf(m_hwnd As Long)
    
    If IsWindow(m_hwnd) = 0 Then Exit Sub
    
    If SubClass Is Nothing Then
        Set SubClass = New cSubclass
        hwnd = m_hwnd
        SubClass.SubClass hwnd, Me
        'SubClass.AddMsg hwnd, WM_KEYDOWN
        SubClass.AddMsg hwnd, WM_KEYUP
        SubClass.AddMsg hwnd, WM_MOUSEMOVE
        SubClass.AddMsg hwnd, WM_VScroll
        'SubClass.AddMsg hwnd, WM_MBUTTONUP
        'SubClass.AddMsg hwnd, WM_CHAR, MSG_BEFORE
    End If
    
End Sub

Public Function LMouseDown() As Boolean
    GetAsyncKeyState vbKeyLButton
    LMouseDown = Not (GetAsyncKeyState(vbKeyLButton) And &HFFFF) = 0
End Function

Public Function RMouseDown() As Boolean
    GetAsyncKeyState vbKeyRButton
    RMouseDown = Not (GetAsyncKeyState(vbKeyRButton) And &HFFFF) = 0
End Function

Private Sub Class_Terminate()
    If Not SubClass Is Nothing Then
        SubClass.Unsubclass hwnd
    End If
End Sub

 


Private Property Get SelStart() As Long
    If hwnd <> 0 Then
        Dim RECR As RECHARRANGE
        SendMessage hwnd, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
        SelStart = RECR.Min
    End If
End Property

Private Property Let SelStart(ByVal Value As Long)
    If hwnd <> 0 Then
        If Value >= 0 Then
            Dim RECR As RECHARRANGE
            RECR.Min = Value
            RECR.Max = Value
            SendMessage hwnd, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
        Else
            Err.Raise 380
        End If
    End If
End Property

Private Property Get SelLength() As Long
    If hwnd <> 0 Then
        Dim RECR As RECHARRANGE
        SendMessage hwnd, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
        SelLength = RECR.Max - RECR.Min
    End If
End Property

Private Property Let SelLength(ByVal Value As Long)
    If hwnd <> 0 Then
        If Value >= 0 Then
            Dim RECR As RECHARRANGE
            SendMessage hwnd, EM_EXGETSEL, 0, ByVal VarPtr(RECR)
            RECR.Max = RECR.Min + Value
            SendMessage hwnd, EM_EXSETSEL, 0, ByVal VarPtr(RECR)
        Else
            Err.Raise 380
        End If
    End If
End Property

Private Property Get SelText() As String
    If hwnd <> 0 Then
        Dim REGTLEX As REGETTEXTLENGTHEX, length As Long
        REGTLEX.Flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
        REGTLEX.CodePage = CP_UNICODE
        length = SendMessage(hwnd, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
        If length > 0 Then
            Dim Buffer As String
            Buffer = String(length, vbNullChar)
            length = SendMessageW(hwnd, EM_GETSELTEXT, 0, ByVal StrPtr(Buffer))
            If length > 0 Then SelText = left$(Buffer, length)
        End If
    End If
End Property

Private Property Let SelText(ByVal Value As String)
    If hwnd <> 0 Then SendMessageW hwnd, EM_REPLACESEL, 0, ByVal StrPtr(Value)
End Property



'Function FindWordBreak(ByVal opt As WordBreakOptions, ByVal position As Long) As Long
'      FindWordBreak = SendMessage(hwnd, EM_FINDWORDBREAK, opt, position)
'End Function
'
'Function IsDelimiter(ByVal position As Long) As Boolean
'      IsDelimiter = CBool(SendMessage(hwnd, EM_FINDWORDBREAK, WB_ISDELIMITER, position))
'End Function


Sub LockUpdate()
    If LockCount = 0 Then LockWindowUpdate hwnd
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

Private Function IndexOfFirstCharOnLine(lNo As Long)
    IndexOfFirstCharOnLine = SendMsg(EM_LINEINDEX, lNo)
End Function

Function CharIndexUnderMouse() As Long
    CharIndexUnderMouse = SendMessage(hwnd, EM_CHARFROMPOS, 0&, LastMousePos)
End Function

Function CharUnderMouse() As String
    Dim pos As Long, lNo As Long, lStartIndex As Long, myLine As String
    'we go through some extra stuff only getting one line
    'of text for speed
    
    pos = CharIndexUnderMouse
    If pos < 1 Or pos > Me.TextLength Then Exit Function
    lNo = LineCharIsOn(pos)
    lStartIndex = IndexOfFirstCharOnLine(lNo)
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
    If pos < 1 Or pos > Me.TextLength Then Exit Function
    lNo = LineCharIsOn(pos)
    lStartIndex = IndexOfFirstCharOnLine(lNo)
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
    
    If ns > 0 And ns > fs And fs > 0 Then
        WordUnderMouse = Mid(myLine, fs, ns)
    End If
    
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
    SendMessage hwnd, EM_SETCHARFORMAT, SCF_SELECTION, udtCharFormat
End Sub

Sub HighLightLine(lNo As Long, Optional bgColor = vbYellow, Optional fgColor = vbBlack)
    LockUpdate
    SelectLine lNo
    HighLightSelection bgColor, fgColor
    SelStart = IndexOfFirstCharOnLine(lNo) + Len(GetLine(lNo)) - 2 '2= Len(vbCrLf)
    SelLength = 0
    UnlockUpdate
End Sub

Sub SelectLine(lNo As Long)
    lNo = lNo - 1 'use 1 based index instead of 0 based
    hLine.lineNo = lNo
    hLine.startChar = SendMsg(EM_LINEINDEX, lNo)
    hLine.endChar = SendMsg(EM_LINEINDEX, lNo + 1) - Len(vbCrLf)
    SelStart = hLine.startChar
    If hLine.endChar < 0 Then hLine.endChar = Me.TextLength
    SelLength = hLine.endChar - hLine.startChar
    ScrollToLine CInt(lNo)
End Sub


Friend Function SendMsg(Msg As tmMsgs, Optional wParam As Long = 0, Optional lParam = 0) As Long
    SendMsg = SendMessage(hwnd, Msg, wParam, lParam)
End Function

Property Get IsDirty() As Boolean
    IsDirty = CBool(SendMessage(hwnd, EM_GETMODIFY, 0, 0))
End Property
Property Let IsDirty(s As Boolean)
        SendMessage hwnd, EM_SETMODIFY, IIf(s, 1, 0), 0
End Property


 



'Sub SetMargins(left As Integer, Optional right As Integer = -1)
'     SendMessage hwnd, EM_SETMARGINS, EC_LEFTMARGIN, ByVal MakeLong(0, left)
'     If right > -1 Then
'        SendMessage hwnd, EM_SETMARGINS, EC_RIGHTMARGIN, ByVal MakeLong(right, 0)
'     End If
'End Sub


Function CurrentColumn() As Long
       CurrentColumn = SelStart - SendMessage(hwnd, EM_LINEINDEX, -1&, ByVal 0&) + 1
End Function

Function CurrrentLineIndex() As Long
    CurrrentLineIndex = SendMessage(hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1
End Function

'Function GetLine(lineNo As Long) As String
'    'When retrieving a line into a buffer the first word
'    'of the buffer specifies the max number of bytes to read,
'    'so one has to guess the maximum line length.  When the bytes
'    'are read into the buffer the first word gets overwritten.
'    'Remember that lines in a TextBox are numbered starting at zero.
'
'    Dim lret As Long
'    Dim strBuffer As String
'    Dim intLineLen As Integer
'    Const MAX_LINE_LEN = 300
'
'    strBuffer = Space(MAX_LINE_LEN) 'max width possible for line
'    intLineLen = SendMessageStr(hwnd, EM_GETLINE, lineNo, strBuffer)
'    GetLine = left(strBuffer, intLineLen)
'
'
'End Function

Public Function GetLine(ByVal LineNumber As Long) As String
If LineNumber < 0 Then Err.Raise 380
If hwnd <> 0 Then
    Dim FirstCharPos As Long, length As Long
    FirstCharPos = SendMessage(hwnd, EM_LINEINDEX, LineNumber, ByVal 0&)
    If FirstCharPos > -1 Then
        length = SendMessage(hwnd, EM_LINELENGTH, FirstCharPos, ByVal 0&)
        If length > 0 Then
            Dim Buffer As String
            Buffer = ChrW(length) & String(length - 1, vbNullChar)
            If LineNumber > 0 Then
                If SendMessageW(hwnd, EM_GETLINE, LineNumber, ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            Else
                If SendMessageW(hwnd, EM_GETLINE, SendMessage(hwnd, EM_EXLINEFROMCHAR, 0, ByVal FirstCharPos), ByVal StrPtr(Buffer)) > 0 Then GetLine = Buffer
            End If
        End If
    Else
        Err.Raise 380
    End If
End If
End Function


Property Get lineCount() As Long
    lineCount = SendMessage(hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
End Property

Property Get TopLineIndex() As Long
    TopLineIndex = SendMessage(hwnd, EM_GETFIRSTVISIBLELINE, 0, ByVal 0&)
End Property

Property Get VisibleLines() As Long
    Dim udtRect As Rect, tm As TEXTMETRIC
    Dim hdc As Long, lFont As Long, lOrgFont As Long
    Const WM_GETFONT As Long = &H31
    
    SendMessage hwnd, EM_GETRECT, 0, udtRect

    lFont = SendMessage(hwnd, WM_GETFONT, 0, 0)
    hdc = GetDC(hwnd)

    If lFont <> 0 Then
        lOrgFont = SelectObject(hdc, lFont)
    End If

    GetTextMetrics hdc, tm
    
    If lFont <> 0 Then
        lFont = SelectObject(hdc, lOrgFont)
    End If

    VisibleLines = (udtRect.Bottom - udtRect.top) \ tm.tmHeight

    ReleaseDC hwnd, hdc

End Property

'Friend Function CaretPos() As POINTAPI
'    Dim pos As Long, caret As Long
'    Dim p As POINTAPI
'    Const EM_POSFROMCHAR = &HD6
'
'    caret = selstart - 1
'    pos = SendMessage(hwnd, EM_POSFROMCHAR, caret, 0)
'
'    'pos is really a short of the x and a short of the y, packed into the
'    'x is the lower 2 bytes, y is the upper 2 bytes.
'
'    CaretPos.x = mRtf.left + (LoWord(pos) * 15)
'    CaretPos.y = mRtf.top + (HiWord(pos) * 15)
'
'End Function

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
    r = SendMessage(hwnd, EM_LINESCROLL, 0, ByVal r)
    RaiseEvent Scrolled

End Sub

'Private Sub mRtf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    Dim newCurLine As Long, tLine As Long
'    tLine = TopLineIndex
'    newCurLine = CurrrentLineIndex - 1
'
'    If hLine.endChar > 0 Then
'        LockUpdate
'        HighLightLine hLine.lineNo + 1, vbWhite, vbBlack
'        hLine.endChar = 0: hLine.startChar = 0: hLine.lineNo = 0
'        selstart = CharIndexUnderMouse
'        UnlockUpdate
'    End If
'
'    'ignore block selections!
'    If newCurLine <> myCurLine And sellength = 0 Then
'        LockUpdate
'        RaiseEvent ClickedToNewLine(myCurLine, newCurLine)
'        myCurLine = newCurLine
'        ScrollToLine tLine
'        selstart = CharIndexUnderMouse
'        UnlockUpdate
'    End If
'
'End Sub

Private Sub ICSubclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

On Error Resume Next
Dim x As Integer
Dim y As Integer

    Select Case wMsg
        Case WM_VScroll
            'now make sure we are not between line even a little cause it syucks and happens!
            'we dont do it when they are dragging the scrollbar through cause it
            'makes it jump real bad
            
            If myTopLine <> Me.TopLineIndex Then
                RaiseEvent Scrolled
                myTopLine = Me.TopLineIndex
            End If
            'If Not LMouseDown Then ScrollToLine TopLineIndex 'has to be like this :0-

 
        Case WM_MOUSEMOVE
        
                x = LoWord(lParam)
                y = HiWord(lParam)
                
                With LastMousePos
                     .x = x
                     .y = y
                End With
    
        Case WM_KEYUP

                mRtf_KeyUp CInt(wParam), 0
                
        Case WM_CHAR
        
            'replace tabs with spaces
            If wParam = 9 Then
                bHandled = True
                If TabWidth = 0 Then TabWidth = 5
                If SelLength > 0 Then
                    If GetAsyncKeyState(vbKeyShift) Then 'shift tab - moves sel text left
                        SelText = Replace(SelText, Space(TabWidth), "")
                    Else
                        SelText = Space(TabWidth) & Replace(SelText, vbCrLf, vbCrLf & Space(TabWidth))
                    End If
                Else
                    SelText = Space(TabWidth)
                End If
            End If
            
'            Debug.Print
    End Select
    
    
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
        SelText = spacer
        UnlockUpdate
    End If

        
End Sub

Function charHeight() As Long
    Dim tm As TEXTMETRIC
    Dim hdc As Long, PrevMapMode
    
    Const MM_TWIPS = 6
    
    hdc = GetWindowDC(hwnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC hwnd, hdc
    End If
    
    charHeight = tm.tmHeight

End Function

Function MaxCharWidth() As Long
    Dim tm As TEXTMETRIC
    Dim hdc As Long, PrevMapMode
    
    Const MM_TWIPS = 6
    
    hdc = GetWindowDC(hwnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC hwnd, hdc
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
    charPosOnLine = SelStart - LineCharIndex
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
    charPosOnLine = SelStart - LineCharIndex
    
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


 

Property Get TextLength() As Long
'    Dim lastLine As Long, sl As Long, el As Long
'    lastLine = Me.lineCount
'    sl = IndexOfFirstCharOnLine(lastLine - 1)
'    el = LineLength(lastLine - 1)
'    TextLength = sl + el

    Dim REGTLEX As REGETTEXTLENGTHEX
    TextLength = SendMessage(hwnd, EM_GETTEXTLENGTHEX, VarPtr(REGTLEX), ByVal 0&)
    
End Property
 

Function LineLength(Optional ByVal index As Long = -1) As Long
    If index = -1 Then index = SelStart
    LineLength = SendMessageLong(hwnd, EM_LINELENGTH, index, 0&)
End Function

Function LineStartPos(lineIndex As Long) As Long
    LineStartPos = SendMessageLong(hwnd, EM_LINEINDEX, lineIndex, 0&)
End Function

Function CurrentLine() As Long
    CurrentLine = SendMessageLong(hwnd, EM_LINEFROMCHAR, SelStart, 0&)
End Function

Sub SelBackcolor(lngColor As Long)
    Dim tCF2 As CHARFORMAT2
    tCF2.dwMask = CFM_BACKCOLOR
    tCF2.crBackColor = lngColor
    tCF2.cbSize = Len(tCF2)
    SendMessage hwnd, EM_SETCHARFORMAT, 1, tCF2
End Sub

'Function CurrentWord() As String
'
'    Dim ss As Long, sl As Long
'
'    Dim startAt As Long
'    Dim endAt As Long
'    Dim c As Integer, leng As Long
'    Dim cl As String
'
'    On Error Resume Next
'
'    ss = selstart
'    sl = sellength
'    Me.LockUpdate
'
'    startAt = ss
'
'    cl = Me.CurrentLine
'
'    Do While startAt > 1
'        'If IsDelimiter(startAt) Then Exit Do
'        selstart = startAt
'        sellength = 1
'        c = Asc(seltext)
'        Debug.Print Chr(c)
'        If isAlpha(c, True) Then
'            startAt = startAt - 1
'        Else
'            startAt = startAt + 1
'            Exit Do
'        End If
'    Loop
'
'    endAt = ss
'    Debug.Print "startat: " & startAt
'
'    Dim tl As Long
'    tl = Me.TextLength
'
'    Do While endAt < tl
'        'If IsDelimiter(endAt) Then Exit Do
'        selstart = endAt
'        sellength = 1
'        c = Asc(seltext)
'        Debug.Print Chr(c)
'        If isAlpha(c, True) Then
'            endAt = endAt + 1
'        Else
'            Exit Do
'        End If
'    Loop
'
'    selstart = startAt
'    leng = endAt - startAt
'    sellength = leng
'    CurrentWord = seltext
'
'
'    selstart = ss
'    sellength = sl
'
'    Me.UnlockUpdate
'
'
'End Function

Function isAlpha(c As Integer, Optional numericToo As Boolean = False) As Boolean

    If c >= Asc("a") And c <= Asc("z") Then
        isAlpha = True
        Exit Function
    End If
    
    If c >= Asc("A") And c <= Asc("Z") Then
        isAlpha = True
        Exit Function
    End If
    
    If numericToo Then
        If c >= Asc("0") And c <= Asc("9") Then
            isAlpha = True
            Exit Function
        End If
    End If
    
    If c = Asc("_") Then isAlpha = True
    
End Function

Private Function LongToWords(l As Long) As LongWords
    Dim w(3) As Byte
    CopyMemory w(0), l, 4
    LongToWords.HiWord = Hexit(w(3)) & Hexit(w(2))
    LongToWords.LoWord = Hexit(w(1)) & Hexit(w(0))
End Function

Private Function Hexit(x) As String
    Hexit = Hex(x)
    If Len(Hexit) < 2 Then Hexit = "0" & Hexit
End Function

Private Function HiWord(l As Long) As Integer
    HiWord = CInt("&h" & LongToWords(l).HiWord)
End Function

Private Function LoWord(l As Long) As Integer
    LoWord = CInt("&h" & LongToWords(l).LoWord)
End Function

Private Function MakeLong(ByVal HiWord As Integer, ByVal LoWord As Integer) As Long
      Call CopyMemory(MakeLong, LoWord, 2)
      Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function

Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

'Add 1 Command Button and 3 Text Boxes to your form.
'Set Text1 MultiLine Property to True.
'At Run-Time, insert the Text Box left margin to Text2, and the right margin to Text3,
'and press the button to apply the changes.
'Insert the following code to your form:

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal _
'lParam As Long) As Long
'Private Const EM_SETMARGINS = &HD3
'Private Const EC_LEFTMARGIN = &H1
'Private Const EC_RIGHTMARGIN = &H2
'
'Private Sub Command1_Click()
'SetMargins
'End Sub
'
'Private Sub Form_Load()
'Text2 = "30"
'Text3 = "30"
'SetMargins
'Text1.Text = "This program shows how to set the left and right margins _
'of a TextBox." & vbCrLf & " Its been taken from www.vb-town.com"
'End Sub
'
'Private Sub SetMargins()
'Dim left_margin As Integer
'Dim right_margin As Integer
'Dim long_value As Long
'left_margin = CInt(Text2.Text)
'right_margin = CInt(Text3.Text)
'long_value = left_margin * &H10000 + right_margin
'SendMessage Text1.hwnd, _
'EM_SETMARGINS, _
'EC_LEFTMARGIN Or EC_RIGHTMARGIN, _
'long_value
'End Sub
'
