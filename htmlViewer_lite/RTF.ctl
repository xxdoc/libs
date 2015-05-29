VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl htmlControl 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   ScaleHeight     =   5025
   ScaleWidth      =   9105
   ToolboxBitmap   =   "RTF.ctx":0000
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4770
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   8414
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"RTF.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "htmlControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'copyright 2002 David Zimmer <dzzie@yahoo.com>  http://sandsprite.com
'all rights reserved

'this is a simple and fast html syntax highlight control requiring only a rich text box

'color find supports regex searchs
'Auto Color only highlights intresting html tags for full syntax highlighting, then use the general option for it.
'color find or auto color..all previous highlighting will be lost


Option Explicit

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

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)

Private Const MM_TWIPS = 6
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9

Enum editorChoice
    stdtext
    Highlight
End Enum

Private startAt As Long
Private findWhat As String
Private ww As Boolean
Private dirty As Boolean
Private surpressTools As Boolean
Private DisplayEditor As editorChoice
Private RTFHeader As String
Private myFont As String
Private myFontSize As Long

Private oRegExp As Object 'New RegExp
Private textColor, tagColor, propColor, propColorVal, commentColor

Private Const Footer = "}" '"\par \plain\f2\fs17\cf0" & vbCrLf & "\par }"

Public rtf As RichTextBox
Public Event RightClick()
Public Event KeyPress(KeyCode As Integer)


Public Sub HtmlCleanup()
    Me.ReplaceText ">", ">" & vbCrLf
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "editor", DisplayEditor
        .WriteProperty "ww", ww
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        EditorDisplayed = .ReadProperty("editor", stdtext)
        WordWrap = .ReadProperty("ww", True)
    End With
End Sub

Property Let EditorDisplayed(choice As editorChoice)
    DisplayEditor = choice
    If DisplayEditor = stdtext Then
        rtb.text = rtb.text
    Else
        HtmlHighLight
    End If
    UserControl.PropertyChanged "editor"
End Property

Property Get EditorDisplayed() As editorChoice
    EditorDisplayed = DisplayEditor
End Property

Public Property Let text(it)
    Dim tmp
    
    If Left(it, 6) = "font==" Then
        'hack to get around binary compat :(
        myFont = Replace(it, "font==", "")
        'rtb.Font.Name = myFont
        Exit Property
    End If
    
    If Left(it, 10) = "fontsize==" Then
        'hack to get around binary compat :(
        tmp = Replace(it, "fontsize==", "")
        If IsNumeric(tmp) Then
            myFontSize = CLng(tmp)
            'rtb.Font.Size = myFontSize
        Else
            myFontSize = 17
        End If
        Exit Property
    End If
    
    rtb.text = it
    dirty = False
    WordWrap = ww
    If DisplayEditor = Highlight Then HtmlHighLight
End Property

Public Property Get text()
        text = rtb.text
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = dirty
End Property
Public Property Let IsDirty(x As Boolean)
    dirty = x
End Property

Public Property Get FindString()
    FindString = CollapseConstants(findWhat)
End Property

Public Property Let FindString(it)
    findWhat = ExpandConstants(it)
End Property

Property Let Locked(setting As Boolean)
    rtb.Locked = setting
End Property
Property Get Locked() As Boolean
    Locked = rtb.Locked
End Property

Function ExpandConstants(ByVal strIn) As String
    strIn = Replace(strIn, "<TAB>", vbTab, , , vbTextCompare)
    strIn = Replace(strIn, "<CRLF>", vbCrLf, , , vbTextCompare)
    strIn = Replace(strIn, "<CR>", vbCr, , , vbTextCompare)
    ExpandConstants = CStr(Replace(strIn, "<LF>", vbLf, , , vbTextCompare))
End Function

Function CollapseConstants(ByVal strIn) As String
    strIn = Replace(strIn, vbTab, "<TAB>", , , vbTextCompare)
    strIn = Replace(strIn, vbCrLf, "<CRLF>", , , vbTextCompare)
    strIn = Replace(strIn, vbCr, "<CR>", , , vbTextCompare)
    CollapseConstants = CStr(Replace(strIn, vbLf, "<LF>", , , vbTextCompare))
End Function

Public Property Let WordWrap(on_ As Boolean)
    ww = on_
    If on_ Then rtb.rightMargin = 0 Else Call SetRightMargain
    UserControl.PropertyChanged "ww"
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = ww
End Property

Public Property Let SelText(txt)
   rtb.SelText = txt
End Property

Public Property Get SelText()
   SelText = rtb.SelText
End Property

Public Property Let SelStart(x)
   rtb.SelStart = x
End Property

Public Property Get SelStart()
    SelStart = rtb.SelStart
End Property

Public Property Let SelLen(x)
    rtb.SelLength = x
End Property

Public Property Get SelLen()
   SelLen = rtb.SelLength
End Property

Public Property Let Enabled(t As Boolean)
    rtb.Enabled = t
End Property

Public Property Get Enabled() As Boolean
    Enabled = rtb.Enabled
End Property

Public Property Get hWnd() As Long
    hWnd = rtb.hWnd
End Property

Public Sub SelSpan(start, length)
    On Error Resume Next
    rtb.SelStart = start
    rtb.SelLength = length
End Sub

Public Sub CopySelection()
    Dim sel As String
    
    If rtb.SelLength < 1 Then Exit Sub
    sel = rtb.SelText
    
    Clipboard.Clear
    Clipboard.SetText sel
End Sub

Public Sub Cut()
    CopySelection
    rtb.SelText = Empty
End Sub

Public Sub SelectAll()
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.text)
End Sub

 

Private Sub rtb_Change()
    dirty = True
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then RaiseEvent RightClick
End Sub

Public Sub LoadFile(path As String)
    Dim tmp As String
    tmp = ReadFile(path)
    If InStr(tmp, vbCrLf) < 1 Then 'assume unix style line endings..
        tmp = Replace(tmp, vbLf, vbCrLf)
    End If
    rtb.text = tmp
    WordWrap = ww
    If DisplayEditor = Highlight Then HtmlHighLight
End Sub

Private Sub UserControl_Initialize()
    rtb.Left = 0
    rtb.Top = 0
    startAt = 1
    
    Set rtf = rtb 'so they can hook events and directly modify it..
    textColor = vbBlack
    tagColor = "7B0000"
    propColor = "FF0000"
    propColorVal = "0000FF"
    commentColor = "008800"
    
    On Error GoTo hell
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.IgnoreCase = True
    oRegExp.Global = True
    
    Dim rtfColor(8) As String, colorTbl As String
    rtfColor(0) = GetRTFColor(textColor)
    rtfColor(1) = GetRTFColor(tagColor)
    rtfColor(2) = GetRTFColor(propColor)
    rtfColor(3) = GetRTFColor(propColorVal)
    rtfColor(4) = GetRTFColor(commentColor)
    rtfColor(5) = GetRTFColor("FF0000")
    rtfColor(6) = GetRTFColor("9D06C8")
    rtfColor(7) = GetRTFColor("67825F")
    rtfColor(8) = GetRTFColor("5219C5")
    
    'colorTbl = "{\colortbl" & Join(rtfColor, ";") & ";}" & vbCrLf
    
    colorTbl = "{\colortbl" & Join(rtfColor, "") & "}"
    
    myFont = "Courier"
    myFontSize = 21
    
    'RTFHeader = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{" & _
    '            "\f0\fswiss ______;}{\f1\froman\fcharset2 " & _
    '            "Symbol;}{\f2\fswiss ______;}}" & vbCrLf _
    '            & colorTbl & "\deflang1033\pard\plain\f0\fs17 "
    
    RTFHeader = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033" & _
                "{\fonttbl{\f0\fswiss ______;}}" & _
                colorTbl & _
                "\deflang1033\pard\f0\fs--"
    
Exit Sub
hell:
        MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub ScrollToTop()
    rtb.SelStart = 1
    rtb.SelLength = 0
    On Error Resume Next
    rtb.SetFocus
End Sub

Public Sub ReplaceText(find, changeTo)
   rtb.text = Replace(rtb.text, ExpandConstants(find), ExpandConstants(changeTo), , , vbTextCompare)
   If DisplayEditor = Highlight Then HtmlHighLight
End Sub

Private Function GetRTFHEader() As String
    If Len(Trim(myFont)) = 0 Then myFont = "Courier"
    If myFontSize < 5 Then myFontSize = 5
    GetRTFHEader = Replace(RTFHeader, "______", myFont)
    GetRTFHEader = Replace(RTFHeader, "--", myFontSize)
End Function

Private Sub UserControl_Resize()
    On Error Resume Next
    rtb.Height = UserControl.Height
    rtb.Width = UserControl.Width
End Sub

Public Sub MatchSize(it As Object, Optional border = 0)
    UserControl.Size it.Width - border, it.Height - border
End Sub

Public Sub find()
    If findWhat = Empty Then Exit Sub
    
    Dim x
    
    Me.ScrollToTop
    startAt = 1
    x = rtb.find(findWhat)
    If x >= 0 Then
        rtb.SelStart = x
        rtb.SelLength = Len(findWhat)
        startAt = x + 1
    Else
        MsgBox "String Not Found", vbInformation
    End If
   
End Sub

Public Sub findNext()
    Dim x As Long
    
    If findWhat = Empty Then Exit Sub
    
    x = rtb.find(findWhat, startAt)
    If x > 0 And startAt < (Len(rtb.text) - 1) Then
        rtb.SelStart = x
        rtb.SelLength = Len(findWhat)
        startAt = x + 1
    Else
        startAt = 1
        MsgBox "Search Complete", vbInformation
    End If
    
End Sub


Private Sub rtb_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 6: find
        Case 4: findNext
        Case Else: RaiseEvent KeyPress(KeyAscii)
    End Select
    With rtb
        If Chr(KeyAscii) = "<" Then .SelColor = CLng("&H" & tagColor)
        If InComment = True Then Exit Sub
        If InTag = True Then
            If Chr(KeyAscii) = "-" Then
                ' check if we are in a comment
                .SelStart = .SelStart - 3
                .SelLength = 3
                If .SelText = "<!-" Then .SelColor = CLng("&H" & commentColor)
                .SelStart = .SelStart + 4
            End If
            
            If Chr(KeyAscii) = " " Then
                If InPropval Then
                    .SelColor = CLng("&H" & propColorVal)
                Else
                    .SelColor = CLng("&H" & propColor)
                End If
            ElseIf Chr(KeyAscii) = "=" Then
                    .SelText = "="
                    .SelColor = CLng("&H" & propColorVal)
                    KeyAscii = 0
            ElseIf Chr(KeyAscii) = ">" Then
                    .SelColor = CLng("&H" & tagColor)
                    .SelText = ">"
                    KeyAscii = 0
                    .SelColor = CLng("&H" & textColor)
            End If
        End If
    End With
End Sub

Public Sub AppendIt(it)
   rtb.text = rtb.text & it
End Sub

Public Sub PrePendIt(it)
    rtb.text = it & rtb.text
End Sub

Public Sub SetRightMargain()
    Dim tm As TEXTMETRIC
    Dim longestLine As Long, hdc As Long, i As Long
    Dim lineCount As Long, lineLength As Long, lineIndex As Long
    Dim PrevMapMode
    
    lineCount = SendMessageLong(rtb.hWnd, EM_GETLINECOUNT, 0&, 0&)
    
    For i = 0 To lineCount - 1
        lineIndex = SendMessageLong(rtb.hWnd, EM_LINEINDEX, i, 0&)
        lineLength = SendMessageLong(rtb.hWnd, EM_LINELENGTH, lineIndex, 0&)
        If lineLength > longestLine Then longestLine = lineLength
    Next
        
    hdc = GetWindowDC(rtb.hWnd)
    
    If hdc Then
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        GetTextMetrics hdc, tm
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ReleaseDC rtb.hWnd, hdc
    End If
    
    rtb.rightMargin = longestLine * tm.tmMaxCharWidth
    
End Sub

Public Sub AutoColor()
    Dim tmpstr As String
    
    Screen.MousePointer = vbHourglass
    LockWindowUpdate rtb.hWnd
    
    tmpstr = rtb.text
    
    'regEx to escape RTF
    tmpstr = RegExReplace("([{}\\])", "\$1", tmpstr)
    tmpstr = RegExReplace("(\r)", "\par \r", tmpstr)
    
    HighlightTagAttributes tmpstr, "<form"
    HighlightTagAttributes tmpstr, "<input"
    HighlightTagAttributes tmpstr, "<option"
    HighlightTagAttributes tmpstr, "<select"
        
    tmpstr = WrappedReplace("</option>", 1, tmpstr)
    tmpstr = WrappedReplace("</select>", 1, tmpstr)
    tmpstr = WrappedReplace("</form>", 1, tmpstr)
    tmpstr = WrappedReplace("<script[\w\W]+?</script>", 7, tmpstr) '&HC000C0
    tmpstr = WrappedReplace("<!--[\w\W]+?-->", 6, tmpstr) ' &H808000
     
    rtb.TextRTF = GetRTFHEader() & tmpstr & "\plain\f2\fs17\cf0 " & Footer
    
    Screen.MousePointer = vbDefault
    LockWindowUpdate 0

End Sub

Public Sub HighlightTagAttributes(ByRef strIn, strTag As String)
    'Dim oMatch As Match, oMatches As MatchCollection
    Dim oMatch As Object, oMatches As Object
    Dim tmp, tmpstr 'modifies parent directly
    
    strIn = WrappedReplace(strTag, 1, strIn)
    oRegExp.Pattern = strTag & "[\w\W]+?>"
    
    Set oMatches = oRegExp.Execute(strIn)
   
    For Each oMatch In oMatches
        'tmp = RegExReplace("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+|[\w\W]+?)", "\plain\f2\fs17\cf2 $1\plain\f2\fs17\cf3 $2\plain\f2\fs17\cf3 ", oMatch.value)
        tmp = RegExReplace("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+|[\w\W]+?)", "\cf2 $1\cf3 $2\cf3 ", oMatch.value)
        strIn = Replace(strIn, oMatch.value, tmp)
    Next
    
End Sub

Function WrappedReplace(sFind, intColorIndex, strIn)
    'WrappedReplace = RegExReplace("(" & sFind & ")", "\plain\f2\fs17\cf" & intColorIndex & " $1\plain\f2\fs17\cf0 ", strIn)
    WrappedReplace = RegExReplace("(" & sFind & ")", "\cf" & intColorIndex & " $1\cf0 ", strIn)
End Function


'---------------------------------------------------------------
'----------------- stuff for the text pad bar -------------------
'---------------------------------------------------------------
 

'Private Sub cmdFind_Click()
'    If txtFind <> FindString Then
'        FindString = txtFind
'        find
'    Else
'        findNext
'    End If
'End Sub
'
'Private Sub cmdReplace_Click()
'    ReplaceText txtFind, txtReplace
'End Sub

'Sub SetColor(str, Optional Color As ColorConstants = vbBlack, Optional fSize = 10, Optional bold As Boolean = False, Optional italic As Boolean = False)
Public Sub ColorFind(sFind, intColorIndex As Integer)
    Dim tmpstr As String
    
    If intColorIndex < 0 Or intColorIndex > 8 Then
        MsgBox "invalid color index only 0-8"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    LockWindowUpdate rtb.hWnd
    
    tmpstr = rtb.text
    
    'regEx to escape RTF
    tmpstr = RegExReplace("([{}\\])", "\$1", tmpstr)
    tmpstr = RegExReplace("(\r)", "\par \r", tmpstr)

    'tmpstr = RegExReplace("(" & sFind & ")", "\plain\f2\fs17\cf" & intColorIndex & " $1\plain\f2\fs17\cf0 ", tmpstr)
    tmpstr = RegExReplace("(" & sFind & ")", "\cf" & intColorIndex & " $1\cf0 ", tmpstr)
    
    rtb.TextRTF = GetRTFHEader() & tmpstr & "\plain\f2\fs17\cf0 " & Footer
    
    Screen.MousePointer = vbDefault
    LockWindowUpdate 0
End Sub




'--------------------------------------------------------------------
'highlighting stuff below
'--------------------------------------------------------------------

Private Function GetRTFColor(ByVal Color As Variant) As String
  'this function accepts a VB color (long) or a HTML color (string) and
  'returns a RTF color table def.
  
  Const sHEX = "0123456789ABCDEF"
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  If VarType(Color) = vbLong Then
    lngRed = Color Mod 256&
    lngGreen = (Color Mod 65536) \ 256&
    lngBlue = Color \ 65536
  ElseIf VarType(Color) = vbString Then
    ' the string should be something like this: #D0D5DF
    Color = Right$(Color, 6)
    
    'find the position for each char in sHEX. Position is the value
    lngRed = 16& * (InStr(1, sHEX, Mid$(Color, 1, 1), vbTextCompare) - 1) + _
              1& * (InStr(1, sHEX, Mid$(Color, 2, 1), vbTextCompare) - 1)
    lngGreen = 16& * (InStr(1, sHEX, Mid$(Color, 3, 1), vbTextCompare) - 1) + _
                1& * (InStr(1, sHEX, Mid$(Color, 4, 1), vbTextCompare) - 1)
    lngBlue = 16& * (InStr(1, sHEX, Mid$(Color, 5, 1), vbTextCompare) - 1) + _
               1& * (InStr(1, sHEX, Mid$(Color, 6, 1), vbTextCompare) - 1)
  End If
  
  GetRTFColor = "\red" & CStr(lngRed) & "\green" & CStr(lngGreen) & "\blue" & CStr(lngBlue) & ";"
End Function

Private Sub HtmlHighLight()
    Dim tmpstr As String

    Screen.MousePointer = vbHourglass
    LockWindowUpdate rtb.hWnd
    
    tmpstr = rtb.text
    
    'regEx to escape RTF
    tmpstr = RegExReplace("([{}\\])", "\$1", tmpstr)
    tmpstr = RegExReplace("(\r)", "\par \r", tmpstr)
    
    'tags and prop/value pairs
    'tmpstr = RegExReplace("(<[^>]+>)", "\plain\f2\fs17\cf1 $1\plain\f2\fs17\cf0 ", tmpstr)
    'tmpstr = RegExReplace("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+)", "\plain\f2\fs17\cf2 $1\plain\f2\fs17\cf3 $2\plain\f2\fs17\cf1 ", tmpstr)
    tmpstr = RegExReplace("(<[^>]+>)", "\cf1 $1\cf0 ", tmpstr)
    tmpstr = RegExReplace("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+)", "\cf2 $1\cf3 $2\cf1 ", tmpstr)


    'comments
    'tmpstr = RegExReplace("(<!--[\w\W]+?-->)", "\plain\f2\fs17\cf4 $1\plain\f2\fs17\cf0 ", tmpstr)
    tmpstr = RegExReplace("(<!--[\w\W]+?-->)", "\cf4 $1\cf0 ", tmpstr)
    
    rtb.TextRTF = GetRTFHEader() & tmpstr & "\plain\f2\fs17\cf0 " & Footer
    
    Screen.MousePointer = vbDefault
    LockWindowUpdate 0
End Sub

Private Function RegExReplace(patrn, replStr, textStr)
  oRegExp.Pattern = patrn
  RegExReplace = oRegExp.Replace(textStr, replStr)
End Function

Private Function InTag() As Boolean
    If rtb.SelStart > 0 Then
        If InStrRev(rtb.text, "<", rtb.SelStart, vbTextCompare) > InStrRev(rtb.text, ">", rtb.SelStart, vbTextCompare) Then InTag = True
    End If
End Function

Private Function InComment() As Boolean
    If rtb.SelStart > 0 Then
        If InStrRev(rtb.text, "<!--", rtb.SelStart, vbTextCompare) > InStrRev(rtb.text, "-->", rtb.SelStart, vbTextCompare) Then InComment = True
    End If
End Function

Private Function InPropval() As Boolean
    Dim x, Y As Long
    x = InStrRev(rtb.text, """", rtb.SelStart, vbTextCompare)
    Y = InStrRev(rtb.text, "=", rtb.SelStart, vbTextCompare)
    If x > Y Then
        If InStrRev(rtb.text, """", x - 1, vbTextCompare) < InStrRev(rtb.text, "=", x - 1, vbTextCompare) Then InPropval = True
    End If
End Function

Function ReadFile(filename)
    On Error GoTo hell
    If Len(filename) = 0 Then Exit Function
    Dim f As Long
    Dim temp
      f = FreeFile
      temp = ""
       Open filename For Binary As #f        ' Open file.(can be text or image)
         temp = Input(FileLen(filename), #f) ' Get entire Files data
       Close #f
       ReadFile = temp
hell:
End Function

