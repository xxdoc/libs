Attribute VB_Name = "modSyntaxHighlighting"
Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

'THIS MODULE IS USED IN MULTIPLE LIBRARIES
'EDIT IT CAREFULLY,
'DEPENDANCIES - CURRENT DEVCONTROL WITH RTF EXTENDER
'
'Updated Oct 19 03 - removed global pDevControl made local to 2 fx
'change in frmIntellisense and devcontrol.ctl
'


Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private vbTokens() As String
Private jsTokens() As String
Private lvbTokens() As String
Private ljsTokens() As String
Private bvbtokens() As Byte
Private bjsTokens() As Byte

Private tokensInitalized As Boolean

Sub InitalizeTokens()
    
    tokensInitalized = True
    
    vbTokens() = Split("If,Then,Else,ElseIf,Case,Default,With,End," & _
                       "Select,cStr,Not,Exit,Function,Sub,And,Or,Xor," & _
                       "For,While,Next,Wend,Do,Loop,Until,Mid,InStr," & _
                       "Left,Trim,LTrim,Right,UBound,LBound,Len,Split," & _
                       "Call,True,False,Set,To,Each,In,Is,Nothing,Dim," & _
                       "Redim,Preserve,On,Error,Resume,True,False,IsArray," & _
                       "IsObject,IsNumeric,Const", ",")
                    
    jsTokens() = Split("if,else,switch,new,var,function,eval,break,exit," & _
                       "for,while,case,default,true,false,NaN", ",")
    
    ReDim lvbTokens(UBound(vbTokens))
    ReDim ljsTokens(UBound(jsTokens))
    ReDim bvbtokens(UBound(vbTokens))
    ReDim bjsTokens(UBound(jsTokens))

   Dim i As Integer

    For i = 0 To UBound(vbTokens)
        lvbTokens(i) = LCase(vbTokens(i))
        bvbtokens(i) = LCaseLeft1(vbTokens(i))
    Next

    For i = 0 To UBound(jsTokens)
        ljsTokens(i) = LCase(jsTokens(i))
        bjsTokens(i) = LCaseLeft1(jsTokens(i))
    Next
    
End Sub

Sub SyntaxHighlightLine(pDevControl As ctlDevControl, lineIndex As Long, isVbs As Boolean, Optional MakeSureClear As Boolean = True)
    Dim tmp As String, i As Integer, J As Integer
    Dim commentStart As Integer, commentlength As Integer
    Dim lineStart As Long
    Dim words() As String
    Dim tokens() As String
    Dim lTokens() As String
    Dim bTokens() As Byte
    
    If Not tokensInitalized Then InitalizeTokens
    
    If isVbs Then
        tokens() = vbTokens()
        lTokens() = lvbTokens()
        bTokens() = bvbtokens()
    Else
        tokens() = jsTokens()
        lTokens() = ljsTokens()
        bTokens() = bjsTokens()
    End If
    
    With pDevControl.clsRtf
        tmp = .GetLine(lineIndex)
        If Len(Trim(tmp)) = 2 Then Exit Sub 'account for the vbcrlf
        
        lineStart = .IndexOfFirstCharOnLine(lineIndex)
        
        If MakeSureClear Then  'clear previous formatting
            PreformHighlight pDevControl, lineStart, Len(tmp), vbBlack
        End If
        
        commentStart = CommentStartChar(tmp, isVbs)
        commentlength = Len(tmp) - commentStart
        
        If commentStart > 0 Then
            tmp = Mid(tmp, 1, commentStart)
            PreformHighlight pDevControl, (lineStart + commentStart), commentlength, RGB(0, &H88, 0)
        End If
        
        tmp = Replace(tmp, vbCrLf, Empty) 'cleanup
        tmp = Replace(tmp, vbTab, " ") 'cleanup
        tmp = Replace(tmp, "(", " ") 'valid word divider
        tmp = Replace(tmp, ",", " ") 'valid word divider
        tmp = Replace(tmp, "{", " ") 'valid word divider
        tmp = Replace(tmp, ";", " ") 'valid word divider
        tmp = Replace(tmp, ":", " ") 'valid word divider
        
        
        'tmp is not whole line no comments
        If Len(Trim(tmp)) = 0 Then Exit Sub
        
        'now block out all the quoted strings
        'word & character indexes remain unchanged
        RemoveQuotedStrings tmp
        
        words() = Split(tmp, " ")
        
        Dim wordStartIndex As Integer
        
        wordStartIndex = 0
        Dim lCasethisWord As String
        Dim thisWordFirstLetter As Byte
        For i = 0 To UBound(words)
            If Len(words(i)) > 0 Then
                lCasethisWord = LCase(words(i))
                thisWordFirstLetter = LCaseLeft1(words(i))
                'Debug.Print lCasethisWord & ":" & thisWordFirstLetter
                For J = 0 To UBound(tokens)
                    'Debug.Print LCase(tokens(j)) & " " & LCase(words(i))
                     If bTokens(J) = thisWordFirstLetter Then
                        If lTokens(J) = lCasethisWord Then
                            PreformHighlight pDevControl, lineStart + wordStartIndex, Len(words(i)), RGB(0, 0, &H88), tokens(J)
                        End If
                     End If
                Next
                wordStartIndex = wordStartIndex + Len(words(i))
            End If
            wordStartIndex = wordStartIndex + 1 'for the spaces
        Next
                   
    End With
End Sub

Function LCaseLeft1(s As String) As Byte
    If Len(s) = 0 Then Exit Function
    LCaseLeft1 = AscW(s)
    If LCaseLeft1 < 97 Then
        LCaseLeft1 = LCaseLeft1 + (97 - 65)
    End If
End Function

Sub PreformHighlight(pDevControl As ctlDevControl, selStart, selLength, color, Optional selText = "")
    
    Dim rtf As RichTextBox
    
    Set rtf = pDevControl.clsRtf.GetRtf
 
reSelect:
        rtf.selStart = IIf(selStart >= 0, selStart, 1)
        rtf.selLength = IIf(selLength > 0, selLength, 1)
        
        If Len(selText) > 0 Then
            rtf.selText = selText 'for proper case
            selText = ""
            GoTo reSelect
        End If
        
        pDevControl.clsRtf.HighLightSelection vbWhite, color
 
    
End Sub

Sub RemoveQuotedStrings(sIn As String)
    Dim dq As Integer
    Dim match As Integer, sLen As Integer
    Dim tmp As String
    
    dq = InStr(sIn, """")
    While dq > 0
        dq = InStr(sIn, """")
        If dq < 1 Then Exit Sub
        match = InStr(dq + 1, sIn, """") 'find its closing dq
        If match < 1 Then Exit Sub 'err.raise turn line red?
        sLen = match - dq + 1  'entire length of the quoted string
        tmp = Mid(sIn, 1, dq - 1) & String(sLen, "-") & Mid(sIn, match + 1, Len(sIn))
        sIn = tmp
    Wend
    
End Sub

Function CommentStartChar(ByVal sLine As String, Optional isVbs As Boolean) As Integer
    Dim commentChar As String
    Dim sq As Integer, dq As Integer, startAt As Integer
    
    commentChar = IIf(isVbs, "'", "//")
    
    If Not isVbs Then
        'so we dont have to deal with the possibility of single or double quoted strings
        sLine = Replace(sLine, "'", " ")
    End If
    
    startAt = 1
    
top:
    sq = InStr(startAt, sLine, commentChar)
    dq = InStr(startAt, sLine, """")
    
    If sq < 1 Then Exit Function
    
    If dq > 1 And dq < sq Then
        'we are in a quoted string, find the end quote and change startAt
        dq = InStr(dq + 1, sLine, """")
        If dq < 1 Then 'no close quote found? exit
            Exit Function
        Else
            startAt = dq + 1
            GoTo top
        End If
    End If
    
    'by the time we get here, sq should be at the first char of comment block
    CommentStartChar = sq - IIf(sq = 1, 0, 1)
    'think this leaves a bug but acceptable for now
End Function


Public Function LMouseDown() As Boolean
    GetAsyncKeyState vbKeyLButton
    LMouseDown = Not (GetAsyncKeyState(vbKeyLButton) And &HFFFF) = 0
End Function

Public Function RMouseDown() As Boolean
    GetAsyncKeyState vbKeyRButton
    RMouseDown = Not (GetAsyncKeyState(vbKeyRButton) And &HFFFF) = 0
End Function

