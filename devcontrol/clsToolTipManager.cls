VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTipManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

Private Type typeArgInfo
    startAt As Integer
    argLen As Integer
End Type

Private argInfo() As typeArgInfo
Public ArgCount As Integer

Private WithEvents rtf As RichTextBox
Attribute rtf.VB_VarHelpID = -1
Private mExt As New clsRtfExtender
Private WithEvents parentRTF As RichTextBox
Attribute parentRTF.VB_VarHelpID = -1

Public ToolTipText As String
Public pDevControl As ctlDevControl

Private argSections() As String
Private prePend As String
Private apend As String

Private keyBuffer() As Byte
Private MaxCharWidth As Long
Private charHeight As Long
Private BoldedSection As Integer
Private inQuotedString As Boolean
Private CurrentCharPos As Long
 
Public isVisible As Boolean

Private Sub ShowIt()
    isVisible = True
    rtf.Visible = True
End Sub

Private Sub HideIt()
    isVisible = False
    rtf.Visible = False
End Sub

Property Get HighlightedArg() As Integer
    HighlightedArg = BoldedSection
End Property

Friend Sub SetRTFObject(oRtf As Object, pRtf As Object)
    Set parentRTF = pRtf
    Set rtf = oRtf
    mExt.SetRtf rtf
    MaxCharWidth = mExt.MaxCharWidth
    charHeight = mExt.charHeight
End Sub

Sub HideToolTip()
    HideIt
End Sub

Sub BoldPreviousSection()
    On Error Resume Next
    If UBound(argInfo) = -1 Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    
    If BoldedSection <= UBound(argInfo) Then
        mExt.LockUpdate
        BoldedSection = BoldedSection - 1
        rtf.Text = rtf.Text 'remove all formatting
        rtf.selStart = argInfo(BoldedSection).startAt
        rtf.selLength = argInfo(BoldedSection).argLen
        rtf.SelBold = True
        rtf.selStart = 1
        mExt.UnlockUpdate
     End If
End Sub

Sub BoldNextSection()
    
    On Error Resume Next
    If UBound(argInfo) = -1 Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    
    If BoldedSection <= UBound(argInfo) Then
        mExt.LockUpdate
        BoldedSection = BoldedSection + 1
        rtf.Text = rtf.Text 'remove all formatting
        rtf.selStart = argInfo(BoldedSection).startAt
        rtf.selLength = argInfo(BoldedSection).argLen
        rtf.SelBold = True
        rtf.selStart = 1
        mExt.UnlockUpdate
     End If
                     
End Sub

Sub ShowToolTip(sToolTipText As String)
    Dim i As Integer
    Dim p As Integer
    
    sToolTipText = Trim(sToolTipText)
    ToolTipText = sToolTipText
    
    prePend = Empty
    apend = Empty
    BoldedSection = 0
    ReDim argInfo(0)
    ReDim keyBuffer(0)
    
    p = InStr(sToolTipText, "(")
    If p > 0 Then
        prePend = Mid(sToolTipText, 1, p)
        sToolTipText = Mid(sToolTipText, p + 1)
        p = InStrRev(sToolTipText, ")")
        If p > 0 Then
            apend = Mid(sToolTipText, p)
            sToolTipText = Mid(sToolTipText, 1, p - 1)
        End If
    End If
    
    'sToolTipText is now just the function arguments
    
   
    
    If Len(sToolTipText) = 0 Then sToolTipText = " "
        
    argSections() = Split(sToolTipText, ",")
    
    ReDim argInfo(UBound(argSections))
        
    Dim pt As POINTAPI

    pt = pDevControl.clsRtf.CaretPos()
    
    With rtf
        .Move pt.x - 100, pt.y + 370

        .Text = prePend
        For i = 0 To UBound(argSections)
            p = Len(.Text)
            .Text = .Text & argSections(i) & IIf(i <> UBound(argSections), ",", "")
            argInfo(i).startAt = p
            argInfo(i).argLen = Len(argSections(i))
        Next
         
        .Text = .Text & apend
        .selStart = argInfo(BoldedSection).startAt
        .selLength = argInfo(BoldedSection).argLen
        .SelBold = True
        
        ResizeRTFFor Len(ToolTipText)
        
        .selStart = 1
        
        ShowIt
         
        
    End With
    
    
    
    
    'will the tooltip be to far offscreen? reposition left and up as needed
    'is the text too long for it? resize width and height of rtftt
    

End Sub

 

Private Function WidthAvailable() As Long
    'max width available from where tooltip .left is on parent UC to
    'UC.width,
    
End Function

Private Sub ResizeRTFFor(cLen As Long)
    Dim oneLineWidth As Long
    Dim twoLineWidthNeeded As Long
    Dim availWidth As Long
    Dim pRtfWidth As Long
    Dim lineCount As Long
    
    
    oneLineWidth = (cLen * MaxCharWidth) + 200
    pRtfWidth = pDevControl.clsRtf.GetRtf.Width
    
    If pRtfWidth < 500 Then Exit Sub 'no tooltips to small
    
    'If rtf.left + 1200 > pRtfWidth Then 'to close to right border move left
    
    If rtf.left + oneLineWidth < pRtfWidth Then
        Debug.Print "one line no reposition"
        rtf.Width = oneLineWidth + 350
        rtf.Height = charHeight + 70
    ElseIf (rtf.left + oneLineWidth - 800) < pRtfWidth Then
        Debug.Print "one line -800 reposition"
        rtf.Width = oneLineWidth
        rtf.Height = charHeight + 70
        rtf.left = rtf.left - 800
    Else
        availWidth = pRtfWidth - rtf.left
        'never let it be on more than 2 lines
        twoLineWidthNeeded = oneLineWidth / 2 + 200
        
        If twoLineWidthNeeded > availWidth Then
            rtf.left = pRtfWidth - 800 - twoLineWidthNeeded
            Debug.Print "two lines w/reposition"
        Else
            Debug.Print "two lines no reposition"
        End If
        
         
        rtf.Height = (charHeight * 2) + 40
        rtf.Width = twoLineWidthNeeded + 200
        
    End If
    
    If rtf.left < 50 Then
        rtf.left = 50
        Debug.Print "had to reposition tooltip right..it was offscreen"
    End If
        
    
    
End Sub

 

Private Sub parentRTF_Click()
    If isVisible Then HideIt
End Sub




 
Private Sub parentRTF_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    'these routines can still bug out...especially when quoted arg contains
    'a comma. Does not recgonize if user clicked to different argument position
    
    Dim prevWord As String, indexFound As Integer
     
    If Not isVisible Then Exit Sub
    
    prevWord = pDevControl.clsRtf.WordBeforeCursor(indexFound, " ", "(")
   
    If parentRTF.selStart = CurrentCharPos Then
        'they pressed a key but nothing happened, like right arrow when already at end of line
        Exit Sub
    Else
        CurrentCharPos = parentRTF.selStart
    End If
    
    If KeyCode = 13 Then
        HideIt 'return key
        Exit Sub
    End If
    
    'if keycode is a letter or number then then
    'not arrow keys or delete anyway
    If KeyCode > 60 Then
        push keyBuffer, KeyCode
        If KeyCode = 222 And Shift = 1 Then  'doubel quote
            inQuotedString = Not inQuotedString
        End If
    End If
    
    If KeyCode = 39 Then 'right arrow, manually append buffer
        Dim lastCharByPos As Byte
        lastCharByPos = pDevControl.clsRtf.CharBeforeCursor
        If lastCharByPos = 44 Then 'ascii ,'
            lastCharByPos = 188 'keycode for , pretty gay huh
        End If
        push keyBuffer, lastCharByPos
        If lastCharByPos = 222 Then inQuotedString = Not inQuotedString
    End If
    
    If KeyCode = 188 Then
        BoldNextSection
    End If
    
    If KeyCode = 37 Or KeyCode = 8 Then  'leftarrow or delete, decrement buffer
        If keyBuffer(UBound(keyBuffer)) = 188 Then BoldPreviousSection
        If keyBuffer(UBound(keyBuffer)) = 222 Then inQuotedString = Not inQuotedString
        If UBound(keyBuffer) = 0 Then Exit Sub
        ReDim Preserve keyBuffer(UBound(keyBuffer) - 1)
    End If
    
    If KeyCode = 39 Then 'right arrow
        If keyBuffer(UBound(keyBuffer)) = 188 Then BoldNextSection
        'add stepped by char to keybuffer
    End If

 
End Sub

Private Sub rtf_GotFocus()
    On Error Resume Next
    parentRTF.SetFocus
End Sub
