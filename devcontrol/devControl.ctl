VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl ctlDevControl 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   LockControls    =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   7875
   Begin MSScriptControlCtl.ScriptControl ScriptControl 
      Left            =   3660
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin RichTextLib.RichTextBox rtfToolTip 
      Height          =   280
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   503
      _Version        =   393217
      BackColor       =   11206655
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"devControl.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   3315
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5847
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1e7
      TextRTF         =   $"devControl.ctx":0085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pNums 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3255
      ScaleWidth      =   435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   435
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   80
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "devControl.ctx":0105
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "devControl.ctx":050F
            Key             =   "func"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "devControl.ctx":08A9
            Key             =   "class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "devControl.ctx":0CF7
            Key             =   "closedFolder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "devControl.ctx":1011
            Key             =   "openFolder"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlDevControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

'THIS FILE IS USED IN MULTIPLE PROJECTS EDIT CAREFULLY
'
'DEPENDS ON MODSYNTAXHIGHLIGHT, FRMiNTELLISENSE, CLSRTFEXTENDER
'NEEDS PROJECT REFERENCE TO RTFCONTROL, CLSTOOLTIPMANAGER

Private Type obj
    id As String
    member() As String
    proto() As String
End Type

Private Objects() As obj
Private LineNumbersOn As Boolean
Private IntellisenseOn As Boolean

Public MDI_OffsetTop As Long
Public MDI_OffsetLeft As Long

Public isVbs As Boolean
Public WithEvents clsRtf As clsRtfExtender
Attribute clsRtf.VB_VarHelpID = -1
Public ToolTip As New clsToolTipManager
Public sc As Object

Event RightClicked()
Event ScriptError(description, source, line)

 

Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = rtf.Text
End Property
Property Let Text(t As String)
    rtf.Text = t
    SyntaxHighlight
End Property

Property Let useLineNumbers(b As Boolean)
    LineNumbersOn = b
    UserControl_Resize
    AddLineNums
End Property
Property Get useLineNumbers() As Boolean
    useLineNumbers = LineNumbersOn
End Property

Property Let useIntellisense(b As Boolean)
    IntellisenseOn = b
End Property
Property Get useIntellisense() As Boolean
    useIntellisense = IntellisenseOn
End Property

Sub Intellisense_SelectionMade(Text As String, prototype As String)
    rtf.selText = Text
    UserControl.Parent.Caption = prototype
    ToolTip.ShowToolTip prototype
    rtf.SetFocus
End Sub

Private Sub clsRtf_ArrowDownLine(prevlineIndex As Long)
   Dim curColumn As Integer
   curColumn = clsRtf.CurrentColumn
   SyntaxHighlight prevlineIndex - 1
   rtf.selStart = clsRtf.IndexOfFirstCharOnLine(prevlineIndex) + (curColumn - 1)
   rtf.SelColor = vbBlack
   hide rtfToolTip
End Sub

Private Sub clsRtf_ArrowUpLine(prevlineIndex As Long)
   Dim curColumn As Integer
   curColumn = clsRtf.CurrentColumn
   SyntaxHighlight prevlineIndex - 1
   rtf.selStart = clsRtf.IndexOfFirstCharOnLine(prevlineIndex - 2) + (curColumn - 1)
   rtf.SelColor = vbBlack
   hide rtfToolTip
End Sub

Private Sub clsRtf_ClickedToNewLine(PrevLine As Long, curLine As Long)
    SyntaxHighlight PrevLine
    ToolTip.HideToolTip
End Sub

Private Sub clsRtf_newLine(lineIndex As Long)
    SyntaxHighlight lineIndex - 1
    rtf.selStart = clsRtf.IndexOfFirstCharOnLine(lineIndex)
    rtf.SelColor = vbBlack
    hide rtfToolTip
End Sub

Private Sub rtf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        'Debug.Print "bastard double paste bug! IMA cheatoorrrr!"
        rtf.selText = Clipboard.GetText
        KeyCode = 0
    End If
    
    'If KeyCode = 40 Then SyntaxHighlight clsRtf.CurrrentLineIndex - 1
    'Debug.Print "devcontrol.rtf_KeyDown:" & KeyCode & " " & Shift
    
End Sub

Private Sub rtf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then RaiseEvent RightClicked
End Sub

'Private Sub rtf_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim fName As String
'    fName = Data.Files(1)
'    If clsFso.FileExists(fName) Then
'        rtf.Text = clsFso.ReadFile(fName)
'    End If
'End Sub

Sub PrepareForForcedTearDown()
    'this will run the code in class terminate removing subclasses
    Set clsRtf = Nothing
End Sub

Sub ExecuteCurrentScript()
    ScriptControl.ExecuteStatement rtf.Text
End Sub

Private Sub ScriptControl_Error()

        With ScriptControl.Error
            Dim c As String
                    
            c = clsRtf.GetLine(.line)
            
            clsRtf.LockUpdate
            clsRtf.HighLightLine .line
            If lineNumber > 6 Then
                clsRtf.ScrollToLine .line - 6
            End If
            clsRtf.UnlockUpdate
            
            RaiseEvent ScriptError(.description, .source, .line)
        
        End With
        
End Sub


Private Sub UserControl_Initialize()
        
    Set sc = ScriptControl
    Set clsRtf = New clsRtfExtender
    Set ToolTip.pDevControl = Me
    
    clsRtf.SetRtf rtf
    ToolTip.SetRTFObject rtfToolTip, rtf
    
    isVbs = True 'just force it, spScript does not support JS or others
    
    AddLineNums
    
End Sub

Property Get RTFTEXT() As String
    RTFTEXT = rtf.TextRTF
End Property

Public Sub GetCaretPos(x, y)
  'fills in these variables
  Dim p As POINTAPI
  p = clsRtf.CaretPos
  x = p.x
  y = p.y
  
  If MDI_OffsetLeft > 0 Or MDI_OffsetTop > 0 Then
    x = x + UserControl.Parent.devControl.left
    y = y + UserControl.Parent.devControl.top + 500
  Else
    x = x + UserControl.Parent.left + UserControl.Parent.devControl.left
    y = y + UserControl.Parent.top + UserControl.Parent.devControl.top + 500
  End If
  
End Sub
Private Sub clsRtf_AutoComplete()
    
    If Not IntellisenseOn Then Exit Sub
    
    'if preceeding word = " " then
    DisplayTopLevelObjects
    'else locate preceeding class, if recgonized then
    'if only one match then autocomplete
    'else show listing for this window with first letter entries highlighted
    
End Sub

Private Sub clsRtf_Scrolled()
    AddLineNums
End Sub

Private Sub rtf_Click()
    On Error Resume Next
    Unload frmIntellisense
    'If rtfToolTip.Visible Then rtfToolTip.Visible = False
End Sub

Private Sub rtf_KeyUp(KeyCode As Integer, Shift As Integer)
            
    If Not IntellisenseOn Then Exit Sub
    
    If KeyCode = 190 Then '"."
    
        Start = rtf.selStart
        spac = InStrRev(rtf.Text, " ", Start) + 1
        crlf = InStrRev(rtf.Text, vbCrLf, Start) + 1
        
        If crlf > spac And spac > 0 Then spac = crlf + 1
        
        'if we cant find a whole word, then get to beginning of line
        'or in teh caseof the first line..to the beginning of file
        If spac < 1 Then spac = InStrRev(rtf.Text, vbLf, Start) + 1
        If spac < 1 Then spac = 1
               
        If spac > 0 Then
            'get the word to match..but did we take to much?
            s = LCase(Mid(rtf.Text, spac, Start - spac + 1))
            
            'are we in a function arg list?
            spac = InStrRev(s, ",")
            If spac > 0 Then s = Mid(s, spac + 1, Len(s))
            
            'are we in a function call?
            spac = InStrRev(s, "(")
            If spac > 0 Then s = Mid(s, spac + 1, Len(s))
            
            'UserControl.Parent.Caption = s
            For i = 1 To UBound(Objects)
                If s Like Objects(i).id & "." Then
                    'UserControl.Parent.Caption = Objects(i).id & " " & UBound(Objects(i).member)
                    FillandPositionListView frmIntellisense.lv, Objects(i)
                    Exit For
                End If
            Next
        End If
        
    End If
    
    If KeyCode = 71 And Shift = 2 Then GotoLine 'ctrl-G
       
    
    
End Sub

Sub GotoLine(Optional x As Integer = -1)
     On Error Resume Next
     If x < 0 Then
        x = InputBox("Goto Line Number")
     End If
     If IsNumeric(x) Then clsRtf.ScrollToLine CLng(x)
End Sub

Private Sub AddLineNums()
        
    If Not LineNumbersOn Then Exit Sub
    
    With pNums
        .Cls
        .CurrentX = 0
        .CurrentY = 0
        
        Start = clsRtf.TopLineIndex
        Max = clsRtf.VisibleLines + Start
        For i = Start + 1 To Max + 3
            pNums.Print IIf(i < 100, " ", "") & IIf(Len(i) < 2, "0", "") & i
        Next
        
        .Refresh
    End With
    
End Sub


'-----------------------------------------------------------------
'|    General Subs Here
'-----------------------------------------------------------------

Sub LoadProtoTypes(fPath As String)
    
    Dim f As Long
    Dim tmp
    On Error GoTo hell
    
    fPath = Replace(fPath, "%AP%", App.path)
    
    ReDim Objects(0)
    
    If Not FileExists(fPath) Then
        MsgBox "Could not Load prototype file: " & vbCrLf & vbCrLf & fPath, vbInformation
        Exit Sub
    End If
    
    f = FreeFile
    Open fPath For Input As f
    
    While Not EOF(f)
        Line Input #f, tmp
        
        tmp = Replace(Trim(tmp), vbCrLf, "")
        
        If tmp = "" Then GoTo nextOne
        If left(tmp, 1) = "[" Then
            ReDim Preserve Objects(UBound(Objects) + 1)
            Objects(UBound(Objects)).id = Mid(tmp, 2, (InStr(2, tmp, "]") - 2))
            'Debug.Print "Objects(" & UBound(Objects) & " ).id" & " : " & Objects(UBound(Objects)).id
        ElseIf left(tmp, 1) = "#" Then
            GoTo nextOne
        Else
            If InStr(tmp, ":") < 1 Then
                MsgBox "Improper member format: Section: " & Objects(UBound(Objects)).id & vbCrLf & vbCrLf & tmp, vbInformation
                GoTo nextOne
            End If
            
            tmp = Split(tmp, ":")
            push Objects(UBound(Objects)).member(), Trim(tmp(0))
            push Objects(UBound(Objects)).proto(), Trim(tmp(1))
        End If
    
nextOne:
    Wend
    
Exit Sub
hell: MsgBox Err.description
End Sub

Private Sub FillandPositionListView(l As ListView, o As obj)
    Dim J As Integer
    Dim li As ListItem
    
    l.ListItems.Clear
    
    For J = 0 To UBound(o.member)
        Set li = l.ListItems.Add(, , o.member(J))
        li.Tag = Replace(Trim(o.proto(J)), vbTab, "")
        If InStr(1, li.Tag, "Sub", vbTextCompare) > 0 Or _
           InStr(1, li.Tag, "Function", vbTextCompare) > 0 Then
                li.SmallIcon = "func"
        Else
                li.SmallIcon = "prop"
        End If
    Next
    
    'Debug.Print l.ListItems.Count & " ListItems"
    frmIntellisense.ResizeAndActivate J, Me
    
End Sub

Sub DisplayTopLevelObjects()
    Dim J As Integer
    Dim li As ListItem
    
    With frmIntellisense
    
        .lv.ListItems.Clear
        
        For J = 1 To UBound(Objects)
            Set li = .lv.ListItems.Add()
            li.Text = Replace(Objects(J).id, "*", "")
            li.Tag = Trim(Objects(J).id)
            li.SmallIcon = "class"
        Next
        
    End With
    
    frmIntellisense.ResizeAndActivate J, Me
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        useIntellisense = .ReadProperty("intellisense", True)
        useLineNumbers = .ReadProperty("linenos", True)
    End With
End Sub

Private Sub UserControl_Terminate()
    Set clsRtf = Nothing
    Set ToolTip = Nothing
    Set sc = Nothing
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "intellisense", IntellisenseOn
        .WriteProperty "linenos", LineNumbersOn
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    pNums.left = IIf(LineNumbersOn, 0, 0 - pNums.Width - 15)
    rtf.left = pNums.left + pNums.Width + 15
    rtf.Height = UserControl.Height - 50
    rtf.Width = UserControl.Width - rtf.left - 50
    pNums.Height = rtf.Height
    
    AddLineNums
    
End Sub

Sub SyntaxHighlight(Optional lineIndex As Long = -1)
    Dim i As Long
    On Error GoTo hell
        
    With clsRtf
            
        .LockUpdate
            
        If lineIndex < 0 Then
            'clear all comment formatting in one bulk operation
            rtf.selStart = 1
            rtf.selLength = Len(rtf.Text)
            rtf.SelColor = vbBlack
            rtf.selLength = 0
    
            For i = 0 To .lineCount
                SyntaxHighlightLine Me, i, isVbs, False
            Next
            rtf.selStart = 1
            rtf.SelColor = vbBlack
        Else
            SyntaxHighlightLine Me, lineIndex, isVbs
            'rtf.selStart = .IndexOfFirstCharOnLine(lineIndex + 1)
            'rtf.SelColor = vbBlack
        End If
        
    End With
    
hell:
    rtf.selLength = 0
    clsRtf.UnlockUpdate
    
End Sub

Private Sub hide(o As Object)
    On Error Resume Next
    o.Visible = False
End Sub

Friend Function CaretPos() As POINTAPI
    CaretPos = clsRtf.CaretPos
End Function
