VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl CodeEdit 
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ScaleHeight     =   1950
   ScaleWidth      =   3150
   Begin VB.PictureBox picLineNumbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   1695
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   1e7
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"CodeEdit.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CodeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : CodeEdit Active-X Control
' Date      : Dec 2002
' Author    : Karl Durrance (KDurrance@Hotmail.com)
' Purpose   : CodeEdit Control
'             Provides Efficient syntax colouring and basic editor functions
'---------------------------------------------------------------------------------------
' This Control is FreeWare, and may be freely used and distributed in your own projects

' You may not publish any part of this source code
' without giving credit to me for the parts that I have created..

' The control is 'As-is', use it at your own risk

' Using this control means that you understand that I will not be responsible
' for any damage this may cause to any file you open with this control.
'---------------------------------------------------------------------------------------

Option Explicit

#Const USE_SUBCLASS = 0
#If USE_SUBCLASS Then
   Private sc As CSubclass
   'Implements ISubclass
#End If

#Const ECOLOR_DEFINED = 1
#If ECOLOR_DEFINED = 0 Then
Public Enum eColor
   clrBlack = 0 '&H0, Black, RGB(0, 0, 0), #000000, clrBlack = 0x0
   clrBlue = 16711680 '&HFF0000, Blue, RGB(0, 0, 255), #0000FF, clrBlue = 0xFF0000
   clrBrown = 13209&
   clrBrown1 = 2763429 '&H2A2AA5, RGB(165, 42, 42), #A52A2A, Unnamed color = 0x2A2AA5
   clrCrimson = 3937500 '&H3C14DC, RGB(220, 20, 60), #DC143C, Unnamed color = 0x3C14DC
   clrCyan = &HFFFF00
   clrGold = 52479
   clrGold1 = 55295 '&HD7FF, RGB(255, 215, 0), #FFD700, Unnamed color = 0xD7FF
   clrGreen = 32768 '&H8000, Green, RGB(0, 128, 0), #008000, clrGreen = 0x8000
   clrIndigo = 10040115
   clrIndigo1 = 8519755 '&H82004B, RGB(75, 0, 130), #4B0082, Unnamed color = 0x82004B
   clrIvory = 15794175 '&HF0FFFF, RGB(255, 255, 240), #FFFFF0, Unnamed color = 0xF0FFFF
   clrLime = 52377
   clrLime1 = 65280 '&HFF00, Bright Green, RGB(0, 255, 0), #00FF00, clrBrightGreen = 0xFF00
   clrMagneta = &HFF00FF
   clrOrange = 26367&
   clrOrange1 = 42495 '&HA5FF, RGB(255, 165, 0), #FFA500, Unnamed color = 0xA5FF
   clrPurple = 8388736 '&H800080, Violet, RGB(128, 0, 128), #800080, clrViolet = 0x800080
   clrRed = 255 '&HFF, Red, RGB(255, 0, 0), #FF0000, clrRed = 0xFF
   clrRose = 13408767
   clrSalmon = 7504122 '&H7280FA, RGB(250, 128, 114), #FA8072, Unnamed color = 0x7280FA
   clrSilver = 12632256 '&HC0C0C0, Gray 25%, RGB(192, 192, 192), #C0C0C0, clrGray25 = 0xC0C0C0
   clrViolet = 8388736
   clrViolet1 = 15631086 '&HEE82EE, RGB(238, 130, 238), #EE82EE, Unnamed color = 0xEE82EE
   clrWhite = 16777215 '&HFFFFFF, White, RGB(255, 255, 255), #FFFFFF, clrWhite = 0xFFFFFF
   clrYellow = 65535 '&HFFFF, Yellow, RGB(255, 255, 0), #FFFF00, clrYellow = 0xFFFF
End Enum
#End If

Private Enum eEditMsg
   EM_GETFIRSTVISIBLELINE = &HCE
   EM_GETLINECOUNT = &HBA
   EM_LINEFROMCHAR = &HC9
   EM_LINEINDEX = &HBB
   EM_LINELENGTH = &HC1
End Enum

Private Enum eMsg
   ALL_MESSAGES = -1
   WM_LBUTTONDOWN = &H201&
   WM_RBUTTONDOWN = &H204&
   WM_VSCROLL = &H115&
   WM_HSCROLL = &H114&
   WM_PASTE = &H302&
End Enum

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, lParam As Any) As Long

Private mbIsRunTime As Boolean

'Default Property Values:
Private Const m_def_ColourStrings = clrBrown1 '&HC000C0    ' purple
Private Const m_def_ColourOperator = clrRed '&HFF&      ' red
Private Const m_def_ColourKeyWord = clrIndigo '&HFF0000    ' blue
Private Const m_def_ColourComment = clrGreen '&H8000&     ' green

Private Const COMMENT_IDENTIFER = "'"           ' comment line char

' default keyword assignments
Private Const m_def_ceBoldWords As String = "*Do*Loop**If*Then*Else*End*Error*Exit*Resume*For*Next*Call*Dim*Sub*Function*Set*True*False*Case*Select*Private Const*ReDim*With*"
Private Const m_def_ceOperators As String = "*Not*And*Or*In*To*Nothing*Xor*Err*"
Private Const m_def_ceKeyWords As String = "*Abs*AddressOf*Array*As*Asc*ByVal*ByRef*Const*Private Const*CreateObject*Else*ElseIf*If*Alias*Base*Begin*Binary*Boolean*Byte*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*Chr*CInt*CLng*Close*Compare*Private Const*CSng*CStr*Currency*CVar*CVErr" & _
                  "*Day*Decimal" & "*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Friend*Function*Get*GoSub*GoTo" & _
                  "*Hex*If*Imp*Input*Input*InStr*InStrRev*Integer*Is*LBound*Left*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Mid*New*Next" & _
                  "*Object*On*Open*Option*Output*Print*Private*Property*Public*ReDim*Resume*Return*Replace*Right*Select*Set*Single*Spc*Split*Static*String*Stop*Sub*Tab*Then*Then*Time*True*Type" & _
                  "*UBound*Unlock*Variant*WEnd*WScript*While*With*MsgBox*Now*InputBox*Len*Sleep*Trim*RTrim*LTrim*LCase*UCase*Until*vbCrLf*vbLf*vbCr*VB*Menu*"

Private Const m_def_NormaliseCase As Boolean = True
Private Const m_def_ForeColor As Long = 0
Private Const m_def_BackStyle As Long = 0
Private Const m_def_SyntaxColouring As Boolean = True
Private Const m_def_ProcessStrings As Boolean = True
Private Const m_def_ItalicComments As Boolean = False
Private Const m_def_BoldSelectedKeyWords As Boolean = False
Private Const m_def_WordWrap As Boolean = False
Private Const m_def_LineNumbers As Boolean = True
Private Const m_def_SelStart As Long = 0
Private Const m_def_SelLength As Long = 0
Private Const m_def_SelText As String = vbNullString

'Property Variables:
Private m_ColourStrings         As eColor
Private m_ColourOperator        As eColor
Private m_ColourKeyWord         As eColor
Private m_ColourComment         As eColor
Private m_ProcessStrings        As Boolean
Private m_ItalicComments        As Boolean
Private m_BoldSelectedKeyWords  As Boolean
Private m_WordWrap              As Boolean
Private m_LineNumbers           As Boolean
Private m_SelStart              As Long
Private m_SelLength             As Long
Private m_SelText               As String
Private m_ceBoldWords           As String
Private m_ceOperators           As String
Private m_ceKeyWords            As String
Private m_NormaliseCase         As Boolean
Private m_ForeColor             As Long
Private m_BackStyle             As Integer
Private m_SyntaxColouring       As Boolean
Private bDirty                  As Boolean
Private stexttmp                As String

'rgb values for the long to rgb conversion
Private RGBRed1                 As Long
Private RGBBlue1                As Long
Private RGBGreen1               As Long
Private RGBRed2                 As Long
Private RGBBlue2                As Long
Private RGBGreen2               As Long
Private RGBRed3                 As Long
Private RGBBlue3                As Long
Private RGBGreen3               As Long
Private RGBRed4                 As Long
Private RGBBlue4                As Long
Private RGBGreen4               As Long
Private RGBRed5                 As Long
Private RGBBlue5                As Long
Private RGBGreen5               As Long

' other private variables
Private RaiseEvents         As Boolean
Private lLineTracker        As Long
Private mWndProcOrg         As Long
Private mHWndSubClassed     As Long
Private bScrolling          As Boolean

'Event Declarations:
Public Event VScroll()
Public Event HScroll()
Public Event Change() 'MappingInfo=RTB,RTB,-1,Change
Public Event SelChange() 'MappingInfo=RTB,RTB,-1,SelChange
Public Event Click() 'MappingInfo=RTB,RTB,-1,Click
Public Event DblClick() 'MappingInfo=RTB,RTB,-1,DblClick
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=RTB,RTB,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=RTB,RTB,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=RTB,RTB,-1,KeyUp
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTB,RTB,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTB,RTB,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTB,RTB,-1,MouseUp

Public Property Get BackColor() As eColor
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."

   BackColor = RTB.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As eColor)

   RTB.BackColor() = New_BackColor
   PropertyChanged "BackColor"

End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."

   ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)

   m_ForeColor = New_ForeColor
   PropertyChanged "ForeColor"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

   Enabled = RTB.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

   RTB.Enabled() = New_Enabled
   PropertyChanged "Enabled"

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

   Set Font = RTB.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

   Set RTB.Font = New_Font
   PropertyChanged "Font"

End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."

   BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)

   m_BackStyle = New_BackStyle
   PropertyChanged "BackStyle"

End Property

Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

   BorderStyle = RTB.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)

   RTB.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"

End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."

   RTB.Refresh

End Sub

Private Sub RTB_Click()
   'MsgBox "RTB_Click"
   
   RaiseEvent Click

End Sub

Private Sub RTB_DblClick()
   'MsgBox "RTB_DblClick"
   RaiseEvent DblClick

End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)

' Original code by ChiefRedBull from www.VisualBasicForum.com
   'MsgBox "RTB_KeyDown"
   On Error Resume Next

      Dim lCursor             As Long
      Dim lSelectLen          As Long
      Dim lStart              As Long
      Dim lFinish             As Long
      Dim lLocalTracker       As Long

      'If KeyCode = vbKeyA And Shift = 2 Then
         'Debug.Assert 0
      'End If
      lLocalTracker = CurrentLineNumber

      
      Call WriteLineNumbers

      ' check for Ctrl+Y, this is the delete line shortcut
      If KeyCode = vbKeyY And Shift = 2 Then
         ' delete the current line...
         DeleteCurrentLine
         ' null the keypress to get rid of any 'Y' characters..
         KeyCode = 0
      End If

      ' to handle delete being pressed...
      If KeyCode = 8 Then
         If lLocalTracker <> lLineTracker Then
            lLineTracker = CurrentLineNumber
            bDirty = False
         End If
      End If

      ' reset the line tracker after the del check
      lLineTracker = CurrentLineNumber

      If KeyCode = vbKeyTab Then
         'RTB.SelText = vbTab
         RTB.SelText = Space$(3)
         KeyCode = 0
      End If

      ' check for text being pasted into the box
      ' with Ctrl-V.. we also call the same sub when a WM_Paste message
      ' has been send to the control...
      If KeyCode = vbKeyV And Shift = 2 Then
         Call DoPaste
         ' null the keypress so we don't get any 'V' characters
         KeyCode = 0
      End If

      If KeyCode = 13 Or _
         KeyCode = vbKeyUp Or _
         KeyCode = vbKeyDown Or _
         KeyCode = 33 Or KeyCode = 34 Then

         ' only color this line if it's been changed
         If bDirty Or KeyCode = 13 And Shift <> 2 Then

            ' store the current cursor pos
            ' and current selection if there is any
            lCursor = RTB.SelStart
            lSelectLen = RTB.SelLength

            ' sure we need to colour the line.. but lets reset its colour first
            ' to be sure we don't screw the colours up..
            Call ResetColours(CurrentLineNumber - 1)

            ' lock the window and lets colour the line
            LockWindowUpdate RTB.hWnd

            lStart = CurrentLineNumber - 1
            lFinish = CurrentLineNumber - 1

            ColourSelection lStart, lFinish

            ' reset the properties
            RTB.SelStart = lCursor
            RTB.SelLength = lSelectLen
            RTB.SelColor = vbBlack
            RTB.SelBold = False
            RTB.SelItalic = False

            ' reset the flag and release the window
            bDirty = False
            LockWindowUpdate 0&

         End If

      ElseIf Not IsControlKey(KeyCode) And Shift <> 2 Then 'NOT KEYCODE...

         ' this section resets the current lines colour to black
         ' once we are finished, then the above section re-colours the line..
         If bDirty = False Then
            ' reset the colours for this line only!
            Call ResetColours(CurrentLineNumber - 1)
            bDirty = True
         End If

      End If

      RaiseEvent KeyDown(KeyCode, Shift)

End Sub ':(? 'Chr$(160)!!On Error Resume still active

Private Sub RTB_KeyPress(KeyAscii As Integer)
   'MsgBox "RTB_KeyPress"
'don't reset colours on ctrl-c, ctrl+A
   If GetAsyncKeyState(VK_CONTROL) = 0 Then
   'If KeyAscii <> 3 And KeyAscii <> 97 Then
      RTB.SelColor = vbBlack
      RTB.SelBold = False
      RTB.SelItalic = False
   'End If
   End If

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub RTB_KeyUp(KeyCode As Integer, Shift As Integer)
   'MsgBox "RTB_KeyUp"
   
   Call WriteLineNumbers
   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   'MsgBox "RTB_MouseDown"
   
   Call WriteLineNumbers
   RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub RTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   'MsgBox "RTB_MouseUp"
   RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Public Property Get SyntaxColouring() As Boolean

   SyntaxColouring = m_SyntaxColouring

End Property

Public Property Let SyntaxColouring(ByVal New_SyntaxColouring As Boolean)

   m_SyntaxColouring = New_SyntaxColouring
   PropertyChanged "SyntaxColouring"

End Property

Private Sub UserControl_Initialize()
   
   RaiseEvents = True
   'subclassControl RTB

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

   m_ForeColor = m_def_ForeColor
   m_BackStyle = m_def_BackStyle
   m_SyntaxColouring = m_def_SyntaxColouring
   bDirty = True
   m_NormaliseCase = m_def_NormaliseCase
   m_ceBoldWords = m_def_ceBoldWords
   m_ceOperators = m_def_ceOperators
   m_ceKeyWords = m_def_ceKeyWords
   m_SelStart = m_def_SelStart
   m_SelLength = m_def_SelLength
   m_SelText = m_def_SelText
   m_LineNumbers = m_def_LineNumbers
   m_WordWrap = m_def_WordWrap
   m_BoldSelectedKeyWords = m_def_BoldSelectedKeyWords
   m_ItalicComments = m_def_ItalicComments
   m_ProcessStrings = m_def_ProcessStrings
   m_ColourOperator = m_def_ColourOperator
   m_ColourKeyWord = m_def_ColourKeyWord
   m_ColourComment = m_def_ColourComment
   m_ColourStrings = m_def_ColourStrings

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   On Error Resume Next

      Err.Clear
      mbIsRunTime = Ambient.UserMode
      If Err.Number Then
         mbIsRunTime = False
         Err.Clear
      End If

      'Start subclassing to detect 'mouse out' for the splitterbar
      StartSubclassing
   On Error GoTo 0
   RTB.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
   m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
   RTB.Enabled = PropBag.ReadProperty("Enabled", True)
   Set RTB.Font = PropBag.ReadProperty("Font", Ambient.Font)
   m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
   RTB.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   m_SyntaxColouring = PropBag.ReadProperty("SyntaxColouring", m_def_SyntaxColouring)
   RTB.Text = PropBag.ReadProperty("Text", "")
   m_NormaliseCase = PropBag.ReadProperty("NormaliseCase", m_def_NormaliseCase)
   m_ceBoldWords = PropBag.ReadProperty("ceBoldWords", m_def_ceBoldWords)
   m_ceOperators = PropBag.ReadProperty("ceOperators", m_def_ceOperators)
   m_ceKeyWords = PropBag.ReadProperty("ceKeyWords", m_def_ceKeyWords)
   m_SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
   m_SelLength = PropBag.ReadProperty("SelLength", m_def_SelLength)
   m_SelText = PropBag.ReadProperty("SelText", m_def_SelText)
   m_LineNumbers = PropBag.ReadProperty("LineNumbers", m_def_LineNumbers)
   m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
   RTB.HideSelection = PropBag.ReadProperty("HideSelection", False)
   m_BoldSelectedKeyWords = PropBag.ReadProperty("BoldSelectedKeyWords", m_def_BoldSelectedKeyWords)
   m_ItalicComments = PropBag.ReadProperty("ItalicComments", m_def_ItalicComments)
   m_ProcessStrings = PropBag.ReadProperty("ProcessStrings", m_def_ProcessStrings)
   m_ColourOperator = PropBag.ReadProperty("ColourOperator", m_def_ColourOperator)
   m_ColourKeyWord = PropBag.ReadProperty("ColourKeyWord", m_def_ColourKeyWord)
   m_ColourComment = PropBag.ReadProperty("ColourComment", m_def_ColourComment)
   m_ColourStrings = PropBag.ReadProperty("ColourStrings", m_def_ColourStrings)

   picLineNumbers.Visible = m_LineNumbers
   Call UserControl_Resize

   ' split the long values to rgb sub vals
   SplitRGB m_ColourStrings, RGBRed4, RGBGreen4, RGBBlue4
   SplitRGB m_ColourOperator, RGBRed2, RGBGreen2, RGBBlue2
   SplitRGB m_ColourKeyWord, RGBRed1, RGBGreen1, RGBBlue1
   SplitRGB m_ColourComment, RGBRed5, RGBGreen5, RGBBlue5

End Sub

Private Sub UserControl_Resize()

   With RTB

      .Height = UserControl.ScaleHeight
      .Top = UserControl.ScaleTop
      If m_LineNumbers = True Then ':(? 'Chr$(160)!!Remove Pleonasm
         .Left = UserControl.ScaleLeft + picLineNumbers.ScaleWidth
      Else 'NOT M_LINENUMBERS...
         .Left = UserControl.ScaleLeft
      End If
      If m_LineNumbers = True Then ':(? 'Chr$(160)!!Remove Pleonasm
         .Width = UserControl.ScaleWidth - picLineNumbers.Width
      Else 'NOT M_LINENUMBERS...
         .Width = UserControl.ScaleWidth
      End If

   End With 'RTB

   With picLineNumbers

      .Height = UserControl.ScaleHeight
      .Top = UserControl.ScaleTop
      .Left = UserControl.ScaleLeft

   End With 'PICLINENUMBERS

   Call WriteLineNumbers

End Sub

Private Sub UserControl_Terminate()

   StopSubclassing

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("BackColor", RTB.BackColor, &H80000005)
   Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
   Call PropBag.WriteProperty("Enabled", RTB.Enabled, True)
   Call PropBag.WriteProperty("Font", RTB.Font, Ambient.Font)
   Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
   Call PropBag.WriteProperty("BorderStyle", RTB.BorderStyle, 1)
   Call PropBag.WriteProperty("SyntaxColouring", m_SyntaxColouring, m_def_SyntaxColouring)
   Call PropBag.WriteProperty("Text", RTB.Text, "")
   Call PropBag.WriteProperty("NormaliseCase", m_NormaliseCase, m_def_NormaliseCase)
   Call PropBag.WriteProperty("ceBoldWords", m_ceBoldWords, m_def_ceBoldWords)
   Call PropBag.WriteProperty("ceOperators", m_ceOperators, m_def_ceOperators)
   Call PropBag.WriteProperty("ceKeyWords", m_ceKeyWords, m_def_ceKeyWords)
   Call PropBag.WriteProperty("SelStart", m_SelStart, m_def_SelStart)
   Call PropBag.WriteProperty("SelLength", m_SelLength, m_def_SelLength)
   Call PropBag.WriteProperty("SelText", m_SelText, m_def_SelText)
   Call PropBag.WriteProperty("LineNumbers", m_LineNumbers, m_def_LineNumbers)
   Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
   Call PropBag.WriteProperty("HideSelection", RTB.HideSelection, False)
   Call PropBag.WriteProperty("BoldSelectedKeyWords", m_BoldSelectedKeyWords, m_def_BoldSelectedKeyWords)
   Call PropBag.WriteProperty("ItalicComments", m_ItalicComments, m_def_ItalicComments)
   Call PropBag.WriteProperty("ProcessStrings", m_ProcessStrings, m_def_ProcessStrings)
   Call PropBag.WriteProperty("ColourOperator", m_ColourOperator, m_def_ColourOperator)
   Call PropBag.WriteProperty("ColourKeyWord", m_ColourKeyWord, m_def_ColourKeyWord)
   Call PropBag.WriteProperty("ColourComment", m_ColourComment, m_def_ColourComment)
   Call PropBag.WriteProperty("ColourStrings", m_ColourStrings, m_def_ColourStrings)

   picLineNumbers.Visible = m_LineNumbers
   Call UserControl_Resize

   ' split the long values to rgb sub vals
   SplitRGB m_ColourStrings, RGBRed4, RGBGreen4, RGBBlue4
   SplitRGB m_ColourOperator, RGBRed2, RGBGreen2, RGBBlue2
   SplitRGB m_ColourKeyWord, RGBRed1, RGBGreen1, RGBBlue1
   SplitRGB m_ColourComment, RGBRed5, RGBGreen5, RGBBlue5

End Sub

Private Sub RTB_Change()
   'MsgBox "RTB_Change"

   If RaiseEvents Then
      Call WriteLineNumbers
      RaiseEvent Change
   End If

End Sub

Private Sub RTB_SelChange()
   'MsgBox "RTB_SelChange"
   'Debug.Print "RTB_SelChange"
   If RaiseEvents Then
      'Debug.Print "RaiseEvents"
      Call WriteLineNumbers
      RaiseEvent SelChange
   End If

End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
Attribute Text.VB_UserMemId = 0

   Text = RTB.Text

End Property

Public Property Let Text(ByVal New_Text As String)

   'MsgBox "Text"
   RTB.Text() = New_Text
   PropertyChanged "Text"

End Property

Public Property Get NormaliseCase() As Boolean

   NormaliseCase = m_NormaliseCase

End Property

Public Property Let NormaliseCase(ByVal New_NormaliseCase As Boolean)

   m_NormaliseCase = New_NormaliseCase
   PropertyChanged "NormaliseCase"

End Property

Public Property Get ceBoldWords() As String

   ceBoldWords = m_ceBoldWords

End Property

Public Property Let ceBoldWords(ByVal New_ceBoldWords As String)

   m_ceBoldWords = New_ceBoldWords
   PropertyChanged "ceBoldWords"

End Property

Public Property Get ceOperators() As String

   ceOperators = m_ceOperators

End Property

Public Property Let ceOperators(ByVal New_ceOperators As String)

   m_ceOperators = New_ceOperators
   PropertyChanged "ceOperators"

End Property

Public Property Get ceKeyWords() As String

   ceKeyWords = m_ceKeyWords

End Property

Public Property Let ceKeyWords(ByVal New_ceKeyWords As String)

   m_ceKeyWords = New_ceKeyWords
   PropertyChanged "ceKeyWords"

End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"

   SelStart = m_SelStart

End Property

Public Property Let SelStart(ByVal New_SelStart As Long)

   If Ambient.UserMode = False Then Err.Raise 387 ':(? 'Chr$(160)!!Expand Structure
   m_SelStart = New_SelStart
   PropertyChanged "SelStart"

End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"

   SelLength = m_SelLength

End Property

Public Property Let SelLength(ByVal New_SelLength As Long)

   If Ambient.UserMode = False Then Err.Raise 387 ':(? 'Chr$(160)!!Expand Structure
   m_SelLength = New_SelLength
   PropertyChanged "SelLength"

End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"

   SelText = m_SelText

End Property

Public Property Let SelText(ByVal New_SelText As String)

   If Ambient.UserMode = False Then Err.Raise 387 ':(? 'Chr$(160)!!Expand Structure
   m_SelText = New_SelText
   PropertyChanged "SelText"

End Property

Public Sub InsertString(InsertString As String)

   RTB.SelText = InsertString

End Sub

Public Property Get LineNumbers() As Boolean

   LineNumbers = m_LineNumbers

End Property

Public Property Let LineNumbers(ByVal New_LineNumbers As Boolean)

   m_LineNumbers = New_LineNumbers
   PropertyChanged "LineNumbers"
   picLineNumbers.Visible = m_LineNumbers
   Call UserControl_Resize

End Property

Public Property Get WordWrap() As Boolean

   WordWrap = m_WordWrap

End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)

   m_WordWrap = New_WordWrap
   PropertyChanged "WordWrap"
   If m_WordWrap = True Then ':(? 'Chr$(160)!!Remove Pleonasm
      RTB.RightMargin = RTB.Width - 250
   Else 'NOT M_WORDWRAP...
      RTB.RightMargin = 999999
   End If
   Call WriteLineNumbers

End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns/sets a value that specifies if the selected item remains highlighted when a control loses focus."

   HideSelection = RTB.HideSelection

End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)

   RTB.HideSelection() = New_HideSelection
   PropertyChanged "HideSelection"

End Property

Public Property Get BoldSelectedKeyWords() As Boolean

   BoldSelectedKeyWords = m_BoldSelectedKeyWords

End Property

Public Property Let BoldSelectedKeyWords(ByVal New_BoldSelectedKeyWords As Boolean)

   m_BoldSelectedKeyWords = New_BoldSelectedKeyWords
   PropertyChanged "BoldSelectedKeyWords"

End Property

Public Property Get ItalicComments() As Boolean

   ItalicComments = m_ItalicComments

End Property

Public Property Let ItalicComments(ByVal New_ItalicComments As Boolean)

   m_ItalicComments = New_ItalicComments
   PropertyChanged "ItalicComments"

End Property

Public Property Get ProcessStrings() As Boolean

   ProcessStrings = m_ProcessStrings

End Property

Public Property Let ProcessStrings(ByVal New_ProcessStrings As Boolean)

   m_ProcessStrings = New_ProcessStrings
   PropertyChanged "ProcessStrings"

End Property

Public Property Get ColourOperator() As eColor

   ColourOperator = m_ColourOperator

End Property

Public Property Let ColourOperator(ByVal New_ColourOperator As eColor)

   m_ColourOperator = New_ColourOperator
   PropertyChanged "ColourOperator"

End Property

Public Property Get ColourKeyWord() As eColor

   ColourKeyWord = m_ColourKeyWord

End Property

Public Property Let ColourKeyWord(ByVal New_ColourKeyWord As eColor)

   m_ColourKeyWord = New_ColourKeyWord
   PropertyChanged "ColourKeyWord"

End Property

Public Property Get ColourComment() As eColor

   ColourComment = m_ColourComment

End Property

Public Property Let ColourComment(ByVal New_ColourComment As eColor)

   m_ColourComment = New_ColourComment
   PropertyChanged "ColourComment"

End Property

Public Property Get ColourStrings() As eColor

   ColourStrings = m_ColourStrings

End Property

Public Property Let ColourStrings(ByVal New_ColourStrings As eColor)

   m_ColourStrings = New_ColourStrings
   PropertyChanged "ColourStrings"

End Property

Public Property Get hWnd() As Long

   hWnd = UserControl.hWnd

End Property

Public Sub ColourSelection(lStartLine As Long, lEndLine As Long)

' go thru the rtb line by line, instead of the traditional way of selecting
' each keyword individually, we will select the entire line, then write
' back to the SelRTF property..

' this does not need to be 'as' fast as the ColourEntireRTB sub, but..
' still needs to be reasonable as this will process blocks of code
' for WM_Paste messages in the RTB..

' Karl Durrance, Dec 2002

' the lStartLine and lEndLine values are zero based..

   Dim x                       As Long
   Dim i                       As Long
   Dim lCurLineStart           As Long
   Dim lCurLineEnd             As Long
   Dim sLineText               As String
   Dim sLineTextRTF            As String
   Dim lnglength               As Long
   Dim nQuoteEnd               As Long
   Dim sCurrentWord            As String
   Dim sChar                   As String
   Dim nWordPos                As Long
   Dim lColour                 As Long
   Dim lLastBreak              As Long
   Dim sBoldStart              As String
   Dim sBoldEnd                As String
   Dim bDone                   As Boolean
   Dim lLineOffset             As Long
   Dim lStartRTFCode           As Long
   Dim stmpstring              As String

   With RTB

      For i = lStartLine To lEndLine

         ' get the details for this line
         lCurLineStart = SendMessage(.hWnd, EM_LINEINDEX, i, 0&)
         lnglength = SendMessage(.hWnd, EM_LINELENGTH, lCurLineStart, 0)

         ' if the line actually has some data in it then we'll process it..
         If lCurLineStart >= 0 And lnglength > 0 Then

            ' select the entire line
            .SelStart = lCurLineStart
            .SelLength = lnglength

            If lCurLineStart = 1 Then lCurLineStart = 0 ':(? 'Chr$(160)!!Expand Structure

            ' get the selected text.. assign to a variable
            sLineText = .SelText

            ' fix up any rtf problems now.. like "\{}"..
            If InStr(1, sLineText, "\") Or InStr(1, sLineText, "{") Or InStr(1, sLineText, "}") Then
               sLineText = Replace$(sLineText, "\", "\\")
               sLineText = Replace$(sLineText, "{", "\{")
               sLineText = Replace$(sLineText, "}", "\}")
               lnglength = Len(sLineText)
            End If

            ' check for comment identifier at the start of the line
            If Left$(LTrim$(sLineText), 1) = "'" Then
               ' colour the lines that are complete comments like this
               ' beats messing around with the RTB codes..
               ' there is no speed loss since the line is already selected..
               .SelColor = m_ColourComment
               If m_ItalicComments = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                  .SelItalic = True
               End If
            Else 'NOT LEFT$(LTRIM$(SLINETEXT),...

               lLastBreak = 1
               For x = 1 To Len(sLineText)

                  sChar = Mid$(sLineText, x, 1)
                  bDone = False

                  Select Case sChar

                  Case COMMENT_IDENTIFER

                     ' write the colours now!
                     If Len(sLineTextRTF) > 0 Then

                        .SelRTF = "{{\colortbl;\red" & RGBRed1 & "\green" & _
                                  RGBGreen1 & "\blue" & RGBBlue1 & ";\red" & RGBRed2 & _
                                  "\green" & RGBGreen2 & "\blue" & RGBBlue2 & ";\red" & _
                                  RGBRed3 & "\green" & RGBGreen3 & "\blue" & RGBBlue3 & _
                                  ";\red" & RGBRed4 & "\green" & RGBGreen4 & "\blue" & _
                                  RGBBlue4 & ";\red" & RGBRed5 & "\green" & RGBGreen5 _
                                  & "\blue" & RGBBlue5 & ";}" & sLineTextRTF & "\I0\B0}\par"

                     End If

                     ' comment, colour the rest of the line
                     ' these can be done the slower way..
                     ' with no real time loss..
                     ' these are rarer than standard comments...
                     .SelStart = lCurLineStart + x - 1
                     .SelLength = (lnglength + 2) - x
                     .SelColor = m_ColourComment
                     If m_ItalicComments = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                        .SelItalic = True
                     End If
                     ' set the flag so we don't colour the line again
                     bDone = True
                     Exit For '>---> Next

                  Case Chr$(34)

                     ' Find the end and reset the for loop
                     nQuoteEnd = InStr(x + 1, sLineText, Chr$(34), vbBinaryCompare)
                     If nQuoteEnd = 0 Then nQuoteEnd = Len(sLineText) ':(? 'Chr$(160)!!Expand Structure

                     If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure

                     If m_ProcessStrings = True Then ':(? 'Chr$(160)!!Remove Pleonasm

                        ' assign the colour codes to the string..
                        stmpstring = "{\cf4" & Mid$(sLineText, x, (nQuoteEnd - x) + 1) & "\cf0}"
                        sLineTextRTF = Replace$(sLineTextRTF, Mid$(sLineText, x, (nQuoteEnd - x) + 1), "{\cf4" & Mid$(sLineText, x, (nQuoteEnd - x) + 1) & "\cf0}")
                        lLineOffset = lLineOffset + 10

                     End If

                     x = nQuoteEnd ':(? 'Chr$(160)!!Modifies active For-Variable

                  Case "a" To "z", "A" To "Z", "_"
                     ' alphanumeric, non string or comment..
                     sCurrentWord = sCurrentWord & sChar
                     ' if we are at the end of a line with no vbCrLf then
                     ' call the colour routine directly so we don't miss
                     ' the last word in the line...
                     If x = Len(sLineText) Then GoTo ColourWord ':(? 'Chr$(160)!!Expand Structure

                  Case Else
                     ' should be a word sep char.. so we could have a word!

ColourWord:

                     If sCurrentWord <> "" Then

                        nWordPos = InStr(1, m_ceKeyWords & m_ceOperators, "*" & sCurrentWord & "*", vbTextCompare)

                        If nWordPos > 0 Then
                           ' this word is a keyword, set the colour
                           If nWordPos > Len(m_ceKeyWords) Then
                              lColour = 2
                           Else 'NOT NWORDPOS...
                              lColour = 1
                           End If

                           ' check if we need to bold the word..
                           If m_BoldSelectedKeyWords = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                              If InStr(1, m_ceBoldWords, "*" & sCurrentWord & "*", vbTextCompare) Then
                                 sBoldStart = "\b1"
                                 sBoldEnd = "\b0"
                              Else 'NOT INSTR(1,...
                                 sBoldStart = ""
                                 sBoldEnd = ""
                              End If
                           End If

                           ' reset the case of the keyword if required...
                           If m_NormaliseCase = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                              sCurrentWord = Mid$(m_ceKeyWords & m_ceOperators, InStr(1, LCase$(m_ceKeyWords & m_ceOperators), "*" & LCase$(sCurrentWord) & "*", vbBinaryCompare) + 1, Len(sCurrentWord))
                           End If

                           ' now colour the word with the rtf codes
                           ' use the custom replaceword function, start at the last breakpoint
                           ' only colour one copy of the word..
                           If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure
                           sLineTextRTF = ReplaceFullWord$(sLineTextRTF, sCurrentWord, "{\cf" & lColour & sBoldStart & sCurrentWord & sBoldEnd & "\cf0}", lLastBreak + lLineOffset, 1, vbTextCompare)
                           'assign the offset because of the RTF codes..
                           lLineOffset = lLineOffset + 10 + IIf(Len(sBoldStart) > 0, 6, 0)

                        End If

                        ' reset the word to nothing
                        sCurrentWord = ""

                     End If

                     lLastBreak = x

                  End Select

               Next x

               If sLineTextRTF <> "" And bDone = False Then

                  .SelRTF = "{{\colortbl;\red" & RGBRed1 & "\green" & _
                            RGBGreen1 & "\blue" & RGBBlue1 & ";\red" & RGBRed2 & _
                            "\green" & RGBGreen2 & "\blue" & RGBBlue2 & ";\red" & _
                            RGBRed3 & "\green" & RGBGreen3 & "\blue" & RGBBlue3 & _
                            ";\red" & RGBRed4 & "\green" & RGBGreen4 & "\blue" & _
                            RGBBlue4 & ";\red" & RGBRed5 & "\green" & RGBGreen5 _
                            & "\blue" & RGBBlue5 & ";}" & sLineTextRTF & "\I0\B0}\par"

               End If

               sLineTextRTF = ""
               lLineOffset = 0

            End If

         End If

      Next i

   End With 'RTB

End Sub

Public Sub ColourEntireRTB(Optional sThisText As String, Optional bUseThisText As Boolean)

' This is for an entire colour of the RTB.. like on load..
' this out performs the line by line methods because we process
' the entire script in memory..

' the structure is basically the same as the ColourSelection sub
' but we write to the TextRTF property at the end instead..
' and do all the line processing in memory

' this obviously clears the entire contents of the rtb..

' Karl Durrance Dec 2002

   Dim x                       As Long
   Dim i                       As Long
   Dim lCurLineStart           As Long
   Dim lCurLineEnd             As Long
   Dim sLineText               As String
   Dim sLineTextRTF            As String
   Dim sAllTextRTF             As String
   Dim lnglength               As Long
   Dim nQuoteEnd               As Long
   Dim sCurrentWord            As String
   Dim sChar                   As String
   Dim nWordPos                As Long
   Dim lColour                 As Long
   Dim lLastBreak              As Long
   Dim sBoldStart              As String
   Dim sBoldEnd                As String
   Dim sItalicStart            As String
   Dim sItalicEnd              As String
   Dim bDone                   As Boolean
   Dim lLineOffset             As Long
   Dim lStartRTFCode           As Long
   Dim stmpstring              As String
   Dim sBuffer                 As String
   Dim asBuffer()              As String
   Dim bForce                  As Boolean
   Dim objAllRTFString         As New CString
   Dim objFinalConcat          As New CString
   Dim sTextRTF                As String

   With RTB

      If m_ItalicComments = True Then ':(? 'Chr$(160)!!Remove Pleonasm
         ' set the RTF italic code because we have it turned on..
         sItalicStart = "\I1"
         sItalicEnd = "\I0"
      End If
      
      If bUseThisText Then
         sBuffer = sThisText
      Else
         sBuffer = .Text
      End If
      asBuffer = Split(sBuffer, vbCrLf)

      ' set the text buffer for the CString class..
      ' we'll set the size initially to triple the size of the script
      ' in plain text.. this is pretty big, but will speed up execution
      ' because memory won't need to be reallocated during load..
      ' we will release the extra memory at the end by resetting the buffer..
      objAllRTFString.SetBufferSize Len(sBuffer) * 3
      objFinalConcat.SetBufferSize Len(sBuffer) * 3

      For i = LBound(asBuffer) To UBound(asBuffer)

         ' get the selected text.. assign to a variable for readability
         sLineText = asBuffer(i)

         ' fix up any rtf problems now.. like "\{}"..
         If InStr(1, sLineText, "\") Or InStr(1, sLineText, "{") Or InStr(1, sLineText, "}") Then
            sLineText = Replace$(sLineText, "\", "\\")
            sLineText = Replace$(sLineText, "{", "\{")
            sLineText = Replace$(sLineText, "}", "\}")
         End If

         ' check for comment identifier at the start of the line
         If Left$(LTrim$(sLineText), 1) = "'" Then
            sLineTextRTF = "{\cf5" & sItalicStart & sLineText & "\cf0" & sItalicEnd & "}"
            objAllRTFString.Append sLineTextRTF & "\par" & vbCrLf
            ' reset the variables now.. we are done for this line..
            sLineTextRTF = ""
            lLineOffset = 0
         Else 'NOT LEFT$(LTRIM$(SLINETEXT),...

            lLastBreak = 1
            For x = 1 To Len(sLineText)

               sChar = Mid$(sLineText, x, 1)
               bDone = False

               Select Case sChar

               Case COMMENT_IDENTIFER

                  If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure

                  ' comment, colour the rest of the line
                  sLineTextRTF = Mid$(sLineTextRTF, 1, (x + lLineOffset) - 1) & "{\cf5" & sItalicStart & Mid$(sLineTextRTF, x + lLineOffset) & "\cf0" & sItalicEnd & "}"
                  Exit For '>---> Next

               Case Chr$(34)

                  ' Find the end and reset the for loop
                  nQuoteEnd = InStr(x + 1, sLineText, Chr$(34), vbBinaryCompare)
                  If nQuoteEnd = 0 Then nQuoteEnd = Len(sLineText) ':(? 'Chr$(160)!!Expand Structure

                  If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure

                  If m_ProcessStrings = True Then ':(? 'Chr$(160)!!Remove Pleonasm

                     ' assign the colour codes to the string..
                     stmpstring = "{\cf4" & Mid$(sLineText, x, (nQuoteEnd - x) + 1) & "\cf0}"
                     sLineTextRTF = Replace$(sLineTextRTF, Mid$(sLineText, x, (nQuoteEnd - x) + 1), "{\cf4" & Mid$(sLineText, x, (nQuoteEnd - x) + 1) & "\cf0}")
                     lLineOffset = lLineOffset + 10

                  End If

                  x = nQuoteEnd ':(? 'Chr$(160)!!Modifies active For-Variable

               Case "a" To "z", "A" To "Z", "_"
                  ' alphanumeric, non string or comment..
                  sCurrentWord = sCurrentWord & sChar
                  ' if we are at the end of a line with no vbCrLf then
                  ' call the colour routine directly so we don't miss
                  ' the last word in the line...
                  If x = Len(sLineText) Then GoTo ColourWord ':(? 'Chr$(160)!!Expand Structure

               Case Else
                  ' should be a word sep char.. so we could have a word!

                  ' this tag is basically to handle the last word on a line
                  ' just incase it needs colouring we call the ColourWord tag directly..
ColourWord:

                  If sCurrentWord <> "" Then

                     nWordPos = InStr(1, m_ceKeyWords & m_ceOperators, "*" & sCurrentWord & "*", vbTextCompare)

                     If nWordPos > 0 Then
                        ' this word is a keyword, set the colour
                        If nWordPos > Len(m_ceKeyWords) Then
                           lColour = 2
                        Else 'NOT NWORDPOS...
                           lColour = 1
                        End If

                        ' check if we need to bold the word..
                        If m_BoldSelectedKeyWords = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                           If InStr(1, m_ceBoldWords, "*" & sCurrentWord & "*", vbTextCompare) Then
                              sBoldStart = "\b1"
                              sBoldEnd = "\b0"
                           Else 'NOT INSTR(1,...
                              sBoldStart = ""
                              sBoldEnd = ""
                           End If
                        End If

                        ' reset the case of the keyword if required...
                        If m_NormaliseCase = True Then ':(? 'Chr$(160)!!Remove Pleonasm
                           sCurrentWord = Mid$(m_ceKeyWords & m_ceOperators, InStr(1, LCase$(m_ceKeyWords & m_ceOperators), "*" & LCase$(sCurrentWord) & "*", vbBinaryCompare) + 1, Len(sCurrentWord))
                        End If

                        ' now colour the word with the rtf codes
                        ' use the custom replaceword function, start at the last breakpoint
                        ' only colour one copy of the word..
                        If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure
                        sLineTextRTF = ReplaceFullWord$(sLineTextRTF, sCurrentWord, "{\cf" & lColour & sBoldStart & sCurrentWord & sBoldEnd & "\cf0}", lLastBreak + lLineOffset, 1, vbTextCompare)
                        'assign the offset because of the RTF codes..
                        lLineOffset = lLineOffset + 10 + IIf(Len(sBoldStart) > 0, 6, 0)

                     End If

                     ' reset the word to nothing
                     sCurrentWord = ""

                  End If

                  lLastBreak = x

               End Select

            Next x

            If sLineTextRTF = "" Then sLineTextRTF = sLineText ':(? 'Chr$(160)!!Expand Structure

            ' for LARGE strings, concatenation is a pain..
            ' so we will replace with the fast CString class
            objAllRTFString.Append sLineTextRTF & "\par" & vbCrLf

            sLineTextRTF = ""
            lLineOffset = 0

         End If

      Next i

      sAllTextRTF = objAllRTFString.Value

      ' once again, use the faster CString class
      ' for BIG scripts, this can save up to a second!!

      objFinalConcat.Append "{{\colortbl;\red" & RGBRed1 & "\green" & RGBGreen1 & _
                            "\blue" & RGBBlue1 & ";\red" & RGBRed2 & "\green" & RGBGreen2 & "\blue" & _
                            RGBBlue2 & ";\red" & RGBRed3 & "\green" & RGBGreen3 & "\blue" & RGBBlue3 _
                            & ";\red" & RGBRed4 & "\green" & RGBGreen4 & "\blue" & RGBBlue4 & ";\red" _
                            & RGBRed5 & "\green" & RGBGreen5 & "\blue" & RGBBlue5 & ";}"

      objFinalConcat.Append sAllTextRTF
      objFinalConcat.Append "\I0\B0}\par"

      ' reset the buffer size to the amount of characters.
      objFinalConcat.SetBufferSize objFinalConcat.Length

      ' clear the buffer to release memory now..
      objAllRTFString.SetBufferSize 0, True

      sTextRTF = objFinalConcat.Value

      ' clear the buffer to release memory now..
      objFinalConcat.SetBufferSize 0, True

      ' we are finished...write the full set of RTF to the TextRTF property of the RTB!!
      .TextRTF = "" ' clear the rtb box of all contents before writing the the value.
      .TextRTF = sTextRTF

   End With 'RTB

   Set objFinalConcat = Nothing
   Set objAllRTFString = Nothing

End Sub

Private Function ReplaceFullWord(Source As String, Find As String, ReplaceStr As String, _
                                 Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
                                 Optional Compare As VbCompareMethod = vbBinaryCompare) As String

   Dim findLen             As Long
   Dim replaceLen          As Long
   Dim Index               As Long
   Dim counter             As Long
   Dim charcode            As Long
   Dim replaceIt           As Boolean

   findLen = Len(Find)
   replaceLen = Len(ReplaceStr)

   ' this prevents an endless loop
   If findLen = 0 Then Err.Raise 5 ':(? 'Chr$(160)!!Expand Structure

   If Start < 1 Then Start = 1 ':(? 'Chr$(160)!!Expand Structure
   Index = Start

   ' let's start by assigning the source to the result
   ReplaceFullWord = Source

   Do
      Index = InStr(Index, ReplaceFullWord, Find, Compare)
      If Index = 0 Then Exit Do ':(? 'Chr$(160)!!Expand Structure or consider reversing Condition

      replaceIt = False
      ' check that it is preceded by a punctuation symbol
      If Index > 1 Then
         charcode = Asc(UCase$(Mid$(ReplaceFullWord, Index - 1, 1)))
      Else 'NOT INDEX...
         charcode = 32
      End If
      If charcode < 65 Or charcode > 90 Then
         ' check that it is followed by a punctuation symbol
         charcode = Asc(UCase$(Mid$(ReplaceFullWord, Index + Len(Find), _
                    1)) & " ")
         If charcode < 65 Or charcode > 90 Then
            replaceIt = True
         End If
      End If

      If replaceIt Then
         ' do the replacement
         ReplaceFullWord = Left$(ReplaceFullWord, Index - 1) & ReplaceStr & Mid$ _
                           (ReplaceFullWord, Index + findLen)
         ' skip over the string just added
         Index = Index + replaceLen
         ' increment the replacement counter
         counter = counter + 1
      Else 'REPLACEIT = FALSE
         ' skip over this false match
         Index = Index + findLen
      End If

      ' Note that the Loop Until test will always fail if Count = -1
   Loop Until counter = Count

End Function

Public Sub SelectCurrentLine()

   Dim lStart      As Long
   Dim lFinish     As Long

   ' get the line start and end
   lStart = SendMessage(RTB.hWnd, EM_LINEINDEX, CurrentLineNumber - 1, 0&)
   lFinish = SendMessage(RTB.hWnd, EM_LINELENGTH, lStart, 0)

   RTB.SelStart = lStart
   RTB.SelLength = lFinish

End Sub

Public Sub DeleteCurrentLine()

   Dim lStart      As Long
   Dim lFinish     As Long

   LockWindowUpdate RTB.hWnd

   ' select the entire line, then delete the text
   SelectCurrentLine
   RTB.SelText = ""

   ' take the risk.. delete the line with sendkeys.. YUK!
   RTB.SetFocus
   SendKeys "{DEL}", True
   
   LockWindowUpdate 0&

End Sub

Private Sub ResetColours(lLine As Long)

'lLine is zero based!

   Dim lStart      As Long
   Dim lFinish     As Long
   Dim lCursor     As Long
   Dim lSelectLen  As Long

   LockWindowUpdate RTB.hWnd

   ' get the line start and end
   lStart = SendMessage(RTB.hWnd, EM_LINEINDEX, lLine, 0&)
   lFinish = SendMessage(RTB.hWnd, EM_LINELENGTH, lStart, 0)

   lCursor = RTB.SelStart
   lSelectLen = RTB.SelLength

   RTB.SelStart = lStart
   RTB.SelLength = lFinish
   RTB.SelColor = vbBlack
   RTB.SelBold = False
   RTB.SelItalic = False

   RTB.SelStart = lCursor
   RTB.SelLength = lSelectLen

   LockWindowUpdate 0&

End Sub

Private Sub WriteLineNumbers()

' write the line numbers in the picture box..
' nice and quick way with the Print method.., ie.. no fancy crap, this works nicely.
' only print from the bounds of the top of the page to the bottom.. this way it
' takes no time at all!!
   If Not m_LineNumbers Then
      Exit Sub
   End If
   
   Dim x           As Long
   Dim lStart      As Long
   Dim FontHeight  As Long
   Dim lFinish     As Long

   lStart = SendMessage(RTB.hWnd, EM_GETFIRSTVISIBLELINE, 0, 0) + 1

   picLineNumbers.Cls
   picLineNumbers.Font = RTB.Font.Name
   picLineNumbers.FontSize = RTB.Font.Size
   picLineNumbers.ForeColor = vbBlue
   picLineNumbers.BackColor = &H80000013

   FontHeight = picLineNumbers.TextHeight("1.")

   lFinish = (RTB.Height / FontHeight) + lStart
   If lFinish > LineCount Then lFinish = LineCount ':(? 'Chr$(160)!!Expand Structure

   ' loop from the first visible line in the rtb to the end of the page
   For x = lStart To lFinish
      picLineNumbers.Print x & "."
   Next x

End Sub

Private Sub DoPaste()

' Original code by ChiefRedBull from www.VisualBasicForum.com

   Dim lCursor         As Long
   Dim lStart          As Long
   Dim lFinish         As Long
   Dim sText           As String

   Screen.MousePointer = vbHourglass

   lCursor = RTB.SelStart
   LockWindowUpdate RTB.hWnd
   sText = Clipboard.GetText

   ' the starting line is the line we are currently on..
   lStart = CurrentLineNumber - 1

   RTB.SelText = sText
   lFinish = RTB.GetLineFromChar(RTB.SelStart + RTB.SelLength)

   ColourSelection lStart, lFinish

   ' restore the original values
   RTB.SelStart = lCursor + Len(sText)
   RTB.SelColor = vbBlack

   LockWindowUpdate 0&

   Screen.MousePointer = vbNormal
   RTB.Refresh

End Sub

Private Function IsControlKey(ByVal KeyCode As Long) As Boolean

' Code by ChiefRedBull from www.VisualBasicForum.com

' check if the key is a control key

   Select Case KeyCode
   Case vbKeyLeft, vbKeyRight, vbKeyHome, _
        vbKeyEnd, vbKeyPageUp, vbKeyPageDown, _
        vbKeyShift, vbKeyControl
      IsControlKey = True
   Case Else
      IsControlKey = False
   End Select

End Function

Public Sub LoadFile(sFilePath As String)

' Original Code by ChiefRedBull from www.VisualBasicForum.com

   Dim FileNum     As Long

   Screen.MousePointer = vbHourglass

   'lock the window update so we don't get flicker
   LockWindowUpdate RTB.hWnd
   RaiseEvents = False

   ' load the file
   FileNum = FreeFile
   Open sFilePath For Input As FileNum
   RTB.Text = Input(LOF(FileNum), FileNum)
   Close FileNum

   ' Call the colouring routine
   ' this is destructive!!!
   ColourEntireRTB

   ' reset the cursor postion to the top of the rtb
   RTB.SelStart = 0

   ' write the line numbers
   Call WriteLineNumbers

   ' update the controls view
   RaiseEvents = True
   LockWindowUpdate 0&

   Screen.MousePointer = vbNormal

End Sub

Public Function GetLine(lngLine As Long) As String

   Dim sAllText    As String
   Dim lngindex    As Long
   Dim lnglength   As Long
   Dim x           As Long
   Dim stemp       As String
   Dim sChar       As Long

   sAllText = RTB.Text

   'get the current lines text..
   lngindex = SendMessage(RTB.hWnd, EM_LINEINDEX, lngLine - 1, 0)
   lnglength = SendMessage(RTB.hWnd, EM_LINELENGTH, lngindex, 0) + 2

   stemp = Mid$(sAllText, lngindex + 1, lnglength)

   ' strip any line feed characters as they are going to stuff us up..
   For x = 1 To Len(stemp)

      sChar = Asc(Mid$(stemp, x, 1))

      If Not sChar = 10 And Not sChar = 13 Then
         GetLine = GetLine & Mid$(stemp, x, 1)
      End If

   Next x

End Function

Public Function CurrentWord() As String

' get the current word being typed from bound to bound.

   Dim BreakChrs       As String
   Dim sLineText       As String
   Dim x               As Long
   Dim lStart          As Long
   Dim lLineStart      As Long

   sLineText = GetLine(CurrentLineNumber)
   lStart = CurrentColumnNumber

   ' set the break character criteria for the words..
   BreakChrs = " ,.()<>[]\|:;=/*-+" & _
               Chr$(32) & _
               Chr$(13) & _
               Chr$(10) & _
               Chr$(9) & _
               Chr$(39)

   For x = lStart To 1 Step -1

      If InStr(1, BreakChrs, Mid$(sLineText, x, 1)) Then
         CurrentWord = Mid$(sLineText, x + 1, lStart - x)
         Exit For '>---> Next
      End If

   Next x

   If CurrentWord = "" Then CurrentWord = Mid$(sLineText, 1, lStart) ':(? 'Chr$(160)!!Expand Structure

End Function

Public Function CurrentLineNumber() As Long

' return the current line number in the code window

   CurrentLineNumber = SendMessage(RTB.hWnd, EM_LINEFROMCHAR, ByVal -1, 0&) + 1

End Function

Public Function CurrentColumnNumber() As Long

   Dim lCurLine As Long

   ' Current Line
   lCurLine = 1 + RTB.GetLineFromChar(RTB.SelStart)
   ' Column
   CurrentColumnNumber = SendMessage(RTB.hWnd, EM_LINEINDEX, ByVal lCurLine - 1, 0&)
   CurrentColumnNumber = (RTB.SelStart) - CurrentColumnNumber

End Function

Public Function LineCount() As Long

' return the total line count of the code window

   LineCount = SendMessage(RTB.hWnd, EM_GETLINECOUNT, 0, 0)

End Function

Public Function SaveFile(sFilePath As String) ':(? 'Chr$(160)!!As Variant ?

   RTB.SaveFile sFilePath, rtfText

End Function

Private Sub SplitRGB(ByVal lColor As Long, _
                     ByRef lRed As Long, _
                     ByRef lGreen As Long, _
                     ByRef lBlue As Long)

   lRed = lColor And &HFF
   lGreen = (lColor And &HFF00&) \ &H100&
   lBlue = (lColor And &HFF0000) \ &H10000

End Sub

'======== Subclassing for hiding bar ===============
'
Private Sub StartSubclassing()

#If USE_SUBCLASS Then

   If mbIsRunTime Then
      If sc Is Nothing Then
         Set sc = New CSubclass        'Create a CSubclass instance
         With sc
            Call .AddMsg(WM_VSCROLL, MSG_BEFORE)
            Call .AddMsg(WM_HSCROLL, MSG_BEFORE)
            Call .AddMsg(WM_LBUTTONDOWN, MSG_BEFORE)
            Call .AddMsg(WM_RBUTTONDOWN, MSG_BEFORE)
            Call .AddMsg(WM_PASTE, MSG_BEFORE)
            'Call .AddMsg(WM_CHAR, MSG_BEFORE)
            'Call .AddMsg(WM_KEYDOWN, MSG_BEFORE)
            
            'Call .AddMsg(ALL_MESSAGES, MSG_BEFORE)
            Call .Subclass(RTB.hWnd, Me)
         End With 'SC
      End If
   End If
#End If

End Sub

Private Sub StopSubclassing()

#If USE_SUBCLASS Then

   If mbIsRunTime Then
      If Not sc Is Nothing Then
         Set sc = Nothing
      End If
   End If
   
#End If

End Sub

Private Sub ISubclass_After( _
                            lReturn As Long, ByVal hWnd As Long, _
                            ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub

Private Sub ISubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As eMsg, wParam As Long, lParam As Long)

' SubClassing Sub constructed from example provided by Garrett Sever (The Hand)
' on www.VisualBasicForum.com

' this sub captures the messages and allows us to process them..

   Dim lCurCursor      As Long
   Dim lFirstLine      As Long

   On Local Error Resume Next
   
'   Select Case uMsg
'   Case WM_NCHITTEST, WM_MOUSEFIRST, WM_SETCURSOR, WM_KILLFOCUS, _
'            8270, WM_MOUSEMOVE, WM_TIMER, WM_NCPAINT, WM_ERASEBKGND, WM_IME_SETCONTEXT, _
'            WM_PAINT
'   Case Else
'      Debug.Print eMsgDesc(uMsg)
'      If wParam = VK_TAB Then
'         Debug.Print eMsgDesc(uMsg) & "*** VK_TAB"
'      End If
'   End Select
   
   Select Case uMsg
   Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_VSCROLL, WM_HSCROLL

      If uMsg = WM_VSCROLL Then
         ' write the line numbers on the Vertical Scroll..
         Call WriteLineNumbers
         ' raise the custom scroll event
         RaiseEvent VScroll
      End If

      If uMsg = WM_HSCROLL Then
         ' raise the custom scroll event
         RaiseEvent HScroll
      End If

      ' now be basically need to capture the times we move off a line
      ' and its not coloured.. ie.. on click on the form, scroll etc..
      ' this will only call if the rtb has the dirty flag..

         If bDirty = True Then ':(? 'Chr$(160)!!Remove Pleonasm

            lCurCursor = RTB.SelStart
            LockWindowUpdate RTB.hWnd
            ' colour the dirty line now
            ColourSelection lLineTracker - 1, lLineTracker - 1
            LockWindowUpdate 0&
            ' reset the flag to false
            bDirty = False

            ' reset the caret pos to the place we clicked or left the cursor
            If lCurCursor > 0 Then
               RTB.SelStart = lCurCursor
            End If

         End If

      Case WM_PASTE
      ' when text is being pasted into the control call DoPaste..
      ' not by ctrl-v, but by a msg being sent to the control by SendMessage..
         'Debug.Print "WM_PASTE"
         Call DoPaste
         bHandled = True

'      Case WM_KEYDOWN, WM_CHAR
'         If uMsg = WM_CHAR Then
'            Debug.Print "WM_CHAR"
'         Else
'            Debug.Print "WM_KEYDOWN"
'         End If
'         Debug.Print "wParam=" & wParam
'         If wParam = api.eVirtualKey.VK_TAB Then
'            Debug.Print "VK_TAB"
'         ElseIf wParam = Asc(vbTab) Then
'            Debug.Print "Asc(vbTab)"
'         End If
      End Select
End Sub ':(? 'Chr$(160)!!On Error Resume still active
