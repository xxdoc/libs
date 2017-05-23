VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00D0D0D0&
   Caption         =   "Subclass..."
   ClientHeight    =   7710
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00404080&
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picHeader 
      Height          =   300
      Left            =   -15
      ScaleHeight     =   240
      ScaleWidth      =   8505
      TabIndex        =   11
      Top             =   0
      Width           =   8565
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "######## When.. lReturn. hWnd.... uMsg.... wParam.. lParam.. Message name....... "
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   -15
         TabIndex        =   12
         Top             =   -30
         UseMnemonic     =   0   'False
         Width           =   8580
      End
   End
   Begin VB.PictureBox picMsgSel 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   8550
      ScaleHeight     =   7650
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.CheckBox chkAfter 
         Caption         =   "After original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   3990
         Width           =   1950
      End
      Begin VB.PictureBox picOptAfter 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   240
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   6
         Top             =   4320
         Width           =   2190
         Begin VB.OptionButton optAfter 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1230
         End
         Begin VB.OptionButton optAfter 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkBefore 
         Caption         =   "Before original WndProc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   150
         Width           =   2040
      End
      Begin VB.PictureBox picOptBefore 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   240
         ScaleHeight     =   465
         ScaleWidth      =   2190
         TabIndex        =   2
         Top             =   480
         Width           =   2190
         Begin VB.OptionButton optBefore 
            Caption         =   "Selected messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton optBefore 
            Caption         =   "All messages"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvBefore 
         Height          =   2640
         Left            =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
      Begin MSComctlLib.ListView lvAfter 
         Height          =   2640
         Left            =   150
         TabIndex        =   10
         Top             =   4875
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4260
         EndProperty
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuItm 
         Caption         =   "Do nothing"
         Index           =   0
      End
      Begin VB.Menu mnuItm 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuItm 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'* fMain - cSubclass based sample. Demonstrates adding and removing individual messages,
'*         ALL_MESSAGES and illustrates the range of windows messages available.
'*
'* Note: this project uses iSublass to define the callback interface. If you wish, you can instead
'*  use the WinSubHook3 type library as provided in the PSC submission. Simply remove iSubclass
'*  from this project and and a reference to the type library. No other changes are required.
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'*************************************************************************************************

Option Explicit

Private Type RECT
  Left              As Long
  Top               As Long
  Right             As Long
  Bottom            As Long
End Type

Private nTxtHeight  As Long                                                 'Height of a text line
Private rc          As RECT                                                 'Scrolling rectangle
Private oSub        As cSubclass                                            'cSubclass instance

Implements iSubclass

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function ScrollWindowEx Lib "user32" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  Dim i As Long
  Dim s As String

  Set oSub = New cSubclass
  oSub.Subclass Me.hwnd, Me
  
  With Me
    nTxtHeight = .TextHeight("My")
    rc.Top = .picHeader.Height
    .Move .Left, .Top, 11745, 7965
  End With

  Me.Height = (Me.Height - (Me.ScaleHeight * Screen.TwipsPerPixelY)) + (513 * Screen.TwipsPerPixelX)
  
  'Populate the listview controls
  For i = 0 To &H400
    s = GetMsgName(i)
    If Asc(Left$(s, 1)) <> 48 Then
      lvBefore.ListItems.Add , "k" & i, s
      lvAfter.ListItems.Add , "k" & i, s
    End If
  Next i

  'Sort the listview controls
  lvBefore.Sorted = True
  lvAfter.Sorted = True

  lvBefore.ColumnHeaders(1).Width = 2430#
  lvAfter.ColumnHeaders(1).Width = 2430#
End Sub

Private Sub Form_Resize()
  With Me
    If .WindowState <> vbMinimized Then
      .picHeader.Width = .ScaleWidth - .picMsgSel.Width + 1#
      .lblHeader.Width = .picHeader.ScaleWidth + 30#
      rc.Right = .picMsgSel.Left - 1
      rc.Bottom = .ScaleHeight
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set oSub = Nothing
End Sub

'cSubclass implemented interface callback
Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, lParamUser As Long)
  Dim sWhen As String

  If uMsg = eMsg.WM_PAINT Then
    'If we try to display the paint message we'll just cause another paint message... vicious circle.
    Exit Sub
  End If

  If bBefore Then
    sWhen = "Before "
  Else
    sWhen = "After  "
  End If

  lParamUser = lParamUser + 1
  Display sWhen, lReturn, lng_hWnd, uMsg, wParam, lParam, lParamUser
End Sub

'Enable/disable *after* subclassing
Private Sub chkAfter_Click()
  If chkAfter = 0 Then
    optAfter(0).Enabled = False
    optAfter(0).Value = False
    optAfter(1).Enabled = False
    optAfter(1).Value = False
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Deselect lvAfter
    'Delete all after messages
    oSub.DelMsg Me.hwnd, ALL_MESSAGES, MSG_AFTER
  Else
    optAfter(0).Enabled = True
    optAfter(0).Value = False
    optAfter(1).Enabled = True
    optAfter(1).Value = False
  End If
End Sub

'Enable/disable *before* subclassing
Private Sub chkBefore_Click()
  If chkBefore.Value = 0 Then
    optBefore(0).Enabled = False
    optBefore(0).Value = False
    optBefore(1).Enabled = False
    optBefore(1).Value = False
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Deselect lvBefore
    'Delete all before messages
    oSub.DelMsg Me.hwnd, ALL_MESSAGES, MSG_BEFORE
  Else
    optBefore(0).Enabled = True
    optBefore(0).Value = False
    optBefore(1).Enabled = True
    optBefore(1).Value = False
  End If
End Sub

'Uncheck all of the messages in the passed listview
Private Sub Deselect(ByVal lv As ListView)
  Dim itm As MSComctlLib.ListItem

  For Each itm In lv.ListItems
    itm.Checked = False
  Next itm
End Sub

'After list check box set/unset
Private Sub lvAfter_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As eMsg

  nMsg = Val(Mid$(Item.Key, 2))

  If Item.Checked Then
    'Add the after message
    oSub.AddMsg Me.hwnd, nMsg, MSG_AFTER
  Else
    'Delete the after message
    oSub.DelMsg Me.hwnd, Val(Mid$(Item.Key, 2)), MSG_AFTER
  End If

  If nMsg = eMsg.WM_MOUSEWHEEL Then
    'Mousewheel events will be captured by the listview, so set the focus elsewhere
    chkAfter.SetFocus
  End If
End Sub

'Before list check box set/unset
Private Sub lvBefore_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Dim nMsg As eMsg

  nMsg = Val(Mid$(Item.Key, 2))

  If Item.Checked Then
    'Add the before message
    oSub.AddMsg Me.hwnd, nMsg, MSG_BEFORE
  Else
    'Delete the before message
    oSub.DelMsg Me.hwnd, nMsg, MSG_BEFORE
  End If

  If nMsg = eMsg.WM_MOUSEWHEEL Then
    'Mousewheel events will be captured/stolen  by the listview, so set the focus elsewhere
    chkBefore.SetFocus
  End If
End Sub

Private Sub mnuItm_Click(Index As Integer)
  If Index = 2 Then
    Unload Me
  End If
End Sub

'After all or selected
Private Sub optAfter_Click(Index As Integer)
  If Index = 0 Then
    lvAfter.Enabled = False
    lvAfter.TextBackground = lvwTransparent
    Deselect lvAfter
    'Add all after messages
    oSub.AddMsg Me.hwnd, ALL_MESSAGES, MSG_AFTER
  Else
    lvAfter.TextBackground = lvwOpaque
    lvAfter.Enabled = True
    'Delete all after messages
    oSub.DelMsg Me.hwnd, ALL_MESSAGES, MSG_AFTER
  End If
End Sub

'Before all or selected
Private Sub optBefore_Click(Index As Integer)
  If Index = 0 Then
    lvBefore.Enabled = False
    lvBefore.TextBackground = lvwTransparent
    Deselect lvBefore
    'Add all before messages
    oSub.AddMsg Me.hwnd, ALL_MESSAGES, MSG_BEFORE
  Else
    lvBefore.TextBackground = lvwOpaque
    lvBefore.Enabled = True
    'Delete all before messages
    oSub.DelMsg Me.hwnd, ALL_MESSAGES, MSG_BEFORE
  End If
End Sub

'Display a line
Private Sub Display(ByVal sWhen As String, ByVal lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long, ByVal lParamUser As Long)
  Const SW_INVALIDATE As Long = &H2
  
  With Me
    ScrollWindowEx .hwnd, 0, -nTxtHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE
    UpdateWindow .hwnd
    .CurrentY = .ScaleHeight - nTxtHeight
    Print FmtHex(lParamUser) & sWhen & FmtHex(lReturn) & FmtHex(lng_hWnd) & FmtHex(uMsg) & FmtHex(wParam) & FmtHex(lParam) & GetMsgName(uMsg)
  End With
End Sub

'Return the passed Long value as a hex string with leading zeros, if required, to a width of eight characters, plus a trailing space
Private Function FmtHex(ByVal nValue As Long) As String
  FmtHex = Right$("0000000" & Hex$(nValue), 8) & " "
End Function

'Return the name of the passed messag number
Private Function GetMsgName(ByVal uMsg As eMsg) As String
  If uMsg = eMsg.WM_ACTIVATE Then
    GetMsgName = "WM_ACTIVATE"
  ElseIf uMsg = eMsg.WM_ACTIVATEAPP Then:           GetMsgName = "WM_ACTIVATEAPP"
  ElseIf uMsg = eMsg.WM_ASKCBFORMATNAME Then:       GetMsgName = "WM_ASKCBFORMATNAME"
  ElseIf uMsg = eMsg.WM_CANCELJOURNAL Then:         GetMsgName = "WM_CANCELJOURNAL"
  ElseIf uMsg = eMsg.WM_CANCELMODE Then:            GetMsgName = "WM_CANCELMODE"
  ElseIf uMsg = eMsg.WM_CAPTURECHANGED Then:        GetMsgName = "WM_CAPTURECHANGED"
  ElseIf uMsg = eMsg.WM_CHANGECBCHAIN Then:         GetMsgName = "WM_CHANGECBCHAIN"
  ElseIf uMsg = eMsg.WM_CHAR Then:                  GetMsgName = "WM_CHAR"
  ElseIf uMsg = eMsg.WM_CHARTOITEM Then:            GetMsgName = "WM_CHARTOITEM"
  ElseIf uMsg = eMsg.WM_CHILDACTIVATE Then:         GetMsgName = "WM_CHILDACTIVATE"
  ElseIf uMsg = eMsg.WM_CLEAR Then:                 GetMsgName = "WM_CLEAR"
  ElseIf uMsg = eMsg.WM_CLOSE Then:                 GetMsgName = "WM_CLOSE"
  ElseIf uMsg = eMsg.WM_COMMAND Then:               GetMsgName = "WM_COMMAND"
  ElseIf uMsg = eMsg.WM_COMPACTING Then:            GetMsgName = "WM_COMPACTING"
  ElseIf uMsg = eMsg.WM_COMPAREITEM Then:           GetMsgName = "WM_COMPAREITEM"
  ElseIf uMsg = eMsg.WM_CONTEXTMENU Then:           GetMsgName = "WM_CONTEXTMENU"
  ElseIf uMsg = eMsg.WM_COPY Then:                  GetMsgName = "WM_COPY"
  ElseIf uMsg = eMsg.WM_COPYDATA Then:              GetMsgName = "WM_COPYDATA"
  ElseIf uMsg = eMsg.WM_CREATE Then:                GetMsgName = "WM_CREATE"
  ElseIf uMsg = eMsg.WM_CTLCOLORBTN Then:           GetMsgName = "WM_CTLCOLORBTN"
  ElseIf uMsg = eMsg.WM_CTLCOLORDLG Then:           GetMsgName = "WM_CTLCOLORDLG"
  ElseIf uMsg = eMsg.WM_CTLCOLOREDIT Then:          GetMsgName = "WM_CTLCOLOREDIT"
  ElseIf uMsg = eMsg.WM_CTLCOLORLISTBOX Then:       GetMsgName = "WM_CTLCOLORLISTBOX"
  ElseIf uMsg = eMsg.WM_CTLCOLORMSGBOX Then:        GetMsgName = "WM_CTLCOLORMSGBOX"
  ElseIf uMsg = eMsg.WM_CTLCOLORSCROLLBAR Then:     GetMsgName = "WM_CTLCOLORSCROLLBAR"
  ElseIf uMsg = eMsg.WM_CTLCOLORSTATIC Then:        GetMsgName = "WM_CTLCOLORSTATIC"
  ElseIf uMsg = eMsg.WM_CUT Then:                   GetMsgName = "WM_CUT"
  ElseIf uMsg = eMsg.WM_DEADCHAR Then:              GetMsgName = "WM_DEADCHAR"
  ElseIf uMsg = eMsg.WM_DELETEITEM Then:            GetMsgName = "WM_DELETEITEM"
  ElseIf uMsg = eMsg.WM_DESTROY Then:               GetMsgName = "WM_DESTROY"
  ElseIf uMsg = eMsg.WM_DESTROYCLIPBOARD Then:      GetMsgName = "WM_DESTROYCLIPBOARD"
  ElseIf uMsg = eMsg.WM_DRAWCLIPBOARD Then:         GetMsgName = "WM_DRAWCLIPBOARD"
  ElseIf uMsg = eMsg.WM_DRAWITEM Then:              GetMsgName = "WM_DRAWITEM"
  ElseIf uMsg = eMsg.WM_DROPFILES Then:             GetMsgName = "WM_DROPFILES"
  ElseIf uMsg = eMsg.WM_ENABLE Then:                GetMsgName = "WM_ENABLE"
  ElseIf uMsg = eMsg.WM_ENDSESSION Then:            GetMsgName = "WM_ENDSESSION"
  ElseIf uMsg = eMsg.WM_ENTERIDLE Then:             GetMsgName = "WM_ENTERIDLE"
  ElseIf uMsg = eMsg.WM_ENTERMENULOOP Then:         GetMsgName = "WM_ENTERMENULOOP"
  ElseIf uMsg = eMsg.WM_ENTERSIZEMOVE Then:         GetMsgName = "WM_ENTERSIZEMOVE"
  ElseIf uMsg = eMsg.WM_ERASEBKGND Then:            GetMsgName = "WM_ERASEBKGND"
  ElseIf uMsg = eMsg.WM_EXITMENULOOP Then:          GetMsgName = "WM_EXITMENULOOP"
  ElseIf uMsg = eMsg.WM_EXITSIZEMOVE Then:          GetMsgName = "WM_EXITSIZEMOVE"
  ElseIf uMsg = eMsg.WM_FONTCHANGE Then:            GetMsgName = "WM_FONTCHANGE"
  ElseIf uMsg = eMsg.WM_GETDLGCODE Then:            GetMsgName = "WM_GETDLGCODE"
  ElseIf uMsg = eMsg.WM_GETFONT Then:               GetMsgName = "WM_GETFONT"
  ElseIf uMsg = eMsg.WM_GETHOTKEY Then:             GetMsgName = "WM_GETHOTKEY"
  ElseIf uMsg = eMsg.WM_GETMINMAXINFO Then:         GetMsgName = "WM_GETMINMAXINFO"
  ElseIf uMsg = eMsg.WM_GETTEXT Then:               GetMsgName = "WM_GETTEXT"
  ElseIf uMsg = eMsg.WM_GETTEXTLENGTH Then:         GetMsgName = "WM_GETTEXTLENGTH"
  ElseIf uMsg = eMsg.WM_HOTKEY Then:                GetMsgName = "WM_HOTKEY"
  ElseIf uMsg = eMsg.WM_HSCROLL Then:               GetMsgName = "WM_HSCROLL"
  ElseIf uMsg = eMsg.WM_HSCROLLCLIPBOARD Then:      GetMsgName = "WM_HSCROLLCLIPBOARD"
  ElseIf uMsg = eMsg.WM_ICONERASEBKGND Then:        GetMsgName = "WM_ICONERASEBKGND"
  ElseIf uMsg = eMsg.WM_IME_CHAR Then:              GetMsgName = "WM_IME_CHAR"
  ElseIf uMsg = eMsg.WM_IME_COMPOSITION Then:       GetMsgName = "WM_IME_COMPOSITION"
  ElseIf uMsg = eMsg.WM_IME_COMPOSITIONFULL Then:   GetMsgName = "WM_IME_COMPOSITIONFULL"
  ElseIf uMsg = eMsg.WM_IME_CONTROL Then:           GetMsgName = "WM_IME_CONTROL"
  ElseIf uMsg = eMsg.WM_IME_ENDCOMPOSITION Then:    GetMsgName = "WM_IME_ENDCOMPOSITION"
  ElseIf uMsg = eMsg.WM_IME_KEYDOWN Then:           GetMsgName = "WM_IME_KEYDOWN"
  ElseIf uMsg = eMsg.WM_IME_KEYLAST Then:           GetMsgName = "WM_IME_KEYLAST"
  ElseIf uMsg = eMsg.WM_IME_KEYUP Then:             GetMsgName = "WM_IME_KEYUP"
  ElseIf uMsg = eMsg.WM_IME_NOTIFY Then:            GetMsgName = "WM_IME_NOTIFY"
  ElseIf uMsg = eMsg.WM_IME_SELECT Then:            GetMsgName = "WM_IME_SELECT"
  ElseIf uMsg = eMsg.WM_IME_SETCONTEXT Then:        GetMsgName = "WM_IME_SETCONTEXT"
  ElseIf uMsg = eMsg.WM_IME_STARTCOMPOSITION Then:  GetMsgName = "WM_IME_STARTCOMPOSITION"
  ElseIf uMsg = eMsg.WM_INITDIALOG Then:            GetMsgName = "WM_INITDIALOG"
  ElseIf uMsg = eMsg.WM_INITMENU Then:              GetMsgName = "WM_INITMENU"
  ElseIf uMsg = eMsg.WM_INITMENUPOPUP Then:         GetMsgName = "WM_INITMENUPOPUP"
  ElseIf uMsg = eMsg.WM_KEYDOWN Then:               GetMsgName = "WM_KEYDOWN"
  ElseIf uMsg = eMsg.WM_KEYFIRST Then:              GetMsgName = "WM_KEYFIRST"
  ElseIf uMsg = eMsg.WM_KEYLAST Then:               GetMsgName = "WM_KEYLAST"
  ElseIf uMsg = eMsg.WM_KEYUP Then:                 GetMsgName = "WM_KEYUP"
  ElseIf uMsg = eMsg.WM_KILLFOCUS Then:             GetMsgName = "WM_KILLFOCUS"
  ElseIf uMsg = eMsg.WM_LBUTTONDBLCLK Then:         GetMsgName = "WM_LBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_LBUTTONDOWN Then:           GetMsgName = "WM_LBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_LBUTTONUP Then:             GetMsgName = "WM_LBUTTONUP"
  ElseIf uMsg = eMsg.WM_MBUTTONDBLCLK Then:         GetMsgName = "WM_MBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_MBUTTONDOWN Then:           GetMsgName = "WM_MBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_MBUTTONUP Then:             GetMsgName = "WM_MBUTTONUP"
  ElseIf uMsg = eMsg.WM_MDIACTIVATE Then:           GetMsgName = "WM_MDIACTIVATE"
  ElseIf uMsg = eMsg.WM_MDICASCADE Then:            GetMsgName = "WM_MDICASCADE"
  ElseIf uMsg = eMsg.WM_MDICREATE Then:             GetMsgName = "WM_MDICREATE"
  ElseIf uMsg = eMsg.WM_MDIDESTROY Then:            GetMsgName = "WM_MDIDESTROY"
  ElseIf uMsg = eMsg.WM_MDIGETACTIVE Then:          GetMsgName = "WM_MDIGETACTIVE"
  ElseIf uMsg = eMsg.WM_MDIICONARRANGE Then:        GetMsgName = "WM_MDIICONARRANGE"
  ElseIf uMsg = eMsg.WM_MDIMAXIMIZE Then:           GetMsgName = "WM_MDIMAXIMIZE"
  ElseIf uMsg = eMsg.WM_MDINEXT Then:               GetMsgName = "WM_MDINEXT"
  ElseIf uMsg = eMsg.WM_MDIREFRESHMENU Then:        GetMsgName = "WM_MDIREFRESHMENU"
  ElseIf uMsg = eMsg.WM_MDIRESTORE Then:            GetMsgName = "WM_MDIRESTORE"
  ElseIf uMsg = eMsg.WM_MDISETMENU Then:            GetMsgName = "WM_MDISETMENU"
  ElseIf uMsg = eMsg.WM_MDITILE Then:               GetMsgName = "WM_MDITILE"
  ElseIf uMsg = eMsg.WM_MEASUREITEM Then:           GetMsgName = "WM_MEASUREITEM"
  ElseIf uMsg = eMsg.WM_MENUCHAR Then:              GetMsgName = "WM_MENUCHAR"
  ElseIf uMsg = eMsg.WM_MENUSELECT Then:            GetMsgName = "WM_MENUSELECT"
  ElseIf uMsg = eMsg.WM_MOUSEACTIVATE Then:         GetMsgName = "WM_MOUSEACTIVATE"
  ElseIf uMsg = eMsg.WM_MOUSEMOVE Then:             GetMsgName = "WM_MOUSEMOVE"
  ElseIf uMsg = eMsg.WM_MOUSEWHEEL Then:            GetMsgName = "WM_MOUSEWHEEL"
  ElseIf uMsg = eMsg.WM_MOVE Then:                  GetMsgName = "WM_MOVE"
  ElseIf uMsg = eMsg.WM_MOVING Then:                GetMsgName = "WM_MOVING"
  ElseIf uMsg = eMsg.WM_NCACTIVATE Then:            GetMsgName = "WM_NCACTIVATE"
  ElseIf uMsg = eMsg.WM_NCCALCSIZE Then:            GetMsgName = "WM_NCCALCSIZE"
  ElseIf uMsg = eMsg.WM_NCCREATE Then:              GetMsgName = "WM_NCCREATE"
  ElseIf uMsg = eMsg.WM_NCDESTROY Then:             GetMsgName = "WM_NCDESTROY"
  ElseIf uMsg = eMsg.WM_NCHITTEST Then:             GetMsgName = "WM_NCHITTEST"
  ElseIf uMsg = eMsg.WM_NCLBUTTONDBLCLK Then:       GetMsgName = "WM_NCLBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_NCLBUTTONDOWN Then:         GetMsgName = "WM_NCLBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_NCLBUTTONUP Then:           GetMsgName = "WM_NCLBUTTONUP"
  ElseIf uMsg = eMsg.WM_NCMBUTTONDBLCLK Then:       GetMsgName = "WM_NCMBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_NCMBUTTONDOWN Then:         GetMsgName = "WM_NCMBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_NCMBUTTONUP Then:           GetMsgName = "WM_NCMBUTTONUP"
  ElseIf uMsg = eMsg.WM_NCMOUSEMOVE Then:           GetMsgName = "WM_NCMOUSEMOVE"
  ElseIf uMsg = eMsg.WM_NCPAINT Then:               GetMsgName = "WM_NCPAINT"
  ElseIf uMsg = eMsg.WM_NCRBUTTONDBLCLK Then:       GetMsgName = "WM_NCRBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_NCRBUTTONDOWN Then:         GetMsgName = "WM_NCRBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_NCRBUTTONUP Then:           GetMsgName = "WM_NCRBUTTONUP"
  ElseIf uMsg = eMsg.WM_NEXTDLGCTL Then:            GetMsgName = "WM_NEXTDLGCTL"
  ElseIf uMsg = eMsg.WM_NEXTMENU Then:              GetMsgName = "WM_NEXTMENU"
  ElseIf uMsg = eMsg.WM_NULL Then:                  GetMsgName = "WM_NULL"
  ElseIf uMsg = eMsg.WM_PAINT Then:                 GetMsgName = "WM_PAINT"
  ElseIf uMsg = eMsg.WM_PAINTCLIPBOARD Then:        GetMsgName = "WM_PAINTCLIPBOARD"
  ElseIf uMsg = eMsg.WM_PAINTICON Then:             GetMsgName = "WM_PAINTICON"
  ElseIf uMsg = eMsg.WM_PALETTECHANGED Then:        GetMsgName = "WM_PALETTECHANGED"
  ElseIf uMsg = eMsg.WM_PALETTEISCHANGING Then:     GetMsgName = "WM_PALETTEISCHANGING"
  ElseIf uMsg = eMsg.WM_PARENTNOTIFY Then:          GetMsgName = "WM_PARENTNOTIFY"
  ElseIf uMsg = eMsg.WM_PASTE Then:                 GetMsgName = "WM_PASTE"
  ElseIf uMsg = eMsg.WM_PENWINFIRST Then:           GetMsgName = "WM_PENWINFIRST"
  ElseIf uMsg = eMsg.WM_PENWINLAST Then:            GetMsgName = "WM_PENWINLAST"
  ElseIf uMsg = eMsg.WM_POWER Then:                 GetMsgName = "WM_POWER"
  ElseIf uMsg = eMsg.WM_QUERYDRAGICON Then:         GetMsgName = "WM_QUERYDRAGICON"
  ElseIf uMsg = eMsg.WM_QUERYENDSESSION Then:       GetMsgName = "WM_QUERYENDSESSION"
  ElseIf uMsg = eMsg.WM_QUERYNEWPALETTE Then:       GetMsgName = "WM_QUERYNEWPALETTE"
  ElseIf uMsg = eMsg.WM_QUERYOPEN Then:             GetMsgName = "WM_QUERYOPEN"
  ElseIf uMsg = eMsg.WM_QUEUESYNC Then:             GetMsgName = "WM_QUEUESYNC"
  ElseIf uMsg = eMsg.WM_QUIT Then:                  GetMsgName = "WM_QUIT"
  ElseIf uMsg = eMsg.WM_RBUTTONDBLCLK Then:         GetMsgName = "WM_RBUTTONDBLCLK"
  ElseIf uMsg = eMsg.WM_RBUTTONDOWN Then:           GetMsgName = "WM_RBUTTONDOWN"
  ElseIf uMsg = eMsg.WM_RBUTTONUP Then:             GetMsgName = "WM_RBUTTONUP"
  ElseIf uMsg = eMsg.WM_RENDERALLFORMATS Then:      GetMsgName = "WM_RENDERALLFORMATS"
  ElseIf uMsg = eMsg.WM_RENDERFORMAT Then:          GetMsgName = "WM_RENDERFORMAT"
  ElseIf uMsg = eMsg.WM_SETCURSOR Then:             GetMsgName = "WM_SETCURSOR"
  ElseIf uMsg = eMsg.WM_SETFOCUS Then:              GetMsgName = "WM_SETFOCUS"
  ElseIf uMsg = eMsg.WM_SETFONT Then:               GetMsgName = "WM_SETFONT"
  ElseIf uMsg = eMsg.WM_SETHOTKEY Then:             GetMsgName = "WM_SETHOTKEY"
  ElseIf uMsg = eMsg.WM_SETREDRAW Then:             GetMsgName = "WM_SETREDRAW"
  ElseIf uMsg = eMsg.WM_SETTEXT Then:               GetMsgName = "WM_SETTEXT"
  ElseIf uMsg = eMsg.WM_SHOWWINDOW Then:            GetMsgName = "WM_SHOWWINDOW"
  ElseIf uMsg = eMsg.WM_SIZE Then:                  GetMsgName = "WM_SIZE"
  ElseIf uMsg = eMsg.WM_SIZECLIPBOARD Then:         GetMsgName = "WM_SIZECLIPBOARD"
  ElseIf uMsg = eMsg.WM_SIZING Then:                GetMsgName = "WM_SIZING"
  ElseIf uMsg = eMsg.WM_SPOOLERSTATUS Then:         GetMsgName = "WM_SPOOLERSTATUS"
  ElseIf uMsg = eMsg.WM_SYNCPAINT Then:             GetMsgName = "WM_SYNCPAINT"
  ElseIf uMsg = eMsg.WM_SYSCHAR Then:               GetMsgName = "WM_SYSCHAR"
  ElseIf uMsg = eMsg.WM_SYSCOLORCHANGE Then:        GetMsgName = "WM_SYSCOLORCHANGE"
  ElseIf uMsg = eMsg.WM_SYSCOMMAND Then:            GetMsgName = "WM_SYSCOMMAND"
  ElseIf uMsg = eMsg.WM_SYSDEADCHAR Then:           GetMsgName = "WM_SYSDEADCHAR"
  ElseIf uMsg = eMsg.WM_SYSKEYDOWN Then:            GetMsgName = "WM_SYSKEYDOWN"
  ElseIf uMsg = eMsg.WM_SYSKEYUP Then:              GetMsgName = "WM_SYSKEYUP"
  ElseIf uMsg = eMsg.WM_TIMECHANGE Then:            GetMsgName = "WM_TIMECHANGE"
  ElseIf uMsg = eMsg.WM_TIMER Then:                 GetMsgName = "WM_TIMER"
  ElseIf uMsg = eMsg.WM_UNDO Then:                  GetMsgName = "WM_UNDO"
  ElseIf uMsg = eMsg.WM_UNINITMENUPOPUP Then:       GetMsgName = "WM_UNINITEDMENUPOPUP"
  ElseIf uMsg = eMsg.WM_USER Then:                  GetMsgName = "WM_USER"
  ElseIf uMsg = eMsg.WM_VKEYTOITEM Then:            GetMsgName = "WM_VKEYTOITEM"
  ElseIf uMsg = eMsg.WM_VSCROLL Then:               GetMsgName = "WM_VSCROLL"
  ElseIf uMsg = eMsg.WM_VSCROLLCLIPBOARD Then:      GetMsgName = "WM_VSCROLL"
  ElseIf uMsg = eMsg.WM_WINDOWPOSCHANGED Then:      GetMsgName = "WM_WINDOWPOSCHANGED"
  ElseIf uMsg = eMsg.WM_WINDOWPOSCHANGING Then:     GetMsgName = "WM_WINDOWPOSCHANGING"
  ElseIf uMsg = eMsg.WM_WININICHANGE Then:          GetMsgName = "WM_WININICHANGE"
  Else:                                             GetMsgName = FmtHex(uMsg)
  End If
End Function
