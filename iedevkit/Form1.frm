VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Global IE Security Settnigs "
      Height          =   1335
      Index           =   1
      Left            =   4860
      TabIndex        =   33
      Top             =   6120
      Width           =   4335
      Begin VB.CheckBox chkSecuritySetting 
         Caption         =   "Allow Persistance"
         Height          =   315
         Index           =   4
         Left            =   2640
         TabIndex        =   38
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkSecuritySetting 
         Caption         =   "Allow Scripting"
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkSecuritySetting 
         Caption         =   "Allow Form Submission"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   2115
      End
      Begin VB.CheckBox chkSecuritySetting 
         Caption         =   "Allow Cookies"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   540
         Width           =   1815
      End
      Begin VB.CheckBox chkSecuritySetting 
         Caption         =   "OverRide Safe for Scripting"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Navigation && Right Click"
      Height          =   1395
      Left            =   0
      TabIndex        =   26
      Top             =   3600
      Width           =   2475
      Begin VB.CheckBox chkNoRightClick 
         Caption         =   "Disable Right Click menu"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   900
         Width           =   2115
      End
      Begin VB.CheckBox chkCustomRightClickMenu 
         Caption         =   "Show Custom  Rt Clk Menu"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   540
         Width           =   2295
      End
      Begin VB.CheckBox chkEditBeforeNavigate 
         Caption         =   "Edit URL before  Navigate"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2235
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Typing "
      Height          =   1335
      Left            =   6300
      TabIndex        =   20
      Top             =   3660
      Width           =   2895
      Begin VB.TextBox txtWbChar 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkOverrideWbTyping 
         Caption         =   "Override Display Char"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   1875
      End
      Begin VB.TextBox txtOverrideChar 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   22
         Text            =   "*"
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkBlockTyping 
         Caption         =   "Block Typing into Wb"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Last Character Typed"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Interactive Demos "
      Height          =   1335
      Left            =   0
      TabIndex        =   19
      Top             =   6120
      Width           =   4815
      Begin VB.CommandButton cmdDhtmlEvents 
         Caption         =   "Hook DHTML Events"
         Height          =   435
         Left            =   240
         TabIndex        =   32
         Top             =   780
         Width           =   2295
      End
      Begin VB.CommandButton cmdIEOptions 
         Caption         =   "Show IE Options"
         Height          =   435
         Left            =   2700
         TabIndex        =   31
         Top             =   300
         Width           =   1995
      End
      Begin VB.CommandButton cmdWindowExternal 
         Caption         =   "Demo Window.GetExternal"
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   300
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Window Attributes "
      Height          =   1035
      Left            =   0
      TabIndex        =   11
      Top             =   5040
      Width           =   9195
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "In PLace Navigation"
         Height          =   375
         Index           =   6
         Left            =   7080
         TabIndex        =   18
         Top             =   300
         Width           =   1755
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "Use Flat Scroll Bars"
         Height          =   375
         Index           =   5
         Left            =   4440
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "Links In New Window"
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   16
         Top             =   240
         Width           =   1875
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "No Scrollbars"
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "No 3d Border"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "Disable Help Menus"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkWindowAttribute 
         Caption         =   "Disable Selections"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Block Special Commands "
      Height          =   1335
      Index           =   0
      Left            =   2580
      TabIndex        =   3
      Top             =   3660
      Width           =   3555
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "SelectAll"
         Height          =   315
         Index           =   6
         Left            =   1440
         TabIndex        =   10
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "Refresh"
         Height          =   315
         Index           =   5
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "Print"
         Height          =   315
         Index           =   4
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "New Window"
         Height          =   315
         Index           =   3
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "GoForward"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "GoBack"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
      Begin VB.CheckBox chkBlockCommand 
         Caption         =   "Find"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtUrl 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   3240
      Width           =   7575
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "Navigate"
      Height          =   315
      Left            =   7740
      TabIndex        =   1
      Top             =   3240
      Width           =   1395
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9075
      ExtentX         =   16007
      ExtentY         =   5424
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSample 
         Caption         =   "Sample Menu 1"
         Index           =   0
      End
      Begin VB.Menu mnuSample 
         Caption         =   "Sample Menu 2"
         Index           =   1
      End
      Begin VB.Menu mnuSample 
         Caption         =   "Sample Menu 3"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'clsWbExtender exposed 5 events and host of
'properties not available by default to us in VB
'this sample has alot, but it still isnt everything
'see the html help file for more details.

Dim WithEvents clsExtender As clsWbExtender
Attribute clsExtender.VB_VarHelpID = -1

Dim clsSec As New clsSecurityManager


Private Sub Form_Load()
    wb.Navigate2 "about:blank"
    
    Set clsExtender = New clsWbExtender
    clsExtender.WbAttributes = InPlaceNavigation
    
    clsExtender.HookWebBrowser wb
    
    txtUrl = "http://google.com"
    
    Call loadCurrentInternetSecurityPolicy
    
End Sub

Private Sub cmdIEOptions_Click()
    clsSec.ShowInternetOptions
End Sub

Private Sub cmdNavigate_Click()
    On Error Resume Next
    wb.Navigate2 CStr(txtUrl)
End Sub

Private Sub cmdWindowExternal_Click()
    wb.Navigate2 App.Path & "\external_demo.html"
End Sub


'----------------------------------------------------------------
'  These are the new events exposed for the
'                               Webbrowser control by clsExtender
'----------------------------------------------------------------

Private Sub clsExtender_GetExternal(oIDispatch As Object)
    'this allows javascript to access the objects we return
    'here is it set so javascript will have access to all functions
    'and objects on this form.
    Set oIDispatch = Me
End Sub

Private Sub clsExtender_EditUrlBeforeNavigate(url As String)
    If chkEditBeforeNavigate.Value Then
        url = InputBox("Here you can Edit Url the browser is requesting.", , url)
    End If
End Sub

Private Sub clsExtender_KeyPress(keyCode As Integer, cancel As Boolean)
    txtWbChar = Chr(keyCode)
    If chkOverrideWbTyping.Value Then keyCode = Asc(txtOverrideChar)
    If chkBlockTyping.Value Then cancel = True
End Sub

Private Sub clsExtender_OnContextMenu(cancel As Boolean)
    
    If chkNoRightClick.Value Then
        cancel = True
        Exit Sub
    End If
    
    If chkCustomRightClickMenu.Value Then
        PopupMenu mnuPopup
        cancel = True
    End If
    
End Sub
 
Private Sub clsExtender_WbCommand(cmdId As WbCommands, BlockCommand As Boolean)
    Dim i As Integer
    
    'these are just some of the commands you can receive and block
    'not all commands are listed in the WbCommands enum, for those
    'that arent, you can still process them by receiving them by numeral
    
    Select Case cmdId
        Case wbFind: i = 0
        Case wbGoBack: i = 1
        Case wbGoForw: i = 2
        Case wbNew: i = 3
        Case wbPrint: i = 4
        Case wbRefresh: i = 5
        Case wbSelectAll: i = 6
    End Select
        
    If chkBlockCommand(i).Value Then BlockCommand = True
    
End Sub

'-------------------------------------------------------------------------



Private Sub chkWindowAttribute_Click(Index As Integer)
    
    Dim i As Long
    
    'here we save the enumerations value to i
    Select Case Index
        Case 0: i = DisableSelections
        Case 1: i = DisableHelpMenu
        Case 2: i = No3DBorder
        Case 3: i = NoScrollBars
        Case 4: i = LinksInNewWindow
        Case 5: i = UseFlatScrollBars
        Case 6: i = InPlaceNavigation
    End Select
    
    'if this attribute is active we add it to the current
    'atttributes multiple options can be set at once by OR'ing
    'them together such as:
    'clsextender.WbAttributes =DisableSelections or NoScrollBars
    
    If chkWindowAttribute(Index).Value Then
        clsExtender.WbAttributes = clsExtender.WbAttributes Or i
    Else
         clsExtender.WbAttributes = clsExtender.WbAttributes Xor i
    End If
    
    wb.Refresh
    
End Sub

Private Sub chkSecuritySetting_Click(Index As Integer)
    Dim i As Long
    'save enumeration value into variable
    Select Case Index
        Case 0: i = OverrideSafeForScripting
        Case 1: i = AllowCookies
        Case 2: i = AllowSubmitForms
        Case 3: i = AllowScripting
        Case 4: i = AllowUserdataPersistance
    End Select
    
    If chkSecuritySetting(Index).Value = 1 Then
        clsSec.SetPolicy i, True
    Else
         clsSec.SetPolicy i, False
    End If
    
End Sub


'this function gets called in demo getexternal
'script from webpage javascript
Function myExternalFunction(StrIn As String) As String
    
    Const v = vbCrLf & vbCrLf
    
    MsgBox "Form Function just called from" & _
           " javascript with argument:" & v & StrIn, vbInformation
           
    myExternalFunction = StrIn & v & "Which has now been modified by Vb Function."
    
End Function
    





Private Sub cmdDhtmlEvents_Click()
    wb.Navigate2 App.Path & "\dhtml_events.html"
    
    'pause until page loaded
    While Not wb.ReadyState = READYSTATE_COMPLETE: DoEvents: Wend
    
    Dim clsHook As New clsDhtmlEvent
    
    'here, we specify:
    '    1) what object has the function to call (me) <--this form
    '    2) the name of the function to call
    '    3) optionally we include an argument to be passed to it.
    clsHook.SetReference Me, "myDhtmlEventHandler", "button 1 pressed"
    
    'note that you will not have intellisense for this line
    wb.Document.getElementById("btnHookme").onclick = clsHook
    
End Sub

Function myDhtmlEventHandler(StrIn)

    MsgBox "Function myDhtmlEventHandler called " & _
           "on main form with argument:" & _
            vbCrLf & vbCrLf & StrIn
            
End Function










Sub loadCurrentInternetSecurityPolicy()
    chkSecuritySetting(0).Value = IIf(clsSec.IsPolicyActive(OverrideSafeForScripting), 1, 0)
    chkSecuritySetting(1).Value = IIf(clsSec.IsPolicyActive(AllowCookies), 1, 0)
    chkSecuritySetting(2).Value = IIf(clsSec.IsPolicyActive(AllowSubmitForms), 1, 0)
    chkSecuritySetting(3).Value = IIf(clsSec.IsPolicyActive(AllowScripting), 1, 0)
    chkSecuritySetting(4).Value = IIf(clsSec.IsPolicyActive(AllowUserdataPersistance), 1, 0)
End Sub
