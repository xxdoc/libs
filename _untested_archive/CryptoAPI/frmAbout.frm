VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5325
   ClientLeft      =   1920
   ClientTop       =   2265
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675.409
   ScaleMode       =   0  'User
   ScaleWidth      =   4817.334
   Begin VB.Frame fraThanks 
      Height          =   3825
      Left            =   45
      TabIndex        =   4
      Top             =   675
      Width           =   5025
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   150
         Picture         =   "frmAbout.frx":030A
         ScaleHeight     =   330
         ScaleWidth      =   4755
         TabIndex        =   16
         Top             =   2835
         Width           =   4755
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VBForums - VB6 or earlier"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   4
            Left            =   1695
            TabIndex        =   17
            Top             =   15
            Width           =   1935
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   150
         Picture         =   "frmAbout.frx":18AC
         ScaleHeight     =   525
         ScaleWidth      =   4755
         TabIndex        =   11
         Top             =   495
         Width           =   4755
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VBNet API code snippets for VB6"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   1695
            TabIndex        =   12
            Top             =   180
            Width           =   2685
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   1
         Left            =   150
         Picture         =   "frmAbout.frx":2DB6
         ScaleHeight     =   465
         ScaleWidth      =   4755
         TabIndex        =   9
         Top             =   1113
         Width           =   4755
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Planet Source Code for Visual Basic"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   1695
            TabIndex        =   10
            Top             =   105
            Width           =   2970
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   150
         Picture         =   "frmAbout.frx":4574
         ScaleHeight     =   510
         ScaleWidth      =   4755
         TabIndex        =   7
         Top             =   1671
         Width           =   4755
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Classic VB code snippets"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   2
            Left            =   1695
            TabIndex        =   8
            Top             =   135
            Width           =   2115
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   3
         Left            =   150
         Picture         =   "frmAbout.frx":6A5E
         ScaleHeight     =   420
         ScaleWidth      =   4755
         TabIndex        =   5
         Top             =   2274
         Width           =   4755
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "All API Network - VB6 reference "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   1695
            TabIndex        =   6
            Top             =   105
            Width           =   2580
         End
      End
      Begin VB.Label lblThankYou 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Acknowledgements"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1425
         TabIndex        =   14
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label lblOperSystem 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lblOperSystem"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   13
         Top             =   3330
         Width           =   4755
      End
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Index           =   1
      Left            =   4335
      Picture         =   "frmAbout.frx":87A0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Return to main screen"
      Top             =   4560
      Width           =   690
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Index           =   0
      Left            =   3585
      Picture         =   "frmAbout.frx":8AAA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "System Information"
      Top             =   4560
      Width           =   690
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblFormTitle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   15
      Top             =   90
      Width           =   4875
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   75
      TabIndex        =   3
      Top             =   4545
      Width           =   2550
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1725
      TabIndex        =   0
      Top             =   465
      Width           =   1680
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmAbout
'
' Description:   This form displays the sites I would like to give thanks.
'                For their code or other information on how to accomplish
'                something.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Verified screen would not be displayed during initial load
' 05-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated hyperlinks and user interface links
' 28-Jan-2014  Kenneth Ives  kenaso@tx.rr.com
'              Update DisplayInfo() routine for easier access to MSINFO32.EXE
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME               As String = "frmAbout"
  Private Const MS_INFO                   As String = "MSInfo32.exe"
  Private Const HKEY_CLASSES_ROOT         As Long = &H80000000
  Private Const FLAG_ICC_FORCE_CONNECTION As Long = 1
  
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' GetDesktopWindow function retrieves a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted. The return
  ' value is a handle to the desktop window.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long

  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hWnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' InternetCheckConnection function allows an application to check
  ' if a connection to the Internet can be established.  I use the URL
  ' "http://www.google.com/" for testing since they are always available
  ' when I am online.  "http://" or "https://" must prefix URL else the
  ' site will not be found. If function succeeds, return value is nonzero.
  Private Declare Function InternetCheckConnection Lib "wininet.dll" _
          Alias "InternetCheckConnectionA" (ByVal sUrl As String, _
          ByVal lFlags As Long, ByVal lReserved As Long) As Long

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------------
' Naming standard:   m str BrowserPath
' Variable name:     mstrBrowserPath
' ***************************************************************************
  Private mstrBrowserPath As String

Private Sub Form_Load()
    
    ' Find default browser path
    mstrBrowserPath = FindDefaultPath(HKEY_CLASSES_ROOT, "http\shell\open\command")
    
    ' If no path returned then try HTTPS
    If Len(Trim$(mstrBrowserPath)) = 0 Then
        mstrBrowserPath = FindDefaultPath(HKEY_CLASSES_ROOT, "https\shell\open\command")
    End If

    ' If no path returned then use system installed Internet Explorer
    If Len(Trim$(mstrBrowserPath)) = 0 Then
        mstrBrowserPath = FindDefaultPath(HKEY_CLASSES_ROOT, "Applications\iexplore.exe\shell\open\command")
    End If
    
    DisableX frmAbout   ' Disable "X" in upper right corner of form
    
    ' Hide this form
    With frmAbout
        .Hide
        .Caption = "About - " & PGM_NAME
    
        ' If operating system is Windows 8 or 8.1
        ' form caption is centered automatically
        If Not gblnWin8or81 Then
            CenterCaption frmAbout   ' Manually center form caption
        End If
    
        .lblFormTitle.Caption = PGM_NAME
        .lblDisclaimer.Caption = "This software is provided without any " & _
                                 "warrantees or guarantees implied or intended."
        .lblOperSystem.Caption = gstrOperSystem
    
        ' center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Based on the unload code the system passes, we determine what to do.
    '
    ' Unloadmode codes
    '     0 - Close from the control-menu box or Upper right "X"
    '     1 - Unload method from code elsewhere in the application
    '     2 - Windows Session is ending
    '     3 - Task Manager is closing the application
    '     4 - MDI Parent is closing
    Select Case UnloadMode
           Case 0    ' return to main form
                frmMain.Show
                frmAbout.Hide
    
           Case Else
                ' Fall thru. Something else is shutting us down.
    End Select

End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
             
           Case 0   ' Display System Information (MSInfo32.exe)
                DisplaySysInfo
    
           Case 1   ' Return to main form
                frmAbout.Hide
                frmMain.Visible = True
    End Select
  
End Sub

Private Sub DisplaySysInfo()

    StopProcessByName MS_INFO   ' Verify MSInfo32.exe is not active
    Shell MS_INFO               ' Display system information
    
End Sub

Public Sub DisplayAbout()

    ' Called by frmMain to display this form
    '
    '             +-------------- Center this form
    '             |         +---- on top of this one
    CenterForm frmAbout, frmMain
    
    ' While this form is displayed deactivate bottom
    ' form. To keep both forms active use vbModeless
    ' instead of vbModal.
    frmAbout.Show vbModal
    
End Sub

Private Sub lblURL_Click(Index As Integer)

    Dim strURL As String
    
    ' Identify URL to be executed
    Select Case Index
           Case 0: strURL = "http://vbnet.mvps.org/"
           Case 1: strURL = "http://www.Planet-Source-Code.com/vb/default.asp"
           Case 2: strURL = "http://vb.mvps.org/"
           Case 3: strURL = "http://allapi.mentalis.org/apilist/apilist.php"
           Case 4: strURL = "http://www.vbforums.com/forumdisplay.php?1-Visual-Basic-6-and-Earlier"
           Case Else: Exit Sub
    End Select
           
    RunShell strURL   ' Make hyperlink call
   
End Sub

Private Sub picURL_Click(Index As Integer)

    Dim strURL As String
    
    ' Identify URL to be executed
    Select Case Index
           Case 0: strURL = "http://vbnet.mvps.org/"
           Case 1: strURL = "http://www.Planet-Source-Code.com/vb/default.asp"
           Case 2: strURL = "http://vb.mvps.org/"
           Case 3: strURL = "http://allapi.mentalis.org/apilist/apilist.php"
           Case 4: strURL = "http://www.vbforums.com/forumdisplay.php?1-Visual-Basic-6-and-Earlier"
           Case Else: Exit Sub
    End Select
           
    RunShell strURL   ' Make hyperlink call
   
End Sub

Private Sub RunShell(ByVal strURL As String)

    ' Called by lblURL_Click()
    '           picURL_Click()
        
    ' When using API ShellExecute() command, sometimes the
    ' browser will hang and show a blank page if it is not
    ' already open.
    
    ' Note:  "http://" must use prefix URL
    '        else site will not be found
    Const TEST_URL_1   As String = "http://www.google.com/"
    Const TEST_URL_2   As String = "http://www.yahoo.com/"
    Const ROUTINE_NAME As String = "RunShell"
    
    If Len(mstrBrowserPath) > 0 Then
        
        ' Attempt to see if first site is available
        If InternetCheckConnection(TEST_URL_1, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        
            ' If first site not available then check a second site
            If InternetCheckConnection(TEST_URL_2, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
        
                InfoMsg "An internet connection is not available at this time." & vbNewLine & _
                        "Please try again later." & vbNewLine & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                Exit Sub
    
            End If
        End If
        
        ' See if default browser is already active
        If FindHandleByExe(mstrBrowserPath) > 0 Then
            
            ' Browser is already open, add another tab
            ShellExecute GetDesktopWindow(), "open", strURL, _
                         vbNullString, vbNullString, vbNormalFocus
        Else
            ' Browser is not open
            Shell mstrBrowserPath & " " & strURL, vbNormalFocus
        End If
    
    Else
        ' No browser path was found
        InfoMsg "An internet connection is not available at this time." & vbNewLine & _
                "Please try again later." & vbNewLine & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    End If
    
End Sub

