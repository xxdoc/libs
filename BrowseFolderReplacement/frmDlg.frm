VERSION 5.00
Begin VB.Form frmDlg 
   Caption         =   "Browse For Folder"
   ClientHeight    =   3975
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   150
      Top             =   90
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   405
      Left            =   5070
      TabIndex        =   10
      Top             =   60
      Width           =   1455
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   480
         Picture         =   "frmDlg.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   390
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   0
         Picture         =   "frmDlg.frx":0440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   390
      End
      Begin VB.CommandButton cmdNewFolder 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   990
         Picture         =   "frmDlg.frx":0880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1005
      Left            =   90
      TabIndex        =   4
      Top             =   2925
      Width           =   6540
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   465
         Left            =   3870
         TabIndex        =   7
         Top             =   495
         Width           =   2580
         Begin VB.CommandButton Command1 
            Caption         =   "Select"
            Height          =   405
            Left            =   1305
            TabIndex        =   9
            Top             =   0
            Width           =   1125
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            Height          =   405
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   1395
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Text            =   "supports drag and drop"
         Top             =   0
         Width           =   5040
      End
      Begin VB.Label Label1 
         Caption         =   "Path"
         Height          =   255
         Left            =   765
         TabIndex        =   6
         Top             =   90
         Width           =   435
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1485
      TabIndex        =   3
      Top             =   90
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   2310
      Left            =   90
      ScaleHeight     =   2250
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   540
      Width           =   1275
      Begin VB.Image imgMyDocs 
         Height          =   810
         Left            =   45
         Picture         =   "frmDlg.frx":0CC0
         Top             =   1215
         Width           =   1170
      End
      Begin VB.Image imgDesktop 
         Height          =   750
         Left            =   0
         Picture         =   "frmDlg.frx":3ECC
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   1485
      TabIndex        =   0
      Top             =   540
      Width           =   5025
   End
   Begin VB.Label Label2 
      Caption         =   "Drive"
      Height          =   195
      Left            =   900
      TabIndex        =   2
      Top             =   135
      Width           =   465
   End
End
Attribute VB_Name = "frmDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Note: we have modified the behavior of the dirlist control so that a single click
'      on an item selects it. this led to the bug below.
'
'8.27.16 - bugfix for visual misselect on automated double click thanks aurel
'          if you clicked on a subfolder that was half way down the sub folder list, and it contained a bunch
'          of subfolders, the list would compact to show the newly selected folder, but if the mouse was still over
'          one of its subfolders, that one would visually highlight (although not be active in .path property)
'          we fix that through some chicanery in Dir1_click

Private Declare Function SHAutoComplete Lib "shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const LEFTDOWN = &H2, LEFTUP = &H4, MIDDLEDOWN = &H20, MIDDLEUP = &H40, RIGHTDOWN = &H8, RIGHTUP = &H10
Private Const SHACF_FILESYS_DIRS = &H20

Private Enum vButtons
    vRightClick = 2
    vDoubleRight = 4
    vLeftClick = 8
    vDoubleLeft = 16
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Const LB_GETCURSEL = &H188
Const LB_ERR = -1
Const LB_GETITEMRECT    As Long = &H198&

Private Type RECT
    Bottom As Long
    Left As Long
    Right As Long
    Top As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private FolderName As String
Private ignoreAutomation As Boolean
Private history() As String
Private ignoreDriveChange As Boolean
Private pt As POINTAPI

Public Enum SpecialFolders
    
    sf_DESKTOP = &H0 '<desktop>
    'sf_INTERNET = &H1 'Internet Explorer (icon on desktop)
    sf_PROGRAMS = &H2 'Start Menu\Programs
    'sf_CONTROLS = &H3'My Computer\Control Panel
    'sf_PRINTERS = &H4'My Computer\Printers
    sf_PERSONAL = &H5 'My Documents
    sf_FAVORITES = &H6 '<user name>\Favourites
    sf_STARTUP = &H7 'Start Menu\Programs\Startup
    sf_RECENT = &H8 '<user name>\Recent
    sf_SENDTO = &H9 '<user name>\SendTo
    sf_BITBUCKET = &HA '<desktop>\Recycle Bin
    sf_STARTMENU = &HB '<user name>\Start Menu
'    sf_MYDOCUMENTS = &HC'logical "My Documents" desktop icon
    sf_MYMUSIC = &HD '"My Music" folder
    sf_MYVIDEO = &HE '"My Videos" folder
    sf_DESKTOPDIRECTORY = &H10 '<user name>\Desktop
    sf_DRIVES = &H11 'My Computer
    'sf_NETWORK = &H12'Network Neighborhood (My Network Places)
'    sf_NETHOOD = &H13'<user name>\nethood
    sf_FONTS = &H14 'windows\fonts
'    sf_TEMPLATES = &H15'templates
    sf_COMMON_STARTMENU = &H16 'All Users\Start Menu
'    sf_COMMON_PROGRAMS = &H17 'All Users\Start Menu\Programs
    sf_COMMON_STARTUP = &H18 'All Users\Startup
    sf_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
    sf_APPDATA = &H1A '<user name>\Application Data
'    sf_PRINTHOOD = &H1B'<user name>\PrintHood
    sf_LOCAL_APPDATA = &H1C '<user name>\Local Settings\Application Data (non roaming)
 '   sf_ALTSTARTUP = &H1D'non localized startup
    'non localized common startup
'    sf_COMMON_ALTSTARTUP = &H1E
'    sf_COMMON_FAVORITES = &H1F
'    sf_INTERNET_CACHE = &H20
'    sf_COOKIES = &H21
'    sf_HISTORY = &H22
    'All Users\Application Data
'    sf_COMMON_APPDATA = &H23
    sf_WINDOWS = &H24 'GetWindowsDirectory()
    sf_SYSTEM = &H25 'GetSystemDirectory()
    sf_PROGRAM_FILES = &H26 'C:\Program Files
    sf_MYPICTURES = &H27 'C:\Program Files\My Pictures
    sf_PROFILE = &H28 'USERPROFILE
'    'x86 system directory on RISC
'    sf_SYSTEMX86 = &H29
'    'x86 C:\Program Files on RISC
'    sf_PROGRAM_FILESX86 = &H2A
'    'C:\Program Files\Common
'    sf_PROGRAM_FILES_COMMON = &H2B
'    'x86 Program Files\Common on RISC
'    sf_PROGRAM_FILES_COMMONX86 = &H2C
'     'All Users\Templates
'    sf_COMMON_TEMPLATES = &H2D
'     'All Users\Documents
'    sf_COMMON_DOCUMENTS = &H2E
'    'All Users\Start Menu\Programs\Administrative Tools
'    sf_COMMON_ADMINTOOLS = &H2F
'    '<user name>\Start Menu\Programs\Administrative Tools
'    sf_ADMINTOOLS = &H30
'    'Network and Dial-up Connections
'    sf_CONNECTIONS = &H31
'    'All Users\My Music
'    sf_COMMON_MUSIC = &H35
'    'All Users\My Pictures
'    sf_COMMON_PICTURES = &H36
'    'All Users\My Video
'    sf_COMMON_VIDEO = &H37
'    'Resource Directory
'    sf_RESOURCES = &H38
'    'Localized Resource Directory
'    sf_RESOURCES_LOCALIZED = &H39
'    'Links to All Users OEM specific apps
'    sf_COMMON_OEM_LINKS = &H3A
'    'USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
'    sf_CDBURN_AREA = &H3B
'    'unused                                      = &H3C
'    'Computers Near Me (computered from Workgroup membership)
'    sf_COMPUTERSNEARME = &H3D
End Enum


Private Sub cmdNewFolder_Click()
    Dim fName As String
    fName = InputBox("Create new folder in: " & vbCrLf & vbCrLf & Dir1.path)
    If Len(fName) = 0 Then Exit Sub
    On Error Resume Next
    MkDir Dir1.path & "\" & fName
    If Err.Number <> 0 Then
        MsgBox Err.Description
    Else
        Text1 = Dir1.path & "\" & fName
        'Dir1.Refresh
    End If
End Sub

Private Sub Command1_Click()
    FolderName = Text1
    Me.Visible = False
End Sub

Private Sub Command2_Click()
    FolderName = Empty
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim tmp As String
    Do
        tmp = pop(history)
        If Len(tmp) = 0 Then Exit Do
    Loop While tmp = Text1
    If Len(tmp) > 0 Then Text1 = tmp
End Sub

Private Sub Command4_Click()
    Dim tmp As String
    On Error Resume Next
    tmp = GetParentFolder(Text1)
    If Len(tmp) > 0 Then Text1 = tmp
End Sub

Private Sub Dir1_Change()
    Text1 = Dir1.path
    push history, Text1
    Debug.Print "Adding: " & Text1
    SyncDrive
End Sub

Private Sub Dir1_Click()
    
    On Error Resume Next
    Dim selitem As Long
    Dim udtRECT As RECT
    
    If ignoreAutomation Then
        'Debug.Print "ignored"
        Exit Sub
    End If
         
    ignoreAutomation = True
    
    'double click the entry under the mouse
    MouseClick vDoubleLeft

    'get the selected item index (Dir1.ListIndex control property is not yet set)
    selitem = SendMessage(Dir1.hWnd, LB_GETCURSEL, ByVal CLng(0), ByVal CLng(0))
    'Me.Caption = selitem & " " & Dir1.List(selitem) & " index:" & Dir1.ListIndex
    
    'save the current mouse position
    GetCursorPos pt
    
    'get rectangle for the selected item..
    SendMessage Dir1.hWnd, LB_GETITEMRECT, ByVal CLng(selitem - 1), udtRECT
    'Me.Caption = Me.Caption & " " & udtRECT.Left & " " & udtRECT.Top
    
    'now we move the mouse to the selected item and click the item once
    MoveMouseCursor udtRECT.Left, udtRECT.Top, Dir1.hWnd
    MouseClick vLeftClick
    
    'we use a timer to give it a slight delay and ensure it doesnt become a feedback loop
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    SetCursorPos pt.x, pt.Y
    ignoreAutomation = False
End Sub

Sub MoveMouseCursor(ByVal x As Long, ByVal Y As Long, Optional ByVal hWnd As Long)
    If hWnd = 0 Then
        SetCursorPos x, Y
    Else
        Dim lpPoint As POINTAPI
        lpPoint.x = x
        lpPoint.Y = Y
        ClientToScreen hWnd, lpPoint
        SetCursorPos lpPoint.x, lpPoint.Y
    End If
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    If ignoreDriveChange Then Exit Sub
    Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Text1 = GetSpecialFolder(sf_DESKTOP)
    SHAutoComplete Text1.hWnd, SHACF_FILESYS_DIRS
End Sub

Function BrowseForFolder(Optional initDir As String, Optional specialFolder As SpecialFolders = -1, Optional owner As Form = Nothing) As String

    If specialFolder <> -1 Then
        Text1 = GetSpecialFolder(specialFolder)
    ElseIf FolderExists(initDir) Then
        Text1 = initDir
    End If
    
    Me.Show 1, owner  'modal does not return until cancel or save hit..
    BrowseForFolder = FolderName
    Unload Me
    
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Width = Me.Width
    Frame2.Left = Frame1.Width - Frame2.Width - 200
    Frame1.Top = Me.Height - Frame1.Height - 400
    Frame3.Left = Me.Width - Frame3.Width - 200
    Dir1.Height = Me.Height - Frame1.Height - 1000
    Dir1.Width = Me.Width - Dir1.Left - 200
    Text1.Width = Dir1.Width
    Picture1.Height = Dir1.Height
    Drive1.Width = Me.Width - Dir1.Left - Frame3.Width - 400
End Sub

Private Sub imgDesktop_Click()
    Dir1.path = GetSpecialFolder(sf_DESKTOP)
End Sub

Private Sub imgMyComp_Click()
    Dir1.path = GetSpecialFolder(sf_DRIVES)
End Sub

Private Sub imgMyDocs_Click()
    Dir1.path = GetSpecialFolder(sf_PERSONAL)
End Sub

Private Sub Text1_Change()
    If FolderExists(Text1) And Text1 <> Dir1.path Then
        Dir1.path = Text1
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Dim f As String
    f = Data.Files(1)
    If FileExists(f) Then Text1 = GetParentFolder(f)
    If FolderExists(f) Then Text1 = f
End Sub

Private Function SyncDrive()
    On Error Resume Next
    Dim drive_letter As String, i As Long
    
    ignoreDriveChange = True
        
    drive_letter = LCase(VBA.Left(Text1, 2))
    For i = 0 To Drive1.ListCount
        If LCase(Left(Drive1.List(i), 2)) = drive_letter Then
            If Drive1.ListIndex <> i Then Drive1.ListIndex = i
            Exit For
        End If
    Next
        
    ignoreDriveChange = False
    
End Function

Private Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Function GetParentFolder(path) As String
    Dim tmp, ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
    If Right(GetParentFolder, 1) = ":" Then GetParentFolder = GetParentFolder & "\"
End Function

Private Function GetSpecialFolder(sf As SpecialFolders) As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, sf, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      GetSpecialFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
  
End Function

Private Sub MouseClick(Optional b As vButtons)

    Select Case b
    
        Case vRightClick
            mouse_event RIGHTDOWN, 0&, 0&, 0&, 0&
            mouse_event RIGHTUP, 0&, 0&, 0&, 0&
        
        Case vDoubleRight
            mouse_event RIGHTDOWN, 0&, 0&, 0&, 0&
            mouse_event RIGHTUP, 0&, 0&, 0&, 0&
            DoEvents
            mouse_event RIGHTDOWN, 0&, 0&, 0&, 0&
            mouse_event RIGHTUP, 0&, 0&, 0&, 0&
        
        Case vLeftClick
            mouse_event LEFTDOWN, 0&, 0&, 0&, 0&
            mouse_event LEFTUP, 0&, 0&, 0&, 0&
        
        Case vDoubleLeft
            mouse_event LEFTDOWN, 0&, 0&, 0&, 0&
            mouse_event LEFTUP, 0&, 0&, 0&, 0&
            DoEvents
            mouse_event LEFTDOWN, 0&, 0&, 0&, 0&
            mouse_event LEFTUP, 0&, 0&, 0&, 0&
    
    End Select
End Sub




Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function pop(ary) 'this modifies parent ary obj
        
    If AryIsEmpty(ary) Then Exit Function
    
    pop = ary(UBound(ary))
    
    If UBound(ary) = 0 Then
        Erase ary
    Else
        ReDim Preserve ary(UBound(ary) - 1)
    End If

End Function

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
