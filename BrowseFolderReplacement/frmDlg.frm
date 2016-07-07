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
         TabIndex        =   8
         Top             =   495
         Width           =   2580
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   405
            Left            =   1305
            TabIndex        =   10
            Top             =   0
            Width           =   1125
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            Height          =   405
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   1395
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Text            =   "supports drag and drop"
         Top             =   0
         Width           =   5040
      End
      Begin VB.CommandButton cmdNewFolder 
         Caption         =   "New Folder"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   495
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Path"
         Height          =   255
         Left            =   765
         TabIndex        =   7
         Top             =   90
         Width           =   435
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1485
      TabIndex        =   3
      Top             =   90
      Width           =   5055
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
         Picture         =   "frmDlg.frx":0000
         Top             =   1215
         Width           =   1170
      End
      Begin VB.Image imgDesktop 
         Height          =   750
         Left            =   0
         Picture         =   "frmDlg.frx":320C
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

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private FolderName As String

Public Enum SpecialFolders
    '<desktop>
    CSIDL_DESKTOP = &H0
    'Internet Explorer (icon on desktop)
    CSIDL_INTERNET = &H1
    'Start Menu\Programs
    CSIDL_PROGRAMS = &H2
    'My Computer\Control Panel
    CSIDL_CONTROLS = &H3
    'My Computer\Printers
    CSIDL_PRINTERS = &H4
    'My Documents
    CSIDL_PERSONAL = &H5
    '<user name>\Favourites
    CSIDL_FAVORITES = &H6
    'Start Menu\Programs\Startup
    CSIDL_STARTUP = &H7
    '<user name>\Recent
    CSIDL_RECENT = &H8
    '<user name>\SendTo
    CSIDL_SENDTO = &H9
    '<desktop>\Recycle Bin
    CSIDL_BITBUCKET = &HA
    '<user name>\Start Menu
    CSIDL_STARTMENU = &HB
    'logical "My Documents" desktop icon
'    CSIDL_MYDOCUMENTS = &HC
    '"My Music" folder
    CSIDL_MYMUSIC = &HD
    '"My Videos" folder
    CSIDL_MYVIDEO = &HE
    '<user name>\Desktop
    CSIDL_DESKTOPDIRECTORY = &H10
    'My Computer
    CSIDL_DRIVES = &H11
    'Network Neighborhood (My Network Places)
    CSIDL_NETWORK = &H12
    '<user name>\nethood
'    CSIDL_NETHOOD = &H13
    'windows\fonts
    CSIDL_FONTS = &H14
    'templates
'    CSIDL_TEMPLATES = &H15
    'All Users\Start Menu
    CSIDL_COMMON_STARTMENU = &H16
    'All Users\Start Menu\Programs
'    CSIDL_COMMON_PROGRAMS = &H17
    'All Users\Startup
    CSIDL_COMMON_STARTUP = &H18
    'All Users\Desktop
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    '<user name>\Application Data
    CSIDL_APPDATA = &H1A
    '<user name>\PrintHood
'    CSIDL_PRINTHOOD = &H1B
    '<user name>\Local Settings\Application Data (non roaming)
    CSIDL_LOCAL_APPDATA = &H1C
    'non localized startup
 '   CSIDL_ALTSTARTUP = &H1D
    'non localized common startup
'    CSIDL_COMMON_ALTSTARTUP = &H1E
'    CSIDL_COMMON_FAVORITES = &H1F
'    CSIDL_INTERNET_CACHE = &H20
'    CSIDL_COOKIES = &H21
'    CSIDL_HISTORY = &H22
    'All Users\Application Data
'    CSIDL_COMMON_APPDATA = &H23
    'GetWindowsDirectory()
    CSIDL_WINDOWS = &H24
    'GetSystemDirectory()
    CSIDL_SYSTEM = &H25
    'C:\Program Files
    CSIDL_PROGRAM_FILES = &H26
    'C:\Program Files\My Pictures
    CSIDL_MYPICTURES = &H27
    'USERPROFILE
    CSIDL_PROFILE = &H28
'    'x86 system directory on RISC
'    CSIDL_SYSTEMX86 = &H29
'    'x86 C:\Program Files on RISC
'    CSIDL_PROGRAM_FILESX86 = &H2A
'    'C:\Program Files\Common
'    CSIDL_PROGRAM_FILES_COMMON = &H2B
'    'x86 Program Files\Common on RISC
'    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
'     'All Users\Templates
'    CSIDL_COMMON_TEMPLATES = &H2D
'     'All Users\Documents
'    CSIDL_COMMON_DOCUMENTS = &H2E
'    'All Users\Start Menu\Programs\Administrative Tools
'    CSIDL_COMMON_ADMINTOOLS = &H2F
'    '<user name>\Start Menu\Programs\Administrative Tools
'    CSIDL_ADMINTOOLS = &H30
'    'Network and Dial-up Connections
'    CSIDL_CONNECTIONS = &H31
'    'All Users\My Music
'    CSIDL_COMMON_MUSIC = &H35
'    'All Users\My Pictures
'    CSIDL_COMMON_PICTURES = &H36
'    'All Users\My Video
'    CSIDL_COMMON_VIDEO = &H37
'    'Resource Directory
'    CSIDL_RESOURCES = &H38
'    'Localized Resource Directory
'    CSIDL_RESOURCES_LOCALIZED = &H39
'    'Links to All Users OEM specific apps
'    CSIDL_COMMON_OEM_LINKS = &H3A
'    'USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
'    CSIDL_CDBURN_AREA = &H3B
'    'unused                                      = &H3C
'    'Computers Near Me (computered from Workgroup membership)
'    CSIDL_COMPUTERSNEARME = &H3D
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
        Dir1.Refresh
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

Private Sub Dir1_Change()
    Text1 = Dir1.path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Text1 = GetSpecialFolder(CSIDL_DESKTOP)
End Sub

Function BrowseForFolder(Optional initDir As String, Optional specialFolder As SpecialFolders = -1, Optional owner As Form = Nothing) As String

    If specialFolder <> -1 Then
        Text1 = GetSpecialFolder(specialFolder)
    ElseIf FolderExists(initDir) Then
        Text1 = initDir
    End If
    
    Me.Show 1, owner 'modal does not return until cancel or save hit..
    BrowseForFolder = FolderName
    Unload Me
    
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Width = Me.Width
    Frame2.Left = Frame1.Width - Frame2.Width - 200
    Frame1.Top = Me.Height - Frame1.Height - 400
    Dir1.Height = Me.Height - Frame1.Height - 1000
    Dir1.Width = Me.Width - Dir1.Left - 200
    Text1.Width = Dir1.Width
    Picture1.Height = Dir1.Height
    Drive1.Width = Dir1.Width
End Sub

Private Sub imgDesktop_Click()
    Dir1.path = GetSpecialFolder(CSIDL_DESKTOP)
End Sub

Private Sub imgMyComp_Click()
    Dir1.path = GetSpecialFolder(CSIDL_DRIVES)
End Sub

Private Sub imgMyDocs_Click()
    Dir1.path = GetSpecialFolder(CSIDL_PERSONAL)
End Sub

Private Sub Text1_Change()
    If FolderExists(Text1) And Text1 <> Dir1.path Then Dir1.path = Text1
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim f As String
    f = Data.Files(1)
    If FileExists(f) Then Text1 = GetParentFolder(f)
    If FolderExists(f) Then Text1 = f
End Sub




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


