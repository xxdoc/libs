VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "API Demo "
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Play AVI from Memory"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3660
      TabIndex        =   14
      Top             =   2460
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5250
      TabIndex        =   12
      Text            =   "AVI"
      Top             =   1830
      Width           =   2370
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2625
      TabIndex        =   10
      Text            =   "#150"
      Top             =   1830
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2625
      TabIndex        =   8
      Text            =   "shell32.dll"
      Top             =   1140
      Width           =   4995
   End
   Begin VB.PictureBox Picture1 
      Height          =   2085
      Left            =   45
      ScaleHeight     =   2025
      ScaleWidth      =   2445
      TabIndex        =   7
      Top             =   900
      Width           =   2505
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   30
      ScaleHeight     =   1755
      ScaleWidth      =   7545
      TabIndex        =   1
      Top             =   3105
      Width           =   7605
      Begin VB.Label l6 
         BackStyle       =   0  'Transparent
         Caption         =   "API"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   795
         Width           =   1830
      End
      Begin VB.Label l7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0000
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   60
         TabIndex        =   4
         Top             =   1020
         Width           =   7470
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "This demo will show you how to play a resource AVI from memory without creating any temp file on disk."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   60
         TabIndex        =   2
         Top             =   315
         Width           =   7470
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "Example Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Label l8 
         BackStyle       =   0  'Transparent
         Caption         =   "www.binaryworld.net"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   5040
         MouseIcon       =   "Form1.frx":0097
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   60
         Width           =   2565
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   15
         Picture         =   "Form1.frx":03A1
         Top             =   0
         Width           =   4875
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5265
      TabIndex        =   13
      Top             =   1560
      Width           =   1830
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   1830
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resource File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   870
      Width           =   1830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim hRsrc As Long
    Dim hGlobal As Long
    Dim lpString As String
    Dim strCmd As String, ret As Long
    Dim nbuf As Long


    '//Loads the resource
    '//Change the filename argument in the next line to the path and
    '//filename of your resource dll file.
    'hInst = LoadLibrary(App.Path & "\test.exe")    '    <<<<
    hInst = LoadLibrary(Text1)    '    <<<<

    '//Identifier for the AVI in the resource file from the resource.h
    '//file of the resource dll, which must be preceded by '#'.
    'lpString = "#101"    '                                       <<<<
    lpString = Text2    '                                       <<<<

    'hRsrc = FindResource(hInst, lpString, "AVI")
    hRsrc = FindResource(hInst, lpString, Text3)

    hGlobal = LoadResource(hInst, hRsrc)
    lpData = LockResource(hGlobal)
    fileSize = SizeofResource(hInst, hRsrc)

    Call mmioInstallIOProc(MEY, AddressOf IOProc, MMIO_INSTALLPROC + MMIO_GLOBALPROC)
    nbuf = 256

    '//Close all opened MCI device for this app before running any new avi
    Call mciSendString("Close all", 0&, 0&, 0&)

    '//Play the AVI file
    strCmd = "open test.MEY+ type avivideo alias test parent " & Picture1.hWnd & " Style child"
    ret = mciSendString(strCmd, 0&, 0&, 0&)
    If ret > 0 Then ShowMCIError (ret)

    strCmd = "play test repeat"
    ret = mciSendString(strCmd, 0&, 0&, 0&)
    If ret > 0 Then ShowMCIError (ret)

    Call mmioInstallIOProc(MEY, vbNull, MMIO_REMOVEPROC)
    FreeLibrary hInst
End Sub

Private Sub Form_Load()
    Text1 = "Shell32.dll"
    Text2 = "#150"    '//This is predefined AVI in shell32.dll so we pick for demo
    Text3 = "AVI"
    Command1.Caption = "Play AVI from Memory"
End Sub

Private Sub l8_Click()
    Shell "explorer http://www.binaryworld.net", vbNormalFocus
End Sub
