VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   1920
   ClientTop       =   2265
   ClientWidth     =   11250
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11250
   Visible         =   0   'False
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Index           =   1
      Left            =   90
      TabIndex        =   29
      Top             =   1440
      Width           =   11070
      Begin VB.CheckBox chkShowPwd 
         Caption         =   "chkShowPwd"
         Height          =   420
         Left            =   6975
         TabIndex        =   49
         Top             =   5310
         Width           =   1725
      End
      Begin VB.CheckBox chkExtraInfo 
         Caption         =   "chkExtraInfo"
         Height          =   330
         Left            =   9045
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   855
         Width           =   1875
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy to clipboard"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7065
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5175
         Width           =   1770
      End
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "txtPwd"
         Top             =   5340
         Width           =   6660
      End
      Begin VB.Frame fraEncrypt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Index           =   1
         Left            =   8955
         TabIndex        =   36
         Top             =   1170
         Width           =   1995
         Begin VB.ComboBox cboKeyMix 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030A
            Left            =   675
            List            =   "frmMain.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   4050
            Width           =   615
         End
         Begin VB.ComboBox cboBlockLength 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":030E
            Left            =   135
            List            =   "frmMain.frx":0310
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2520
            Width           =   1695
         End
         Begin VB.ComboBox cboKeyLength 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":0312
            Left            =   135
            List            =   "frmMain.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1845
            Width           =   1695
         End
         Begin VB.ComboBox cboHash 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":0316
            Left            =   135
            List            =   "frmMain.frx":0318
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1125
            Width           =   1695
         End
         Begin VB.ComboBox cboEncrypt 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   405
            Width           =   1695
         End
         Begin VB.ComboBox cboRounds 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmMain.frx":031A
            Left            =   675
            List            =   "frmMain.frx":031C
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3195
            Width           =   615
         End
         Begin VB.Label lblKeyMix 
            Caption         =   "Number of rounds to mix primary key"
            Height          =   480
            Left            =   180
            TabIndex        =   50
            Top             =   3645
            Width           =   1650
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Block Length"
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   41
            Top             =   2295
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Password Key Length"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   40
            Top             =   1605
            Width           =   1620
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Encryption Algorithm"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   39
            Top             =   180
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Password Hash Algo"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   38
            Top             =   900
            Width           =   1605
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Rounds of encryption"
            Height          =   240
            Index           =   5
            Left            =   180
            TabIndex        =   37
            Top             =   2970
            Width           =   1755
         End
      End
      Begin VB.Frame fraEncrypt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4980
         Index           =   0
         Left            =   105
         TabIndex        =   32
         Top             =   105
         Width           =   8745
         Begin VB.TextBox txtInputString 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "frmMain.frx":031E
            Top             =   360
            Width           =   8520
         End
         Begin VB.TextBox txtInputFile 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "txtInputFile"
            Top             =   360
            Visible         =   0   'False
            Width           =   7890
         End
         Begin VB.PictureBox picProgressBar 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   90
            ScaleHeight     =   240
            ScaleWidth      =   8460
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   4545
            Width           =   8520
         End
         Begin VB.TextBox txtOutput 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "frmMain.frx":032F
            Top             =   3900
            Width           =   8520
         End
         Begin VB.CommandButton cmdBrowse 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8100
            Picture         =   "frmMain.frx":0339
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   35
            Top             =   3660
            Width           =   5820
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   34
            Top             =   135
            Width           =   5820
         End
      End
      Begin VB.PictureBox picDataType 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   8985
         ScaleHeight     =   645
         ScaleWidth      =   1920
         TabIndex        =   30
         Top             =   225
         Width           =   1920
         Begin VB.OptionButton optDataType 
            Caption         =   "String data"
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   6
            Top             =   330
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optDataType 
            Caption         =   "File"
            Height          =   240
            Index           =   1
            Left            =   1275
            TabIndex        =   7
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Type"
            Height          =   240
            Index           =   0
            Left            =   495
            TabIndex        =   31
            Top             =   45
            Width           =   990
         End
      End
      Begin VB.Label lblPwd 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   180
         TabIndex        =   42
         Top             =   5100
         Width           =   6525
      End
   End
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Index           =   0
      Left            =   90
      TabIndex        =   25
      Top             =   1440
      Width           =   11070
      Begin VB.ComboBox cboRandom 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7815
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   225
         Width           =   2985
      End
      Begin VB.TextBox txtRandom 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "frmMain.frx":043B
         Top             =   630
         Width           =   10860
      End
      Begin VB.Label lblAlgo 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Data Types"
         Height          =   240
         Index           =   7
         Left            =   6315
         TabIndex        =   28
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Index           =   2
      Left            =   9765
      Picture         =   "frmMain.frx":0445
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Display credits"
      Top             =   7335
      Width           =   640
   End
   Begin VB.CommandButton cmdChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Index           =   1
      Left            =   9045
      Picture         =   "frmMain.frx":074F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7335
      Width           =   640
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   10440
      Picture         =   "frmMain.frx":0B91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   300
      Picture         =   "frmMain.frx":0E9B
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   180
      Width           =   480
   End
   Begin VB.CommandButton cmdChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Index           =   3
      Left            =   10485
      Picture         =   "frmMain.frx":11A5
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Terminate this application"
      Top             =   7335
      Width           =   640
   End
   Begin VB.CommandButton cmdChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Index           =   0
      Left            =   9045
      Picture         =   "frmMain.frx":14AF
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7335
      Width           =   640
   End
   Begin VB.Frame fraChoice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   18
      Top             =   810
      Width           =   5895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         ScaleHeight     =   375
         ScaleWidth      =   5670
         TabIndex        =   23
         Top             =   150
         Width           =   5670
         Begin VB.OptionButton optChoice 
            Caption         =   "Encrypt/Decrypt"
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   0
            Top             =   75
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "CRC-32"
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   1
            Top             =   75
            Width           =   1035
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Hash"
            Height          =   315
            Index           =   2
            Left            =   3150
            TabIndex        =   2
            Top             =   75
            Width           =   810
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Random data"
            Height          =   315
            Index           =   3
            Left            =   4185
            TabIndex        =   3
            Top             =   75
            Width           =   1350
         End
      End
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
      Left            =   3330
      TabIndex        =   48
      Top             =   135
      Width           =   4875
   End
   Begin VB.Label lblOperSystem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblOperSystem"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4230
      TabIndex        =   47
      Top             =   7560
      Width           =   4650
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      Height          =   585
      Left            =   135
      TabIndex        =   24
      Top             =   7380
      Width           =   2670
   End
   Begin VB.Label lblEncryptMsg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblEncryptMsg"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   7290
      TabIndex        =   22
      Top             =   765
      Width           =   3690
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
      Height          =   330
      Left            =   5190
      TabIndex        =   19
      Top             =   540
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Routine:       frmMain
'
' ===========================================================================
' To build binary zero filled test files, I use the following method for
' files that are one gigabyte or less in size.
'
' The character being used (ex: 0) only designates the last ASCII value
' in the file. All previous values will be null (ASCII 0) values. This
' method may be used to create files up to one GB in size. Takes about
' one second.
'
'    Const MB_5 As Long = &H500000   ' 5242880 bytes
'
'    hFile = FreeFile                                                 ' Get first free file handle
'    Open "C:\Temp\MB_5_Test.dat" For Binary Access Write As #hFile   ' Open for binary writing
'    Put #hFile, MB_5, Chr$(0)                                        ' Fill with binary 0's
'    Close #hFile                                                     ' Close file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Used clsAPI_Hash class to replace cMD4, cMD5, cSHA1 and
'                cSHA2 classes.
'              - Removed RipeMD classes because they are considered weak.
' 12-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug with selecting number of encryption rounds.
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
'              See ProgressBar() routine.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made updates as per Joe Sova's suggestions
'                1. Locked and unlocked input/output textboxes during
'                   processing to reduce flicker.
'                2. Hide progressbar during string processing.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated Cypher_Processing() routine to reference new property
'              CreateNewFile(). If you want to overwrite file being encrypted
'              or decrypted.
' 06-Sep-2013  Kenneth Ives  kenaso@tx.rr.com
'              Changed txtKeyMix textbox to cboKeyMix combobox
' 17-Sep-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added more security for password. See chkShowPwd_Click(),
'              Cipher_Processing(), txtPwd.GetFocus(), txtPwd.Lost_Focus()
'              routines for updates and documentation.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

' ***************************************************************************
' Module API Declares
' ***************************************************************************
  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

  ' Reduce flicker while loading a control
  ' Lock the control to prevent redrawing
  '     Syntax:  LockWindowUpdate frmMain.hWnd
  ' Unlock the control
  '     Syntax:  LockWindowUpdate 0&
  Private Declare Function LockWindowUpdate Lib "user32" _
          (ByVal hWnd As Long) As Long

' ***************************************************************************
' Module Variables
' Variable name:     mlngCipher
' Naming standard:   m lng Cipher
'                    - --- ---------
'                    |  |    |______ Variable subname
'                    |  |___________ Data type (Long)
'                    |______________ Module level designator
'
' ***************************************************************************
  Private mlngKeyMix        As Long
  Private mlngRounds        As Long
  Private mlngHashAlgo      As Long
  Private mlngKeyLength     As Long
  Private mlngCipherAlgo    As Long
  Private mlngBlockLength   As Long
  Private mlngPwdLength_Min As Long
  Private mlngPwdLength_Max As Long
  Private mabytPwd()        As Byte
  Private mstrFolder        As String
  Private mstrFilename      As String
  Private mblnHash          As Boolean
  Private mblnPrng          As Boolean
  Private mblnCRC32         As Boolean
  Private mblnCipher        As Boolean
  Private mblnLoading       As Boolean
  Private mblnShowPwd       As Boolean
  Private mblnStringData    As Boolean
  Private mblnRetLowercase  As Boolean
  Private mblnCreateNewFile As Boolean
  Private mobjPrng          As kiCrypt.cPrng
  Private mobjKeyEdit       As cKeyEdit

  Private WithEvents mobjHash   As kiCrypt.cHash
Attribute mobjHash.VB_VarHelpID = -1
  Private WithEvents mobjCRC32  As kiCrypt.cCRC32
Attribute mobjCRC32.VB_VarHelpID = -1
  Private WithEvents mobjCipher As kiCrypt.cCipher
Attribute mobjCipher.VB_VarHelpID = -1

Private Sub cboEncrypt_Click()
    
    Dim lngIdx As Long
    
    If mblnLoading Then
        Exit Sub
    End If
                
    mlngCipherAlgo = cboEncrypt.ListIndex
    
    Select Case mlngCipherAlgo
           Case 1   ' Base64
                lblEncryptMsg.Visible = False
                lblPwd.Visible = False
                txtPwd.Visible = False
                    
                ' Prepare combo boxes
                cboHash.Enabled = False
                cboHash.ForeColor = vbGrayText
                cboKeyLength.Enabled = False
                cboKeyLength.ForeColor = vbGrayText
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = False
                cboRounds.ForeColor = vbGrayText
                
                mlngKeyLength = 0
                mlngBlockLength = 0
                mlngRounds = 0
           
           Case 0, 3, 5, 6 ' ArcFour, GOST, Serpent, Skipjack
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 128 To 416 Step 32
                        .AddItem CStr(lngIdx) & " bits"   ' 32 bit increments
                    Next lngIdx
                    
                    For lngIdx = 448 To 1024 Step 64
                        .AddItem CStr(lngIdx) & " bits"   ' 64 bit increments
                    Next lngIdx
                    .ListIndex = 0
                End With
    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboRounds_Click
                        
           Case 2, 7  ' Blowfish, TwoFish
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 32 To 448 Step 32
                        .AddItem CStr(lngIdx) & " bits"   ' 32 bit increments
                    Next lngIdx
                    .ListIndex = 0
                End With
    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = False
                cboBlockLength.ForeColor = vbGrayText
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboRounds_Click
                
           Case 4  ' Rijndael
                ' Prepare combo boxes
                With cboKeyLength
                    .Clear
                    For lngIdx = 128 To 256 Step 32
                        .AddItem CStr(lngIdx) & " bits"
                    Next lngIdx
                    .ListIndex = 0
                End With
                    
                cboHash.Enabled = True
                cboHash.ForeColor = vbBlack
                cboKeyLength.Enabled = True
                cboKeyLength.ForeColor = vbBlack
                cboBlockLength.Enabled = True
                cboBlockLength.ForeColor = vbBlack
                cboRounds.Enabled = True
                cboRounds.ForeColor = vbBlack
                
                lblEncryptMsg.Visible = True
                lblPwd.Visible = True
                txtPwd.Visible = True
                cboBlockLength_Click
                cboRounds_Click
    End Select
                
    Select Case mlngCipherAlgo
    
           Case 2, 3, 5, 7   ' Blowfish, GOST, Serpent, Twofish
                lblKeyMix.Enabled = True
                lblKeyMix.Visible = True
                cboKeyMix.Enabled = True
                cboKeyMix.Visible = True
                cboKeyMix.Text = mlngKeyMix
                
           Case Else  ' Hide textbox and label
                lblKeyMix.Enabled = False
                lblKeyMix.Visible = False
                cboKeyMix.Enabled = False
                cboKeyMix.Visible = False
    End Select
                
End Sub

Private Sub cboHash_Click()

    Dim lngIdx As Long
    
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    mlngHashAlgo = cboHash.ListIndex
    
    ' If performing encryption then leave
    If cboEncrypt.Enabled Then
        Exit Sub
    End If
    
    ' Multiple hash rounds only available
    ' for testing hash output values
    Select Case mlngHashAlgo
           
           ' MD2, MD4, MD5, SHA-1, SHA-256, SHA-384, SHA-512
           ' Whirlpool-224, Whirlpool-256, Whirlpool-384, Whirlpool-512
           Case 0 To 6, 14 To 17
                With cboRounds
                    .Clear
                    For lngIdx = 1 To 10
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 0  ' Default rounds = 1
                End With
                    
           Case 7 To 13    ' Tiger3 family
                With cboRounds
                    .Clear
                    For lngIdx = 3 To 15
                        .AddItem CStr(lngIdx)
                    Next lngIdx
                    .ListIndex = 3   ' Default rounds = 6
                End With
    End Select
    
    mlngRounds = CLng(TrimStr(Left$(cboRounds.Text, 2)))
    
End Sub

Private Sub cboKeyLength_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    ' Select user defined key length
    Select Case mlngCipherAlgo
           Case eCIPHER_BASE64
                mlngKeyLength = 0
                
           Case eCIPHER_BLOWFISH, eCIPHER_TWOFISH
                mlngKeyLength = CLng(TrimStr(Left$(cboKeyLength.Text, 3)))
                
           Case Else
                mlngKeyLength = CLng(TrimStr(Left$(cboKeyLength.Text, 4)))
    End Select
    
End Sub

Private Sub cboBlockLength_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    ' Select user defined block size (RIJNDAEL only)
    Select Case mlngCipherAlgo
           Case eCIPHER_RIJNDAEL: mlngBlockLength = CLng(TrimStr(Left$(cboBlockLength.Text, 3)))
           Case Else:             mlngBlockLength = 0
    End Select
    
End Sub

Private Sub cboKeyMix_Click()

    If mblnLoading Then
        Exit Sub
    End If
    
    mlngKeyMix = CLng(TrimStr(cboKeyMix.Text))
    
    ' Test min and max range
    Select Case mlngKeyMix
           Case Is < 1: mlngKeyMix = 1
           Case Is > 5: mlngKeyMix = 5
    End Select
    
End Sub

Private Sub cboRounds_Click()

    If mblnLoading Then
        Exit Sub
    End If
    
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False

    If cboEncrypt.Enabled Then
        
        ' Select number of rounds for encryption
        Select Case mlngCipherAlgo
               Case eCIPHER_BASE64: mlngRounds = 0
               Case Else:           mlngRounds = CLng(TrimStr(Left$(cboRounds.Text, 2)))
        End Select
    
    Else
        ' Perform hash testing only
        mlngRounds = CLng(TrimStr(Left$(cboRounds.Text, 2)))
    End If
    
End Sub

Private Sub cboRandom_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    txtRandom.Text = vbNullString
    RndData_Processing cboRandom.ListIndex

End Sub

Private Sub chkExtraInfo_Click()
    
    ' Designates if input file is to be overwritten after
    ' performing encryption/decryption.
    '
    ' Cipher processing
    '   String processing:
    '       Checked   - Return encrypted string data in lowercase format
    '       Unchecked - Return encrypted string data in uppercase format
    '   File processing:
    '       Checked   - Create new file to hold encrypted/decrypted data
    '       Unchecked - Overwrite input file after encryption/decryption
    '
    ' Hash processing
    ' Checked   - Return hashed data in lowercase format
    ' Unchecked - Return hashed data in uppercase format
    
    If mblnLoading Then
        Exit Sub
    End If
    
    If mblnCipher Then
        If mblnStringData Then
            mblnRetLowercase = IIf(chkExtraInfo.Value = vbChecked, True, False)
        Else
            mblnCreateNewFile = IIf(chkExtraInfo.Value = vbChecked, True, False)
        End If
        
    ElseIf mblnHash Then
        mblnRetLowercase = IIf(chkExtraInfo.Value = vbChecked, True, False)
    End If
    
End Sub

Private Sub chkShowPwd_Click()
    
    ' Toggle showing password
    mblnShowPwd = IIf(chkShowPwd.Value = vbChecked, True, False)
    txtPwd.Text = vbNullString   ' Empty password textbox
    
    ' Replace any data with asteriks.
    ' Just in case someone is using a
    ' utility to read passwords behind
    ' the asteriks.
    If mblnShowPwd Then
        ' Has password array been initialized
        If CBool(IsArrayInitialized(mabytPwd)) Then
            ' Does password array hold any data
            If UBound(mabytPwd) > 0 Then
                txtPwd.Text = String$(UBound(mabytPwd) + 1, "*")   ' Only show asteriks
            End If
        End If
    Else
        ' Has password array been initialized
        If CBool(IsArrayInitialized(mabytPwd)) Then
            ' Does password array hold any data
            If UBound(mabytPwd) > (-1) Then
                txtPwd.Text = String$(5, "*")   ' Show only five asteriks
            End If
        End If
    End If
    
End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear                   ' clear the clipboard
    Clipboard.SetText txtOutput.Text  ' load clipboard with textbox data

End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub cmdBrowse_Click()
    
    Dim strHold As String
    
    mstrFilename = vbNullString
    txtInputFile.Text = vbNullString
    txtOutput.Text = vbNullString
    cmdCopy.Enabled = False
    strHold = mstrFolder
    
    ' Display "File Open" dialog box
    frmMain.Hide                           ' Hide main form
    SetStartingFolder mstrFolder           ' User defined starting folder
    mstrFilename = ShowFileOpen(frmMain)   ' Capture path\name of file
    frmMain.Show                           ' Show main form
        
    mstrFilename = TrimStr(mstrFilename)   ' Remove unwanted leading/trailing characters
    
    If Len(mstrFilename) > 0 Then
        txtInputFile.Text = ShrinkToFit(mstrFilename, 70)  ' Original file name
        mstrFolder = GetFullPath(mstrFilename)             ' Capture new path
    Else
        mstrFolder = strHold
    End If
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim lngIdx As Long

    Select Case Index
    
           Case 0  ' OK button
                Screen.MousePointer = vbHourglass
                gblnStopProcessing = False
                SetDllProcessingFlag
                ResetProgressBar
                cmdChoice_GotFocus 1
                
                ' Lock controls
                With frmMain
                    .cmdChoice(2).Enabled = False
                    .cmdChoice(3).Enabled = False
                    .fraChoice.Enabled = False
                    .fraEncrypt(1).Enabled = False
                    .picDataType.Enabled = False
                    .txtPwd.Enabled = False
                    .txtOutput.Text = vbNullString
                    .Refresh
                End With
                
                If mblnCipher Then
                    ' Temporarily lock textbox while processing.
                    ' This will speed things up and reduce flicker.
                    If mblnStringData Then
                        LockWindowUpdate frmMain.txtInputString.hWnd
                    Else
                        LockWindowUpdate frmMain.txtOutput.hWnd
                    End If
                    
                    Cipher_Processing     ' Encrypt/Decrypt data
                    LockWindowUpdate 0&   ' Unlock txtOutput textbox after processing
                    DoEvents
                End If
                
                If mblnCRC32 Then
                    ' Temporarily lock textbox while processing.
                    ' This will speed things up and reduce flicker.
                    If mblnStringData Then
                        LockWindowUpdate frmMain.txtInputString.hWnd
                    Else
                        LockWindowUpdate frmMain.txtOutput.hWnd
                    End If
                        
                    CRC32_Processing       ' Calculate CRC
                    LockWindowUpdate 0&    ' Unlock txtOutput textbox after processing
                    DoEvents
                End If
                
                If mblnHash Then
                    ' Temporarily lock textbox while processing.
                    ' This will speed things up and reduce flicker.
                    If mblnStringData Then
                        LockWindowUpdate frmMain.txtInputString.hWnd
                    Else
                        LockWindowUpdate frmMain.txtOutput.hWnd
                    End If
                        
                    Hash_Processing        ' Perform hash
                    LockWindowUpdate 0&    ' Unlock txtOutput textbox after processing
                    DoEvents
                End If
                
                If mblnPrng Then
                    lngIdx = cboRandom.ListIndex
                    RndData_Processing lngIdx
                End If

                DoEvents
                UpdateRegistry
                
                ' Unlock controls
                With frmMain
                    .cmdChoice(2).Enabled = True
                    .cmdChoice(3).Enabled = True
                    .fraChoice.Enabled = True
                    .fraEncrypt(1).Enabled = True
                    .picDataType.Enabled = True
                    .txtPwd.Enabled = True
                End With
                
                ResetProgressBar
                cmdChoice_GotFocus 0
                
           Case 1  ' Stop button
                DoEvents
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                SetDllProcessingFlag
                DoEvents
                
                UpdateRegistry
                ResetProgressBar
                
                ' Unlock controls
                With frmMain
                    .cmdChoice(2).Enabled = True
                    .cmdChoice(3).Enabled = True
                    .fraChoice.Enabled = True
                    .fraEncrypt(1).Enabled = True
                    .picDataType.Enabled = True
                    .txtPwd.Enabled = True
                End With
                
                cmdChoice_GotFocus 0
                LockWindowUpdate 0&    ' unlock textbox after processing
                DoEvents

           Case 2  ' Show About form
                frmAbout.DisplayAbout

           Case Else  ' EXIT button
                DoEvents
                Screen.MousePointer = vbDefault
                gblnStopProcessing = True
                SetDllProcessingFlag
                DoEvents
                
                UpdateRegistry
                ResetProgressBar
                TerminateProgram
    End Select

CleanUp:
    DoEvents
    Screen.MousePointer = vbDefault   ' Reset mouse pointer to normal
    DoEvents

End Sub

Private Sub cmdChoice_GotFocus(Index As Integer)

    Select Case Index
           Case 0
                cmdChoice(0).Enabled = True
                cmdChoice(0).Visible = True
                cmdChoice(1).Enabled = False
                cmdChoice(1).Visible = False
           Case 1
                cmdChoice(0).Enabled = False
                cmdChoice(0).Visible = False
                cmdChoice(1).Enabled = True
                cmdChoice(1).Visible = True
    End Select

    Refresh

End Sub

Private Sub SetDllProcessingFlag()

    ' Called by cmdChoice_Click()
    
    ' If the particular object is active then
    ' set the property value
    If Not mobjCipher Is Nothing Then
        mobjCipher.StopProcessing = gblnStopProcessing
    End If
    
    If Not mobjCRC32 Is Nothing Then
        mobjCRC32.StopProcessing = gblnStopProcessing
    End If
                    
    If Not mobjHash Is Nothing Then
        mobjHash.StopProcessing = gblnStopProcessing
    End If
    
    If Not mobjPrng Is Nothing Then
        mobjPrng.StopProcessing = gblnStopProcessing
    End If
                
    DoEvents
    
End Sub

Private Sub Form_Load()

    gblnStopProcessing = False   ' Set global flag
    mstrFolder = vbNullString

    ' Instantiate class and DLL objects
    Set mobjKeyEdit = New cKeyEdit
    Set mobjCipher = New kiCrypt.cCipher
    Set mobjCRC32 = New kiCrypt.cCRC32
    Set mobjHash = New kiCrypt.cHash
    Set mobjPrng = New kiCrypt.cPrng
    
    GetRegistryData
    DisableX frmMain   ' Disable "X" in upper right corner of form
    LoadComboBox
    
    mblnCipher = True
    mblnCRC32 = False
    mblnHash = False
    mblnPrng = False
    mblnStringData = True
           
    With mobjCipher
        mlngPwdLength_Min = .PasswordLength_Min
        mlngPwdLength_Max = .PasswordLength_Max
    End With
    
    With frmMain
        .Caption = gstrVersion

        ' If operating system is Windows 8 or 8.1
        ' form caption is centered automatically
        If Not gblnWin8or81 Then
            CenterCaption frmMain   ' Manually center form caption
        End If
    
        .lblFormTitle.Caption = PGM_NAME
        .lblDisclaimer.Caption = "This software is provided without any " & _
                                 "warrantees or guarantees implied or intended."
        .lblOperSystem.Caption = gstrOperSystem
        .lblEncryptMsg.Visible = False
        
        .txtRandom.BackColor = &HE0E0E0   ' Light gray
        
        ' Passwords or phrases are case sensitive and
        ' have a potential length of fifty characters
        '
        '  123456789+123456789+123456789+123456789+123456789+
        ' "Soylent Green IS People"     ex:  23 character password
        .txtPwd.PasswordChar = "*"      ' Display only asteriks
        .txtPwd.Text = vbNullString     ' Empty password textbox
        .chkShowPwd.Caption = Space$(5) & "Password length in asteriks"
        .chkShowPwd.Value = vbChecked   ' True-show number of actual asteriks
        
        ' set the command buttons
        .cmdChoice(0).Enabled = True
        .cmdChoice(0).Visible = True
        .cmdChoice(1).Enabled = False
        .cmdChoice(1).Visible = False
        
        .cmdCopy.Enabled = False
        .cmdCopy.Visible = False
        
        optChoice_Click 0
        cboEncrypt_Click
        cboKeyMix_Click
        ResetProgressBar
        
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless   ' reduce flicker
        .Refresh
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    gblnStopProcessing = True

    ' If object is still active then
    ' send a command to stop it
    If Not mobjCipher Is Nothing Then
        mobjCipher.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjCRC32 Is Nothing Then
        mobjCRC32.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjHash Is Nothing Then
        mobjHash.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    If Not mobjPrng Is Nothing Then
        mobjPrng.StopProcessing = gblnStopProcessing
        DoEvents
    End If
    
    Erase mabytPwd()            ' Erase password array
    Set mobjKeyEdit = Nothing   ' Free class objects from memory
    Set mobjCipher = Nothing
    Set mobjCRC32 = Nothing
    Set mobjHash = Nothing
    Set mobjPrng = Nothing
    Screen.MousePointer = vbDefault

    If UnloadMode = 0 Then
        TerminateProgram
    End If
    
End Sub

' 29-Jan-2010 Add events to track cipher progress
Private Sub mobjCipher_CipherProgress(ByVal lngProgress As Long)
    
    If mblnStringData Then
        Exit Sub
    End If
    
    ProgressBar picProgressBar, lngProgress, vbBlue
    DoEvents
    
End Sub

' 29-Jan-2010 Add events to track CRC32 progress
Private Sub mobjCRC32_CRCProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress, vbRed
    DoEvents
    
End Sub

' 29-Jan-2010 Add events to track hash progress
Private Sub mobjHash_HashProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress, vbBlack
    DoEvents
    
End Sub

Private Sub optChoice_Click(Index As Integer)

    Dim intIndex As Integer
    
    mblnCipher = False
    mblnCRC32 = False
    mblnHash = False
    mblnPrng = False
    
    Select Case Index
           Case 0  ' encryption
                mblnCipher = True
                optDataType_Click 0   ' String data

                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = True
                    .fraEncrypt(1).Visible = True
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = False
                    
                    ' Prepare combo boxes
                    .lblAlgo(5).Visible = False
                    .cboRounds.Visible = False
                    .cboEncrypt.Enabled = True
                    .cboEncrypt.ForeColor = vbBlack
                    
                    If mlngCipherAlgo > 3 Then
                        .cboKeyLength.Enabled = True
                        .cboKeyLength.ForeColor = vbBlack
                    Else
                        .cboKeyLength.Enabled = False
                        .cboKeyLength.ForeColor = vbGrayText
                    End If
                    
                    .lblAlgo(2).Caption = "Password Hash Algorithm"
                    With .lblPwd
                        .Visible = True
                        .Caption = vbNullString
                        .Caption = "Enter " & CStr(mlngPwdLength_Min) & " - " & _
                                   CStr(mlngPwdLength_Max) & _
                                   " character password or phrase  [ Case sensitive ]"
                    End With
                    .lblEncrypt(1).Caption = "Output file name and location"
                    .lblEncrypt(1).Visible = True
                    .lblEncryptMsg.Caption = "After encryption, data sizes will not match original sizes.  " & _
                                             "This is due to internal padding and information " & _
                                             "required for later decryption."
                    .lblEncryptMsg.Visible = True
                    .txtOutput.Text = vbNullString
                    .txtPwd.Visible = True
                    .txtPwd.Locked = False
                    .chkExtraInfo.Enabled = True
                    .chkExtraInfo.Visible = True
                    .chkShowPwd.Enabled = True
                    .chkShowPwd.Visible = True
                End With
                 
                intIndex = IIf(mblnStringData, 0, 1)
                optDataType_Click intIndex
                chkShowPwd_Click
                
           Case 1 ' CRC-32
                mblnCRC32 = True
                optDataType_Click 0   ' String data

                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = False
                    .fraEncrypt(1).Visible = False
                    .lblEncrypt(1).Caption = "Calculated CRC-32 value in hex"
                    .lblEncrypt(1).Visible = True
                    .lblPwd.Visible = False
                    .lblPwd.Caption = vbNullString
                    .lblEncryptMsg.Visible = False
                    .txtPwd.Visible = False
                    .txtOutput.Text = vbNullString
                    .txtPwd.Text = vbNullString
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = True
                    .chkExtraInfo.Enabled = True
                    .chkExtraInfo.Visible = True
                    .chkExtraInfo.Caption = "Return leading zeros"
                    .chkShowPwd.Enabled = False
                    .chkShowPwd.Visible = False
                End With

                intIndex = IIf(mblnStringData, 0, 1)
                optDataType_Click intIndex
           
           Case 2  ' Hash
                mblnHash = True
                optDataType_Click 0   ' String data
                
                With frmMain
                    .fraMain(0).Visible = False
                    .fraMain(0).Enabled = False
                    .fraMain(1).Visible = True
                    .fraMain(1).Enabled = True
                    .fraEncrypt(0).Enabled = True
                    .fraEncrypt(0).Visible = True
                    .fraEncrypt(1).Enabled = True
                    .fraEncrypt(1).Visible = True
                    
                    ' Prepare combo boxes
                    .lblAlgo(5).Visible = True
                    .cboRounds.Visible = True
                    .cboEncrypt.Enabled = False
                    .cboEncrypt.ForeColor = vbGrayText
                    .cboKeyLength.Enabled = False
                    .cboKeyLength.ForeColor = vbGrayText
                    
                    .lblEncrypt(1).Caption = "Hashed results"
                    .lblEncrypt(1).Visible = True
                    .lblAlgo(2).Caption = "Hash Algorithm"
                    .lblPwd.Caption = vbNullString
                    .lblPwd.Visible = False
                    .lblEncryptMsg.Caption = vbNewLine & "Be patient." & vbNewLine & _
                                             "Large amounts of data take longer to process."
                    .lblEncryptMsg.Visible = True
                    .txtOutput.Text = vbNullString
                    .txtPwd.Visible = False
                    .txtPwd.Text = vbNullString
                    .chkExtraInfo.Enabled = True
                    .chkExtraInfo.Visible = True
                    .chkShowPwd.Enabled = False
                    .chkShowPwd.Visible = False
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = True
                End With
                
                intIndex = IIf(mblnStringData, 0, 1)
                optDataType_Click intIndex
                cboHash_Click
                cboRounds.ListIndex = 0
                
           Case 3  ' Random data
                mblnPrng = True
                With frmMain
                    .fraMain(0).Visible = True
                    .fraMain(0).Enabled = True
                    .fraMain(1).Visible = False
                    .fraMain(1).Enabled = False
                    .lblEncryptMsg.Visible = False
                    .txtRandom.Text = vbNullString
                    .chkExtraInfo.Enabled = False
                    .chkExtraInfo.Visible = False
                    .chkShowPwd.Enabled = False
                    .chkShowPwd.Visible = False
                    .cmdCopy.Enabled = False
                    .cmdCopy.Visible = False
                    .cboRandom.ListIndex = 0
                End With
                
                cboRandom_Click
    End Select

End Sub

Private Sub optDataType_Click(Index As Integer)

    With frmMain
        
        Select Case Index
    
               Case 0   ' string data
                    mblnStringData = True
                    
                    If mblnCipher Then
                        .lblEncrypt(0).Caption = "Data to be encrypted or decrypted"
                    ElseIf mblnCRC32 Then
                        .lblEncrypt(0).Caption = "Data to be calculated for CRC"
                    ElseIf mblnHash Then
                        .lblEncrypt(0).Caption = "Data to be hashed"
                    End If
                    
                    .optDataType(0).Value = True
                    .optDataType(1).Value = False
                    .cmdBrowse.Enabled = False
                    .cmdBrowse.Visible = False
                    .txtInputFile.Text = vbNullString
                    .txtInputFile.Visible = False
                    .txtOutput.Text = vbNullString
                    
                    If mblnCRC32 Or mblnHash Then
                        With .txtInputString
                            .Height = 3255
                            .Text = vbNullString
                            .Visible = True
                        End With
                        .txtOutput.Visible = True
                        
                    ElseIf mblnCipher Then
                        With .txtInputString
                            .Height = 4400
                            .Text = vbNullString
                            .Visible = True
                        End With
                        .txtOutput.Visible = False
                    End If
                                    
               Case 1   ' file data
                    mblnStringData = False
                    
                    If mblnCipher Then
                        .lblEncrypt(0).Caption = "Data to be encrypted or decrypted"
                    ElseIf mblnCRC32 Then
                        .lblEncrypt(0).Caption = "Data to be calculated for CRC"
                    ElseIf mblnHash Then
                        .lblEncrypt(0).Caption = "Data to be hashed"
                    End If
                    
                    .optDataType(0).Value = False
                    .optDataType(1).Value = True
                    .cmdBrowse.Enabled = True
                    .cmdBrowse.Visible = True
                    .lblEncrypt(1).Visible = True
                    .txtInputString.Text = vbNullString
                    .txtInputString.Visible = False
                    .txtInputFile.Text = vbNullString
                    .txtInputFile.Visible = True
                    .txtOutput.Text = vbNullString
                    .txtOutput.Visible = True
        End Select
        
        If mblnStringData Then
            .picProgressBar.Visible = False
        Else
            .picProgressBar.Visible = True
        End If

        If mblnHash Then
            .chkExtraInfo.Caption = "Return as lowercase"
        
            If mblnRetLowercase Then
                .chkExtraInfo.Value = vbChecked
            Else
                .chkExtraInfo.Value = vbUnchecked
            End If
            
        ElseIf mblnCipher Then
        
            ' String processing
            If mblnStringData Then
                
                .chkExtraInfo.Caption = "Return as lowercase"
                
                If mblnRetLowercase Then
                    .chkExtraInfo.Value = vbChecked
                Else
                    .chkExtraInfo.Value = vbUnchecked
                End If
            Else
                ' File processing
                .chkExtraInfo.Caption = "Create new destination file"
                
                If mblnCreateNewFile Then
                    .chkExtraInfo.Value = vbChecked
                Else
                    .chkExtraInfo.Value = vbUnchecked
                End If
            End If
        End If
    End With
    
    If mblnCRC32 Or mblnHash Then
        cmdCopy.Enabled = False
    End If
    
End Sub

' ***************************************************************************
' Data functions
' ***************************************************************************
Private Sub LoadComboBox()

    Dim lngIdx As Long
    
    mblnLoading = True

    ' Encryption algorithms
    With frmMain
        With .cboEncrypt
            .Clear
            .AddItem "ArcFour"
            .AddItem "Base64"
            .AddItem "Blowfish"
            .AddItem "Gost"
            .AddItem "Rijndael (AES)"
            .AddItem "Serpent"
            .AddItem "Skipjack"
            .AddItem "Twofish"
            .ListIndex = 0
        End With
    
        ' Hash algorithms
        With .cboHash
            .Clear
            .AddItem "MD2"             ' 0
            .AddItem "MD4"             ' 1
            .AddItem "MD5"             ' 2
            .AddItem "SHA-1"           ' 3
            .AddItem "SHA-256"         ' 4
            .AddItem "SHA-384"         ' 5
            .AddItem "SHA-512"         ' 6
            .AddItem "Tiger3-128"      ' 7
            .AddItem "Tiger3-160"      ' 8
            .AddItem "Tiger3-192"      ' 9
            .AddItem "Tiger3-224"      ' 10
            .AddItem "Tiger3-256"      ' 11
            .AddItem "Tiger3-384"      ' 12
            .AddItem "Tiger3-512"      ' 13
            .AddItem "Whirlpool-224"   ' 14
            .AddItem "Whirlpool-256"   ' 15
            .AddItem "Whirlpool-384"   ' 16
            .AddItem "Whirlpool-512"   ' 17
            .ListIndex = 4             ' Default = SHA-256
        End With
    
        ' Encryption algorithms
        With .cboRandom
            .Clear
            .AddItem "Keyboard Chars (1400 bytes)"
            .AddItem "Hex String (800 bytes)"
            .AddItem "Hex Array (440 elements)"
            .AddItem "Byte Array (360 elements)"
            .AddItem "Long Array (120 digits)"
            .AddItem "Double Array (80 digits)"
            .ListIndex = 0
        End With
    
        With .cboKeyLength
            .Clear
            For lngIdx = 128 To 256 Step 32
                .AddItem CStr(lngIdx) & " bits"
            Next lngIdx
            .ListIndex = 0
            .Enabled = True
        End With
            
        With .cboBlockLength
            .Clear
            For lngIdx = 128 To 256 Step 32
                .AddItem CStr(lngIdx) & " bits"
            Next lngIdx
            .ListIndex = 0
            .Enabled = False
        End With
            
        ' Number of rounds for hashing
        With .cboRounds
            .Clear
            For lngIdx = 1 To 10
                .AddItem CStr(lngIdx)
            Next lngIdx
            .ListIndex = 0   ' Default = 1 round
        End With
    
        With .cboKeyMix
            .Clear
            For lngIdx = 1 To 5
                .AddItem CStr(lngIdx)
            Next lngIdx
            .ListIndex = 0
        End With
    End With
    
    mblnLoading = False
    
End Sub

' ***************************************************************************
' Routine:       Cipher_Processing
'
' Description:   Encryption demonstration
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated routine to match new screen
' 01-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected error handling and flow of data
' 03-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified structure of code for easier reading
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added reference to new property CreateNewFile().  Change input
'              to a variable if you want to overwrite file to be encrypted or
'              decrypted
' 17-Sep-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added more security for password.
' ***************************************************************************
Private Sub Cipher_Processing()

    Dim strMsg        As String
    Dim strData       As String
    Dim strTemp       As String
    Dim strOutputFile As String
    Dim abytData()    As Byte
    Dim astrMsgBox()  As String
    Dim hFile         As Long
    Dim lngLength     As Long
    Dim lngEncrypt    As enumCIPHER_ACTION
    
    Const ENCRYPT_EXT As String = ".ENC"
    Const DECRYPT_EXT As String = ".DEC"
    
    Erase abytData()    ' Always start with empty arrays
    Erase astrMsgBox()
    strTemp = vbNullString
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(TrimStr(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process", , , 3
            txtInputString.SetFocus
            GoTo Cipher_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(TrimStr(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing", , , 3
            GoTo Cipher_Processing_CleanUp
        End If
    End If
    
    ' Verify a password is being used
    If Not CBool(IsArrayInitialized(mabytPwd)) Then
        InfoMsg "Need a password to perform Encryption/Decryption.", , , 3
        GoTo Cipher_Processing_CleanUp
    Else
        ' Has password array been initialized
        If CBool(IsArrayInitialized(mabytPwd)) Then
            ' Does password array hold any data
            If UBound(mabytPwd) < 0 Then
                InfoMsg "Need a password to perform Encryption/Decryption.", , , 3
                GoTo Cipher_Processing_CleanUp
            End If
        End If
    End If
    
    strTemp = ByteArrayToString(mabytPwd())   ' Load temp variable
    lngLength = Len(strTemp)                  ' Capture actual length of password
    
' ******************************************
'   Verify password
'   Uncomment Debug line and then
'   press CTRL+G to open immediate window
' ******************************************
'Debug.Print strTemp
' ******************************************
    strTemp = vbNullString    ' Empty temp variable
    
    ' Evaluate password length
    Select Case lngLength
           
           Case 0
                ' Safety net in case something
                ' slips thru the testing above
                InfoMsg "Password is missing." & vbNewLine & vbNewLine & _
                        "Minimum length:   " & CStr(mlngPwdLength_Min) & " characters" & vbNewLine & _
                        "Maximum length:   " & CStr(mlngPwdLength_Max) & " characters", , , 4
                GoTo Cipher_Processing_CleanUp
    
           Case Is < mlngPwdLength_Min
                InfoMsg "Password is too short." & vbNewLine & _
                        "Minimum length:   " & CStr(mlngPwdLength_Min) & " characters", , , 3
                GoTo Cipher_Processing_CleanUp
    
           Case Is > mlngPwdLength_Max
                InfoMsg "Password is too long." & vbNewLine & _
                        "Maximum length:   " & CStr(mlngPwdLength_Max) & " characters", , , 3
                GoTo Cipher_Processing_CleanUp
    End Select
    
    lngLength = 0   ' Reset variable
        
    ' Disable form STOP button
    ' until a choice is made
    cmdChoice(1).Enabled = False
    
    '----------------------------------------------------------
    ' Prepare message box display.
    '
    ' These are the button captions,
    ' in order, from left to right.
    ReDim astrMsgBox(3)
    astrMsgBox(0) = "Encrypt"
    astrMsgBox(1) = "Decrypt"
    astrMsgBox(2) = "Cancel"
    
    ' Prompt user with message box
    Select Case MessageBoxH(Me.hWnd, GetDesktopWindow(), _
                            "What do you want to do?  ", _
                            PGM_NAME, astrMsgBox(), eMSG_ICONQUESTION)
           
           ' These are valid responses
           Case IDYES:    lngEncrypt = eMSG_ENCRYPT
           Case IDNO:     lngEncrypt = eMSG_DECRYPT
           Case IDCANCEL: GoTo Cipher_Processing_CleanUp
    End Select
    '----------------------------------------------------------
    
    ' Enable form STOP button
    ' after a choice is made
    cmdChoice(1).Enabled = True
    
    ' *********************************************************
    ' Encrypt/Decrypt - String
    ' *********************************************************
    If mblnStringData Then
                    
        Screen.MousePointer = vbHourglass
        
        With mobjCipher
            .HashMethod = mlngHashAlgo        ' Type of hash algorithm selected
            .CipherMethod = mlngCipherAlgo    ' Type of cipher algorithm selected
            .KeyLength = mlngKeyLength        ' Ignored by Base64
            .PrimaryKeyRounds = mlngKeyMix    ' Only by Blowfish, GOST, Serpent
            .BlockSize = mlngBlockLength      ' Used by Rijndael only
            .CipherRounds = mlngRounds        ' Ignored by Base64
                        
            Select Case lngEncrypt
                   Case eMSG_ENCRYPT    ' Encrypt string
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mabytPwd()                 ' Process password
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        strData = TrimStr(txtInputString.Text)    ' Remove any hidden characters
                        abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                        
                        ' See if using Base64 cipher
                        If mlngCipherAlgo = eCIPHER_BASE64 Then
                            ' Encrypt data string
                            If .EncryptString(abytData()) Then
                                strData = ByteArrayToString(abytData())   ' Convert byte array to string data
                            Else
                                gblnStopProcessing = .StopProcessing      ' See if processing aborted
                            End If
                        Else
                            ' Encrypt data string
                            If .EncryptString(abytData()) Then
                            
                                strData = ByteArrayToHex(abytData())    ' convert single charaters to hex
                            
                                ' Verify that this is hex data
                                If Not IsHexData(strData) Then
                                    InfoMsg "Failed to convert encrypted data to hex.", , , 3
                                    gblnStopProcessing = True
                                End If
                            Else
                                gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            End If   ' .EncryptString
                        End If   ' mlngCipherAlgo
                
                        strData = TrimStr(strData)  ' Store string data in text box
                
                        If mblnRetLowercase Then
                            strData = LCase$(strData)   ' Convert string to lowercase
                        Else
                            strData = UCase$(strData)   ' Convert string to uppercase
                        End If
                                    
                        txtInputString.Text = vbNullString   ' Empty text box
                        txtInputString.Text = strData        ' Fill text box
                   
                   Case eMSG_DECRYPT    ' Decrypt string
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mabytPwd()                 ' Process password
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        strData = TrimStr(txtInputString.Text)   ' Remove leading and trailing blanks
                        
                        ' See if using Base64 cipher
                        If mlngCipherAlgo = eCIPHER_BASE64 Then
                        
                            abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                            .DecryptString abytData()                 ' Decrypt data string
                            gblnStopProcessing = .StopProcessing      ' See if processing aborted
                        
                        Else
                            
                            ' Verify that this is hex data
                            If IsHexData(strData) Then
                                abytData() = HexToByteArray(strData)  ' Convert hex to single char
                                .DecryptString abytData()             ' Decrypt data string
                                gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            Else
                                strMsg = "This text is not hex data or" & vbNewLine & _
                                         "length is not divisible by two."
                                strMsg = strMsg & vbNewLine & vbNewLine & "Cannot decrypt."
                                InfoMsg strMsg, , , 3
                                GoTo Cipher_Processing_CleanUp
                            End If
                        
                        End If   ' mlngCipherAlgo
            
                        strData = ByteArrayToString(abytData())   ' Convert byte array to string data
                        txtInputString.Text = vbNullString        ' Empty text box
                        txtInputString.Text = TrimStr(strData)    ' Store string data in text box
            End Select
        End With
            
        DoEvents
        If gblnStopProcessing Then
            GoTo Cipher_Processing_CleanUp
        End If
    
    Else
        ' *********************************************************
        ' Encrypt/Decrypt - File
        ' *********************************************************
        If IsPathValid(mstrFilename) Then
        
            ' If creating new file then
            ' change to appropriate extension
            If mblnCreateNewFile Then
                Select Case lngEncrypt
                       
                       Case eMSG_ENCRYPT
                            ' First verify file name has an extension
                            If InStr(mstrFilename, ".") > 0 Then
                                strOutputFile = Mid$(mstrFilename, 1, InStrRev(mstrFilename, ".") - 1) & ENCRYPT_EXT
                            Else
                                ' No file extension
                                strOutputFile = mstrFilename & ENCRYPT_EXT
                            End If
                       
                       Case eMSG_DECRYPT
                            ' First verify file name has an extension
                            If InStr(mstrFilename, ".") > 0 Then
                                strOutputFile = Mid$(mstrFilename, 1, InStrRev(mstrFilename, ".") - 1) & DECRYPT_EXT
                            Else
                                ' No file extension
                                strOutputFile = mstrFilename & DECRYPT_EXT
                            End If
                End Select
            End If
            
        Else
            InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename, , , 3
            txtInputFile.SetFocus
            GoTo Cipher_Processing_CleanUp
        End If
    
        ' If output file exist
        ' then verify it is empty
        DoEvents
        If IsPathValid(strOutputFile) Then
            
            SetFileAttributes strOutputFile, FILE_ATTRIBUTE_NORMAL
            hFile = FreeFile
            Open strOutputFile For Output As #hFile
            Close #hFile
            DoEvents
            
        End If   ' IsPathValid
        
        With mobjCipher
            .HashMethod = mlngHashAlgo          ' Type of hash algorithm selected
            .CipherMethod = mlngCipherAlgo      ' Type of cipher algorithm selected
            .KeyLength = mlngKeyLength          ' Ignored by Base64
            .PrimaryKeyRounds = mlngKeyMix      ' Only by Blowfish, GOST
            .BlockSize = mlngBlockLength        ' Used by Rijndael only
            .CipherRounds = mlngRounds          ' Ignored by Base64, Rijndael
            .CreateNewFile = mblnCreateNewFile  ' True - Create new output file
                                                ' False - Overwrite input file
            Select Case lngEncrypt
                   Case eMSG_ENCRYPT    ' Encrypt file
                        ' Verify overwrite message
                        If Not mblnCreateNewFile Then
                            
                            If ResponseMsg("Are you sure you want to overwrite input file?", _
                                           vbYesNo, "Verify output target") = vbNo Then
                                           
                                txtInputFile.SetFocus
                                GoTo Cipher_Processing_CleanUp
                            End If
                            
                            strOutputFile = mstrFilename
                        End If   ' mblnCreateNewFile
                        
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mabytPwd()                 ' Process password
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        If .EncryptFile(mstrFilename) Then
                        
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            txtOutput.Text = vbNullString
                            Screen.MousePointer = vbDefault
                                                
                            If Not gblnStopProcessing Then
                                txtOutput.Text = strOutputFile
                                strMsg = "Finished encrypting file." & vbNewLine & vbNewLine
                                strMsg = strMsg & strOutputFile & vbNewLine & vbNewLine
                                InfoMsg strMsg, , , 4
                            End If
                        Else
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If   '.EncryptFile
                
                   Case eMSG_DECRYPT    ' Decrypt file
                        .HashMethod = mlngHashAlgo          ' Type of hash algorithm selected
                        .CipherMethod = mlngCipherAlgo      ' Type of cipher algorithm selected
                        .KeyLength = mlngKeyLength          ' Ignored by Base64
                        .PrimaryKeyRounds = mlngKeyMix      ' Only by Blowfish, GOST
                        .BlockSize = mlngBlockLength        ' Used by Rijndael only
                        .CipherRounds = mlngRounds          ' Ignored by Base64, Rijndael
                        .CreateNewFile = mblnCreateNewFile  ' True - Create new output file
                        
                        ' Verify overwrite message
                        If Not mblnCreateNewFile Then
                            
                            If ResponseMsg("Are you sure you want to overwrite input file?", _
                                           vbYesNo, "Verify output target") = vbNo Then
                                           
                                txtInputFile.SetFocus
                                GoTo Cipher_Processing_CleanUp
                            End If
                            
                            strOutputFile = mstrFilename
                        End If   ' mblnCreateNewFile
                        
                        If mlngCipherAlgo <> eCIPHER_BASE64 Then
                            .Password = mabytPwd()                 ' Process password
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
                        End If
                    
                        DoEvents
                        If gblnStopProcessing Then
                            GoTo Cipher_Processing_CleanUp
                        End If
            
                        If .DecryptFile(mstrFilename) Then
                            
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                            txtOutput.Text = vbNullString
                            Screen.MousePointer = vbDefault
                                                
                            If Not gblnStopProcessing Then
                                txtOutput.Text = strOutputFile
                                strMsg = "Finished decrypting file." & vbNewLine & vbNewLine
                                strMsg = strMsg & strOutputFile & vbNewLine & vbNewLine
                                InfoMsg strMsg, , , 4
                            End If
                            
                        Else
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If   ' .DecryptFile
            End Select
        End With
    End If   ' If mblnStringData
    
Cipher_Processing_CleanUp:
    Erase abytData()                  ' Always empty arrays when not needed
    Erase astrMsgBox()
    strTemp = vbNullString            ' Empty temp variable
    Screen.MousePointer = vbDefault   ' Set mouse pointer back to normal
    
End Sub

Private Sub CRC32_Processing()
                    
    Dim blnRetLeadZeros As Boolean
    Dim strHex          As String
    Dim abytData()      As Byte
    
    Screen.MousePointer = vbHourglass
    
    Erase abytData()   ' Always start eith an empty array
    cmdCopy.Enabled = False
    blnRetLeadZeros = IIf(CBool(chkExtraInfo.Value), True, False)
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(TrimStr(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process", , , 3
            GoTo CRC32_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(TrimStr(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing", , , 3
            GoTo CRC32_Processing_CleanUp
        End If
    End If
        
    ' CRC32 string data
    If mblnStringData Then
        With mobjCRC32
            abytData() = StringToByteArray(txtInputString.Text)   ' Convert string data to byte array
            strHex = .CRC32_String(abytData(), blnRetLeadZeros)   ' Calculate CRC
            gblnStopProcessing = .StopProcessing                  ' See if processing aborted
        End With
    Else
        ' CRC32 file data
        If IsPathValid(mstrFilename) Then
            With mobjCRC32
                abytData() = StringToByteArray(mstrFilename)        ' Convert string data to byte array
                strHex = .CRC32_File(abytData(), blnRetLeadZeros)   ' Calculate CRC
                gblnStopProcessing = .StopProcessing                ' See if processing aborted
            End With
        Else
            InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename, , , 3
            txtInputFile.SetFocus
            GoTo CRC32_Processing_CleanUp
        End If
    End If
        
    DoEvents
    If gblnStopProcessing Then
        GoTo CRC32_Processing_CleanUp
    End If
                
    txtOutput.Text = TrimStr(strHex)
    cmdCopy.Enabled = True
    
CRC32_Processing_CleanUp:
    Erase abytData()   ' Always empty arrays when not needed
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Hash_Processing()
                    
    Dim strOutput  As String
    Dim abytData() As Byte
    Dim abytHash() As Byte
    
    Screen.MousePointer = vbHourglass
    
    Erase abytData()    ' Always start with empty arrays
    Erase abytHash()
    
    strOutput = vbNullString
    cmdCopy.Enabled = False
    
    If mblnStringData Then
        ' Test for string data to process
        If Len(TrimStr(txtInputString.Text)) = 0 Then
            InfoMsg "Need some data to process", , , 3
            GoTo Hash_Processing_CleanUp
        End If
    Else
        ' Test for file name to process
        If Len(TrimStr(mstrFilename)) = 0 Then
            InfoMsg "Path\File name missing", , , 3
            GoTo Hash_Processing_CleanUp
        End If
    End If
    
    With mobjHash
        .StopProcessing = False                ' Reset stop flag
        .HashMethod = mlngHashAlgo             ' Hash algorithm selected
        .HashRounds = mlngRounds               ' Number of passes
        .ReturnLowercase = mblnRetLowercase   ' TRUE = Return as lowercase
                                               ' FALSE = Return as uppercase
        ' Hash string data
        If optDataType(0).Value Then
            
            abytData() = StringToByteArray(txtInputString.Text)    ' Convert to byte array
            abytHash() = .HashString(abytData())                   ' Hash string data
            gblnStopProcessing = .StopProcessing                   ' See if processing aborted
            strOutput = ByteArrayToString(abytHash())              ' Convert byte array to string
            
        Else
            ' Hash a file
            If IsPathValid(mstrFilename) Then
                abytData() = StringToByteArray(mstrFilename)   ' Convert to byte array
                abytHash() = .HashFile(abytData())             ' Hash file
                gblnStopProcessing = .StopProcessing           ' See if processing aborted
                strOutput = ByteArrayToString(abytHash())      ' Convert byte array to string
            Else
                InfoMsg "Cannot locate Path\File." & vbNewLine & mstrFilename, , , 3
                GoTo Hash_Processing_CleanUp
            End If
        End If
    
    End With
    
    DoEvents
    If gblnStopProcessing Then
        GoTo Hash_Processing_CleanUp
    End If
    
    txtOutput.Text = TrimStr(strOutput)
    cmdCopy.Enabled = True
    
Hash_Processing_CleanUp:
    Erase abytData()    ' Always empty arrays when not needed
    Erase abytHash()
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub RndData_Processing(ByVal lngChoice As Long)

    Dim lngIndex  As Long
    Dim lngCount  As Long
    Dim strOutput As String
    Dim avntData  As Variant

    Screen.MousePointer = vbHourglass
    avntData = Empty
    lngCount = 0
    strOutput = vbNullString
    txtRandom.Text = vbNullString

    With mobjPrng
        Select Case lngChoice
    
               Case 0   ' ASCII String
                    txtRandom.Text = .BuildWithinRange(1400, 32, 126, ePRNG_ASCII)
    
               Case 1   ' Hex String
                    txtRandom.Text = .BuildRndData(800, ePRNG_HEX)
    
               Case 2   ' Hex Array
                    avntData = .BuildRndData(440, ePRNG_HEX_ARRAY)
    
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo RndData_Processing_CleanUp
                    End If
    
                    For lngIndex = 0 To UBound(avntData) - 1
                        
                        If lngCount = 0 Then
                            strOutput = strOutput & Space$(1) & avntData(lngIndex)
                        Else
                            strOutput = strOutput & Space$(2) & avntData(lngIndex)
                        End If
                        
                        lngCount = lngCount + 1
    
                        If lngCount = 22 Then
                            lngCount = 0
                            strOutput = strOutput & vbNewLine
                        End If
    
                        DoEvents
                        If gblnStopProcessing Then
                            strOutput = vbNullString
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    txtRandom.Text = strOutput
    
               Case 3   ' Byte Array
                    avntData = .BuildRndData(360, ePRNG_BYTE_ARRAY)
    
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo RndData_Processing_CleanUp
                    End If
    
                    For lngIndex = 0 To UBound(avntData) - 1
                        
                        If lngCount = 0 Then
                            strOutput = strOutput & Format$(avntData(lngIndex), "@@@@")
                        Else
                            strOutput = strOutput & Format$(avntData(lngIndex), "@@@@@")
                        End If
                        
                        lngCount = lngCount + 1
    
                        If lngCount = 18 Then
                            lngCount = 0
                            strOutput = strOutput & vbNewLine
                        End If
    
                        DoEvents
                        If gblnStopProcessing Then
                            strOutput = vbNullString
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    txtRandom.Text = strOutput
    
               Case 4   ' Long Array
                    avntData = .BuildRndData(120, ePRNG_LONG_ARRAY)
    
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo RndData_Processing_CleanUp
                    End If
    
                    For lngIndex = 0 To UBound(avntData) - 1
                        
                        If lngCount = 0 Then
                            strOutput = strOutput & Format$(avntData(lngIndex), String(16, "@"))
                        Else
                            strOutput = strOutput & Format$(avntData(lngIndex), String(13, "@"))
                        End If
                        
                        lngCount = lngCount + 1
    
                        If lngCount = 6 Then
                            lngCount = 0
                            strOutput = strOutput & vbNewLine
                        End If
    
                        DoEvents
                        If gblnStopProcessing Then
                            strOutput = vbNullString
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    txtRandom.Text = strOutput
    
               Case 5   ' Double Array (0 to 1)
                    avntData = .BuildRndData(80, ePRNG_DBL_ARRAY)
    
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo RndData_Processing_CleanUp
                    End If
    
                    For lngIndex = 0 To UBound(avntData) - 1
                        
                        If lngCount = 0 Then
                            strOutput = strOutput & Space$(1) & Format$(avntData(lngIndex), String(21, "@"))
                        Else
                            strOutput = strOutput & Space$(3) & Format$(avntData(lngIndex), String(17, "@"))
                        End If
                        lngCount = lngCount + 1
    
                        If lngCount = 4 Then
                            strOutput = strOutput & vbNewLine
                            lngCount = 0
                        End If
    
                        DoEvents
                        If gblnStopProcessing Then
                            strOutput = vbNullString
                            Exit For    ' exit For..Next loop
                        End If
    
                    Next lngIndex
    
                    txtRandom.Text = strOutput
        End Select
    End With
    
RndData_Processing_CleanUp:
    DoEvents
    avntData = Empty
    Screen.MousePointer = vbDefault

End Sub

Private Sub ResetProgressBar()

    ' Resets progressbar to zero
    ' with all white background
    ProgressBar picProgressBar, 0, vbWhite
    
End Sub

' ***************************************************************************
' Routine:       ProgessBar
'
' Description:   Fill a picturebox as if it were a horizontal progress bar.
'
' Parameters:    objProgBar - name of picture box control
'                lngPercent - Current percentage value
'                lngForeColor - Optional-The progression color. Default = Black.
'                           can use standard VB colors or long Integer
'                           values representing a color.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 14-FEB-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for the speed enhancement.
'              This way the progress bar is only initialized once.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Private Sub ProgressBar(ByRef objProgBar As PictureBox, _
                        ByVal lngPercent As Long, _
               Optional ByVal lngForeColor As Long = vbBlue)

    Dim strPercent As String
    
    Const MAX_PERCENT As Long = 100
    
    ' Called by ResetProgressBar() routine
    ' to reinitialize progress bar properties.
    ' If progression color matches background
    ' color then progressbar is being reset
    ' to the starting position.
    If lngForeColor = vbWhite Then
        
        With objProgBar
            .AutoRedraw = True      ' Required to prevent flicker
            .BackColor = &HFFFFFF   ' White
            .DrawMode = 10          ' Not Xor Pen
            .FillStyle = 0          ' Solid fill
            .FontName = "Segoe UI"  ' Name of font (same as MS msgbox)
            .FontSize = 11          ' Font point size
            .FontBold = True        ' Font is bold.  Easier to see.
            Exit Sub                ' Exit this routine
        End With
    
    End If
        
    ' If no progress then leave
    If lngPercent < 1 Then
        Exit Sub
    End If
    
    ' Verify flood display has not exceeded 100%
    If lngPercent > MAX_PERCENT Then
        lngPercent = MAX_PERCENT
    End If

    With objProgBar
    
        ' Error trap in case code attempts to set
        ' scalewidth greater than the max allowable
        If lngPercent > .ScaleWidth Then
            lngPercent = .ScaleWidth
        End If
           
        .Cls                        ' Empty picture box
        .ForeColor = lngForeColor   ' Reset forecolor
     
        ' set picture box ScaleWidth equal to maximum percentage
        .ScaleWidth = MAX_PERCENT
        
        ' format percent into a displayable value (ex: 25%)
        strPercent = Format$(CLng((lngPercent / .ScaleWidth) * 100)) & "%"
        
        ' Calculate X and Y coordinates within
        ' picture box and and center data
        .CurrentX = (.ScaleWidth - .TextWidth(strPercent)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(strPercent)) \ 2
            
        objProgBar.Print strPercent   ' print percentage string in picture box
    
        ' Print flood bar up to new percent position in picture box
        objProgBar.Line (0, 0)-(lngPercent, .ScaleHeight), .ForeColor, BF
    End With

    DoEvents   ' allow flood to complete drawing

End Sub

Private Sub GetRegistryData()

    mstrFolder = GetSetting("kiCrypt", "Settings", "LastPath", App.Path & "\")
    mblnCreateNewFile = GetSetting("kiCrypt", "Settings", "OverwriteFile", True)
    mblnRetLowercase = GetSetting("kiCrypt", "Settings", "Lowercase", True)
    mlngKeyMix = GetSetting("kiCrypt", "Settings", "PrimaryKeyMix", "1")
    
End Sub

Public Sub UpdateRegistry()

    SaveSetting "kiCrypt", "Settings", "LastPath", mstrFolder
    SaveSetting "kiCrypt", "Settings", "OverwriteFile", mblnCreateNewFile
    SaveSetting "kiCrypt", "Settings", "Lowercase", mblnRetLowercase
    SaveSetting "kiCrypt", "Settings", "PrimaryKeyMix", mlngKeyMix

End Sub

Private Sub txtPwd_GotFocus()

    ' See if there is a previous password
    If mblnShowPwd Then
        ' Has password array been initialized
        If CBool(IsArrayInitialized(mabytPwd)) Then
            ' Does password array hold any data
            If UBound(mabytPwd) > (-1) Then
                txtPwd.Text = vbNullString                    ' Verify an empty textbox
                txtPwd.Text = ByteArrayToString(mabytPwd())   ' Convert previous password back to string
            End If
        End If
    End If
    
    mobjKeyEdit.TextBoxFocus txtPwd   ' Highlight contents in text box
    
End Sub

Private Sub txtPwd_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Do not copy data to/from clipboard
    mobjKeyEdit.NoCopyText txtPwd, KeyCode, Shift
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    ' Edit data input
    mobjKeyEdit.ProcessAlphaNumeric KeyAscii
End Sub

Private Sub txtPwd_LostFocus()
    
    txtPwd.Text = TrimStr(txtPwd.Text)   ' Remove unwanted characters
    
    If Len(txtPwd.Text) = 0 Then
        Erase mabytPwd()                 ' Verify an empty array
        Exit Sub
    End If
    
    If Len(txtPwd.Text) < mlngPwdLength_Min Then
        If StrComp("*****", txtPwd.Text, vbBinaryCompare) = 0 Then
            Exit Sub
        End If
    End If
    
    ' Verify there is some data here
    If Len(txtPwd.Text) >= mlngPwdLength_Min Then
        Erase mabytPwd()                              ' Verify an empty array
        mabytPwd() = StringToByteArray(txtPwd.Text)   ' Convert password to byte array
    End If
    
    txtPwd.Text = vbNullString   ' Empty password textbox
    
    ' Replace any data with asteriks.
    ' Just in case someone is using a
    ' utility to read passwords behind
    ' the asteriks.
    '
    ' Has password array been initialized
    If CBool(IsArrayInitialized(mabytPwd)) Then
        If mblnShowPwd Then
            
            ' Does password array hold any data
            If UBound(mabytPwd) > 0 Then
                txtPwd.Text = String$(UBound(mabytPwd) + 1, "*")   ' Only show asteriks
            End If
        
        ElseIf UBound(mabytPwd) > 0 Then
            
            txtPwd.Text = String$(5, "*")   ' Only show five asteriks
        Else
            
            ' Does password array hold any data
            If UBound(mabytPwd) <= 0 Then
                txtPwd.Text = String$(5, "*")   ' Only show five asteriks
            End If
        End If
    End If
    
End Sub

Private Sub txtInputString_GotFocus()
    
    ' Highlight contents in text box
    mobjKeyEdit.TextBoxFocus txtInputString
    txtOutput.Text = vbNullString

End Sub

Private Sub txtInputString_KeyDown(KeyCode As Integer, Shift As Integer)
    ' key control (Ex:   Ctrl+C, etc.)
    mobjKeyEdit.TextBoxKeyDown txtInputString, KeyCode, Shift
End Sub

Private Sub txtInputString_KeyPress(KeyAscii As Integer)
    
    ' edit data input
    Select Case KeyAscii
           Case 9
                ' Tab key
                KeyAscii = 0
                SendKeys "{TAB}"
                
           Case 8, 13, 32 To 126
                ' Backspace, ENTER key and
                ' other valid data keys
                
           Case Else  ' Everything else (invalid)
                KeyAscii = 0
    End Select

End Sub

