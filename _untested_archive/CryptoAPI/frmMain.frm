VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8100
   ClientLeft      =   1920
   ClientTop       =   2265
   ClientWidth     =   11235
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
   ScaleWidth      =   11235
   Visible         =   0   'False
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
      TabIndex        =   40
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
         TabIndex        =   41
         Top             =   150
         Width           =   5670
         Begin VB.OptionButton optChoice 
            Caption         =   "Random data"
            Height          =   315
            Index           =   3
            Left            =   4185
            TabIndex        =   45
            Top             =   75
            Width           =   1350
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Hash"
            Height          =   315
            Index           =   2
            Left            =   3150
            TabIndex        =   44
            Top             =   75
            Width           =   810
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "CRC-32"
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   43
            Top             =   75
            Width           =   1035
         End
         Begin VB.OptionButton optChoice 
            Caption         =   "Encrypt/Decrypt"
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   42
            Top             =   75
            Value           =   -1  'True
            Width           =   1530
         End
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
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   1440
      Width           =   11070
      Begin VB.CheckBox chkShowPwd 
         Caption         =   "chkShowPwd"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7245
         TabIndex        =   48
         Top             =   5265
         Width           =   1635
      End
      Begin VB.CheckBox chkExtraInfo 
         Caption         =   "chkExtraInfo"
         Height          =   330
         Left            =   9045
         TabIndex        =   36
         Top             =   855
         Width           =   1905
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
         Left            =   7200
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5175
         Width           =   1635
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
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "txtPwd"
         Top             =   5340
         Width           =   6930
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
         Height          =   2745
         Index           =   1
         Left            =   8955
         TabIndex        =   27
         Top             =   1215
         Width           =   1995
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
            Left            =   135
            TabIndex        =   47
            Text            =   "cboKeyLength"
            Top             =   990
            Width           =   1680
         End
         Begin VB.ComboBox cboRounds 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2250
            Width           =   600
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
            ItemData        =   "frmMain.frx":030A
            Left            =   135
            List            =   "frmMain.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1665
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
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "AES Key Length"
            Height          =   210
            Index           =   6
            Left            =   135
            TabIndex        =   46
            Top             =   765
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of hash rounds"
            Height          =   390
            Index           =   5
            Left            =   135
            TabIndex        =   39
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Encryption Algorithm"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   29
            Top             =   135
            Width           =   1665
         End
         Begin VB.Label lblAlgo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hash Algorithm"
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   28
            Top             =   1440
            Width           =   1665
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
         TabIndex        =   23
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
            TabIndex        =   0
            Text            =   "frmMain.frx":030E
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   "frmMain.frx":031F
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
            Picture         =   "frmMain.frx":032B
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   3660
            Width           =   5820
         End
         Begin VB.Label lblEncrypt 
            BackStyle       =   0  'Transparent
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   25
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
         TabIndex        =   21
         Top             =   225
         Width           =   1920
         Begin VB.OptionButton optDataType 
            Caption         =   "String data"
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   2
            Top             =   330
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton optDataType 
            Caption         =   "File"
            Height          =   240
            Index           =   1
            Left            =   1275
            TabIndex        =   3
            Top             =   330
            Width           =   600
         End
         Begin VB.Label lblAlgo 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Type"
            Height          =   240
            Index           =   0
            Left            =   495
            TabIndex        =   22
            Top             =   45
            Width           =   990
         End
      End
      Begin VB.Label lblAES_Msg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AES strong processing not available with this version of Windows"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   8955
         TabIndex        =   37
         Top             =   4305
         Width           =   1995
      End
      Begin VB.Label lblPwd 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   285
         TabIndex        =   30
         Top             =   5100
         Width           =   6300
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
      TabIndex        =   16
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
         TabIndex        =   18
         Top             =   225
         Width           =   2985
      End
      Begin VB.TextBox txtRandom 
         Appearance      =   0  'Flat
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
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "frmMain.frx":042D
         Top             =   630
         Width           =   10860
      End
      Begin VB.Label lblAlgo 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Data Types"
         Height          =   240
         Index           =   4
         Left            =   6315
         TabIndex        =   19
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
      Picture         =   "frmMain.frx":0437
      Style           =   1  'Graphical
      TabIndex        =   8
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
      Picture         =   "frmMain.frx":0741
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Picture         =   "frmMain.frx":0B83
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
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
      Picture         =   "frmMain.frx":0E8D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
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
      Picture         =   "frmMain.frx":1197
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Picture         =   "frmMain.frx":14A1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7335
      Width           =   640
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
      TabIndex        =   35
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
      TabIndex        =   34
      Top             =   7560
      Width           =   4650
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDisclaimer"
      Height          =   630
      Left            =   195
      TabIndex        =   15
      Top             =   7380
      Width           =   2670
   End
   Begin VB.Label lblEncryptMsg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblEncryptMsg"
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   7335
      TabIndex        =   14
      Top             =   855
      Width           =   3645
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
      TabIndex        =   11
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
' 17-Sep-2016  Kenneth Ives  kenaso@tx.rr.com
'              Added more security for password. See chkShowPwd_Click(),
'              Cipher_Processing(), txtPwd.GetFocus(), txtPwd.Lost_Focus()
'              routines for updates and documentation.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const BACKGRND_GRAY         As Long = &HE0E0E0
  Private Const BACKGRND_ACTIVE       As Long = vbWindowBackground
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
  Private mlngRounds        As Long
  Private mlngKeyLength     As Long
  Private mlngBlockSize     As Long
  Private mlngHashAlgo      As Long
  Private mlngCipherAlgo    As Long
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
  Private mblnAES_Ready     As Boolean
  Private mblnStringData    As Boolean
  Private mblnRetLowercase  As Boolean
  Private mblnCreateNewFile As Boolean
  Private mobjKeyEdit       As cKeyEdit
  Private mobjPrng          As kiCryptoAPI.cPrng

  Private WithEvents mobjHash   As kiCryptoAPI.cHash
Attribute mobjHash.VB_VarHelpID = -1
  Private WithEvents mobjCRC32  As kiCryptoAPI.cCRC32
Attribute mobjCRC32.VB_VarHelpID = -1
  Private WithEvents mobjCipher As kiCryptoAPI.cCipher
Attribute mobjCipher.VB_VarHelpID = -1

Private Sub cboEncrypt_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
                
    mlngCipherAlgo = cboEncrypt.ListIndex
    
    lblEncryptMsg.Visible = True
    lblPwd.Visible = True
    txtPwd.Visible = True
        
    ' Prepare combo boxes
    cboHash.Enabled = True
    cboHash.BackColor = BACKGRND_ACTIVE
    
    If mblnAES_Ready Then
        Select Case mlngCipherAlgo
               Case 0 To 3
                    lblAlgo(1).Caption = "Encryption Algorithm"
                    cboKeyLength.Enabled = False
                    cboKeyLength.BackColor = BACKGRND_GRAY
               Case 4 To 6
                    ' AES selection
                    lblAlgo(1).Caption = "AES Block Length"
                    mlngBlockSize = Val(Right$(Trim$(cboEncrypt.Text), 3))
                    cboKeyLength.Enabled = True
                    cboKeyLength.BackColor = BACKGRND_ACTIVE
               Case Else
                    lblAlgo(1).Caption = "Encryption Algorithm"
                    cboKeyLength.Enabled = False
                    cboKeyLength.BackColor = BACKGRND_GRAY
                    cboHash.Enabled = False
                    cboHash.BackColor = BACKGRND_GRAY
        End Select
    Else
        If mlngCipherAlgo = 3 Then
            cboKeyLength.Enabled = False
            cboKeyLength.BackColor = BACKGRND_GRAY
            cboHash.Enabled = False
            cboHash.BackColor = BACKGRND_GRAY
        End If
    End If
    
End Sub

Private Sub cboHash_Click()

    If mblnLoading Then
        Exit Sub
    End If
                
    cmdCopy.Enabled = False
    mlngHashAlgo = cboHash.ListIndex
    
    ' If performing encryption then leave
    If Not cboEncrypt.Enabled Then
        
        cboKeyLength.Enabled = False
        cboKeyLength.BackColor = BACKGRND_GRAY
        txtOutput.Text = vbNullString
    
    End If
    
End Sub

Private Sub cboKeyLength_Click()

    If mblnLoading Then
        Exit Sub
    End If
    
    mlngKeyLength = Val(Trim$(cboKeyLength.Text))
    txtOutput.Text = vbNullString
    
End Sub

Private Sub cboRandom_Click()
    
    If mblnLoading Then
        Exit Sub
    End If
    
    txtRandom.Text = vbNullString
    RndData_Processing cboRandom.ListIndex

End Sub

Private Sub cboRounds_Click()

    If mblnLoading Then
        Exit Sub
    End If
    
    mlngRounds = Val(Trim$(cboRounds.Text))
    txtOutput.Text = vbNullString
    
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
    Select Case mblnShowPwd
            
           Case True   ' Show actual length in asteriks
                ' Has password array been initialized
                If CBool(IsArrayInitialized(mabytPwd)) Then
                    ' Does password array hold any data
                    If UBound(mabytPwd) > 0 Then
                        txtPwd.Text = String$(UBound(mabytPwd) + 1, "*")   ' Only show asteriks
                    End If
                End If
                
           Case False   ' Show only five asteriks
                ' Has password array been initialized
                If CBool(IsArrayInitialized(mabytPwd)) Then
                    ' Does password array hold any data
                    If UBound(mabytPwd) > (-1) Then
                        txtPwd.Text = String$(5, "*")   ' Show only five asteriks
                    End If
                End If
    End Select
    
End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear                    ' Clear clipboard area
    Clipboard.SetText txtOutput.Text   ' Load clipboard with textbox data

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
        txtInputFile.Text = ShrinkToFit(mstrFilename, 70)   ' Original file name
        mstrFolder = GetFullPath(mstrFilename)              ' Capture new path
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
    
    ' If a particular object is active then
    ' set the property value
    If Not mobjCipher Is Nothing Then
        mobjCipher.StopProcessing = gblnStopProcessing
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
    mblnLoading = True
    mstrFolder = vbNullString
    Erase mabytPwd()             ' Empty password array
    
    ' Instantiate class and DLL objects
    Set mobjKeyEdit = New cKeyEdit
    Set mobjCipher = New kiCryptoAPI.cCipher
    Set mobjCRC32 = New kiCryptoAPI.cCRC32
    Set mobjHash = New kiCryptoAPI.cHash
    Set mobjPrng = New kiCryptoAPI.cPrng
    
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

        ' If operating system is Windows 8 or 8.1,
        ' form caption is centered automatically
        If Not gblnWin8or81 Then
            CenterCaption frmMain   ' Manually center form caption
        End If
    
        .lblFormTitle.Caption = PGM_NAME
        .lblDisclaimer.Caption = "This software is provided without any " & _
                                 "warrantees or guarantees implied or intended."
        .lblOperSystem.Caption = gstrOperSystem
        .lblEncryptMsg.Visible = False
        
        .txtRandom.BackColor = BACKGRND_GRAY
        
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
        cboHash_Click
        cboRounds_Click
        ResetProgressBar
        
        ' Center the form on the screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless   ' reduce flicker
        .Refresh
    End With

    mblnLoading = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    gblnStopProcessing = True

    ' If the object is still active then
    ' send a command to stop it.
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

Private Sub mobjCipher_CipherProgress(ByVal lngProgress As Long)
    
    If mblnStringData Then
        Exit Sub
    End If
    
    ProgressBar picProgressBar, lngProgress, vbBlue
    DoEvents
    
End Sub

Private Sub mobjCRC32_CRCProgress(ByVal lngProgress As Long)
    
    ProgressBar picProgressBar, lngProgress, vbRed
    DoEvents
    
End Sub

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
                    .cboEncrypt.BackColor = BACKGRND_ACTIVE
                    
                    If mlngCipherAlgo > 3 Then
                        .cboKeyLength.Enabled = True
                        .cboKeyLength.BackColor = BACKGRND_ACTIVE
                    Else
                        .cboKeyLength.Enabled = False
                        .cboKeyLength.BackColor = BACKGRND_GRAY
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
                    .cboEncrypt.BackColor = BACKGRND_GRAY
                    .cboKeyLength.Enabled = False
                    .cboKeyLength.BackColor = BACKGRND_GRAY
                    
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
    
    ' Encryption algorithms
    With frmMain

        mblnAES_Ready = mobjPrng.AES_Ready
        
        If mblnAES_Ready Then
        
            ' Advanced processing is available
            With .cboHash
                .Clear
                .AddItem "MD2 (32-bit)"       ' 0
                .AddItem "MD4 (32-bit)"       ' 1
                .AddItem "MD5 (32-bit)"       ' 2
                .AddItem "SHA-1 (32-bit)"     ' 3
                .AddItem "SHA-256 (32-bit)"   ' 4
                .AddItem "SHA-384 (64-bit)"   ' 5
                .AddItem "SHA-512 (64-bit)"   ' 6
                .ListIndex = 3                ' Default = SHA-1
            End With
            
            With .cboEncrypt
                .Clear
                .AddItem "RC2"       ' 0
                .AddItem "RC4"       ' 1
                .AddItem "DES"       ' 2
                .AddItem "3DES"      ' 3
                .AddItem "AES-128"   ' 4
                .AddItem "AES-192"   ' 5
                .AddItem "AES-256"   ' 6
                .AddItem "BASE64"    ' 7
                .ListIndex = 1       ' Default = RC4
            End With
    
            With .cboKeyLength
                .Clear
                .AddItem "128"       ' 0
                .AddItem "192"       ' 1
                .AddItem "256"       ' 2
                .ListIndex = 0       ' Default = 128 bit
                .Enabled = False     ' Deactivate
                .BackColor = BACKGRND_GRAY
            End With
                
            .lblAlgo(6).Visible = True
            .lblAES_Msg.Visible = False  ' Hide message at bottom of form
            
        Else
            ' Advanced processing NOT available
            With .cboHash
                .Clear
                .AddItem "MD2 (32-bit)"     ' 0
                .AddItem "MD4 (32-bit)"     ' 1
                .AddItem "MD5 (32-bit)"     ' 2
                .AddItem "SHA-1 (32-bit)"   ' 3
                .ListIndex = 3              ' Default = SHA-1
            End With
            
            With .cboEncrypt
                .Clear
                .AddItem "RC2"      ' 0
                .AddItem "RC4"      ' 1
                .AddItem "DES"      ' 2
                .AddItem "BASE64"   ' 3
                .ListIndex = 1      ' Default = RC4
            End With
    
            .cboKeyLength.Visible = False
            .lblAlgo(6).Visible = False
            .lblAES_Msg.Visible = True   ' Show message at bottom of form
            
        End If
        
        ' Number of rounds for hashing
        With .cboRounds
            .Clear
            For lngIdx = 1 To 10
                .AddItem CStr(lngIdx)
            Next lngIdx
            .ListIndex = 0   ' Default = 1 round
        End With
        
        ' Random data combobox
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
    End With
    
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

    Dim blnEncrypt    As Boolean
    Dim strMsg        As String
    Dim strData       As String
    Dim strTemp       As String
    Dim strOutputFile As String
    Dim astrMsgBox()  As String
    Dim abytData()    As Byte
    Dim hFile         As Long
    Dim lngHold       As Long
    Dim lngLength     As Long
    Dim lngEncrypt    As enumCIPHER_ACTION
    
    Const ENCRYPT_EXT As String = ".ENC"
    Const DECRYPT_EXT As String = ".DEC"

    Erase abytData()         ' Always start with empty arrays
    Erase astrMsgBox()
    strTemp = vbNullString   ' Empty temp variable
    lngHold = 0
    
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
    
    ' Reset for Base64
    If Not mblnAES_Ready Then
        If mlngCipherAlgo = 3 Then
            lngHold = mlngCipherAlgo
            mlngCipherAlgo = 7
        End If
    End If
    
    ' *********************************************************
    ' Encrypt/Decrypt - String
    ' *********************************************************
    If mblnStringData Then
                    
        Screen.MousePointer = vbHourglass
        
        With mobjCipher
            .StopProcessing = False                ' Set flag to continue processing
            .HashMethod = mlngHashAlgo             ' Type of hash algorithm selected
            .CipherMethod = mlngCipherAlgo         ' Type of cipher algorithm selected
            .Blocksize = mlngBlockSize             ' AES block size (128, 192, 256)
            .Keylength = mlngKeyLength             ' AES key length (128, 192, 256)
            .Password = mabytPwd()                 ' Process password
            gblnStopProcessing = .StopProcessing   ' See if processing aborted
        
            DoEvents
            If gblnStopProcessing Then
                GoTo Cipher_Processing_CleanUp
            End If
        
            Select Case lngEncrypt
                   Case eMSG_ENCRYPT    ' Encrypt string
                        
                        blnEncrypt = True
            
                        strData = Trim$(txtInputString.Text)      ' Remove leading and trailing blanks
                        abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                        
                        ' Encrypt data string
                        If .StringProcessing(abytData(), blnEncrypt) Then
                        
                            Select Case mlngCipherAlgo
                                    
                                   Case 0 To 6
                                        strData = ByteArrayToHex(abytData())  ' convert single charaters to hex
                                         
                                        DoEvents
                                        If gblnStopProcessing Then
                                            GoTo Cipher_Processing_CleanUp
                                        End If
                    
                                        If mblnRetLowercase Then
                                            strData = LCase$(strData)   ' Convert string to lowercase
                                        Else
                                            strData = UCase$(strData)   ' Convert string to uppercase
                                        End If
                                    
                                   Case 7   ' Base64
                                        strData = ByteArrayToString(abytData())
                            End Select
                        Else
                            gblnStopProcessing = .StopProcessing  ' See if processing aborted
                        End If
                
                        txtInputString.Text = vbNullString   ' Empty text box
                        txtInputString.Text = strData        ' Store string data in text box
                
                   Case eMSG_DECRYPT    ' Decrypt string
                        
                        DoEvents
                        If gblnStopProcessing Then
                            txtPwd.SetFocus
                            GoTo Cipher_Processing_CleanUp
                        End If
                        
                        blnEncrypt = False
            
                        With mobjCipher
                            .StopProcessing = False                ' Set flag to continue processing
                            .HashMethod = mlngHashAlgo             ' Type of hash algorithm selected
                            .CipherMethod = mlngCipherAlgo         ' Type of cipher algorithm selected
                            .Blocksize = mlngBlockSize             ' AES block size (128, 192, 256)
                            .Keylength = mlngKeyLength             ' AES key length (128, 192, 256)
                            .Password = mabytPwd()                 ' Process password
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
                        
                            DoEvents
                            If gblnStopProcessing Then
                                GoTo Cipher_Processing_CleanUp
                            End If
                
                            strData = Trim$(txtInputString.Text)   ' Remove leading and trailing blanks
                            
                            Select Case mlngCipherAlgo
                                    
                                   Case 0 To 6
                                        abytData() = HexToByteArray(strData)      ' Convert hex to single char
                                   Case 7   ' Base64
                                        abytData() = StringToByteArray(strData)   ' Convert string data to byte array
                            End Select
                                            
                            DoEvents
                            If gblnStopProcessing Then
                                GoTo Cipher_Processing_CleanUp
                            End If
        
                            ' Decrypt data string
                            If .StringProcessing(abytData(), blnEncrypt) Then
                                strData = ByteArrayToString(abytData())   ' Convert byte array to string data
                            Else
                                GoTo Cipher_Processing_CleanUp
                            End If
                                            
                        End With
                
                        txtInputString.Text = vbNullString   ' Empty text box
                        txtInputString.Text = strData        ' Store string data in text box
            End Select
        End With
            
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
            
        End If
        
        With mobjCipher
            .StopProcessing = False                ' Set flag to continue processing
            .HashMethod = mlngHashAlgo             ' Type of hash algorithm selected
            .CipherMethod = mlngCipherAlgo         ' Type of cipher algorithm selected
            .Blocksize = mlngBlockSize             ' AES block size (128, 192, 256)
            .Keylength = mlngKeyLength             ' AES key length (128, 192, 256)
            .Password = mabytPwd()                 ' Process password
            .CreateNewFile = mblnCreateNewFile     ' True - Create new output file
                                                   ' False - Overwrite input file
            gblnStopProcessing = .StopProcessing   ' See if processing aborted
            
            DoEvents
            If gblnStopProcessing Then
                GoTo Cipher_Processing_CleanUp
            End If

            Select Case lngEncrypt
                   Case eMSG_ENCRYPT    ' Encrypt file
                        
                        blnEncrypt = True
                        
                        ' Verify overwrite message
                        If Not mblnCreateNewFile Then
                            
                            If ResponseMsg("Are you sure you want to overwrite input file?", _
                                           vbYesNo, "Verify output target") = vbNo Then
                                           
                                txtInputFile.SetFocus
                                GoTo Cipher_Processing_CleanUp
                            End If
                            
                            strOutputFile = mstrFilename
                        End If
                        
                        If .FileProcessing(mstrFilename, blnEncrypt) Then
                        
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
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
                        End If
                
                   Case eMSG_DECRYPT    ' Decrypt file
                   
                        blnEncrypt = False
                        
                        ' Verify overwrite message
                        If Not mblnCreateNewFile Then
                            
                            If ResponseMsg("Are you sure you want to overwrite input file?", _
                                           vbYesNo, "Verify output target") = vbNo Then
                                           
                                txtInputFile.SetFocus
                                GoTo Cipher_Processing_CleanUp
                            End If
                            
                            strOutputFile = mstrFilename
                        End If
                        
                        If .FileProcessing(mstrFilename, blnEncrypt) Then
                            
                            gblnStopProcessing = .StopProcessing   ' See if processing aborted
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
                        End If
            End Select
        End With
    End If
    
Cipher_Processing_CleanUp:
    Erase abytData()         ' Always empty arrays when not needed
    Erase astrMsgBox()
    strTemp = vbNullString   ' Empty temp variable
    
    ' Reset selection back to original value
    If lngHold > 0 Then
        mlngCipherAlgo = lngHold
    End If
    
    Screen.MousePointer = vbDefault  ' Set mouse pointer back to normal
    
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
            .StopProcessing = False                               ' Set flag to continue processing
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
    txtOutput.Text = vbNullString
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
        .StopProcessing = False               ' Set flag to continue processing
        .HashMethod = mlngHashAlgo            ' Hash algorithm selected
        .HashRounds = mlngRounds              ' Number of rounds for hashing
        .ReturnLowercase = mblnRetLowercase   ' TRUE = Return as lowercase
                                              ' FALSE = Return as uppercase
        ' Hash string data
        If optDataType(0).Value Then
            
            abytData() = StringToByteArray(txtInputString.Text)   ' Convert to byte array
            abytHash() = .HashString(abytData())                  ' Hash string data
            gblnStopProcessing = .StopProcessing                  ' See if processing aborted
            strOutput = ByteArrayToString(abytHash())             ' Convert byte array to string
            
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
        .StopProcessing = False   ' Set flag to continue processing
        
        Select Case lngChoice
    
               Case 0   ' ASCII String
                    txtRandom.Text = .BuildWithinRange(1400, 32, 126, ePRNG_ASCII)
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
               Case 1   ' Hex String
                    txtRandom.Text = .BuildRndData(800, ePRNG_HEX)
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
               Case 2   ' Hex Array
                    avntData = .BuildRndData(440, ePRNG_HEX_ARRAY)
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
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
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
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
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
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
                    gblnStopProcessing = .StopProcessing   ' See if processing aborted
    
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

    mstrFolder = GetSetting("kiCryptoAPI", "Settings", "LastPath", App.Path & "\")
    mblnCreateNewFile = GetSetting("kiCryptoAPI", "Settings", "OverwriteFile", True)
    mblnRetLowercase = GetSetting("kiCryptoAPI", "Settings", "Lowercase", True)
    
End Sub

Public Sub UpdateRegistry()

    SaveSetting "kiCryptoAPI", "Settings", "LastPath", mstrFolder
    SaveSetting "kiCryptoAPI", "Settings", "OverwriteFile", mblnCreateNewFile
    SaveSetting "kiCryptoAPI", "Settings", "Lowercase", mblnRetLowercase

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

Private Sub txtInputString_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' key control (Ex:   Ctrl+V, etc.)
    mobjKeyEdit.TextBoxKeyDown txtInputString, KeyCode, Shift
    
End Sub


