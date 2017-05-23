VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{DD72A457-E7F1-11D5-B111-0004AC98CB59}#1.0#0"; "sptbdock.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "cEdit Code Editor"
   ClientHeight    =   5235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13245
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin TabDock.TTabDock fDock 
      Left            =   4080
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      BorderStyle     =   7
      CaptionStyle    =   4
      Persistant      =   -1  'True
      Gradient1       =   0
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   370
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   13245
      TabIndex        =   0
      Top             =   4545
      Width           =   13245
      Begin MSComctlLib.ImageList TabImage 
         Left            =   4800
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483644
         ImageWidth      =   10
         ImageHeight     =   10
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1042
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":150C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TabStrip tb 
         Height          =   360
         Left            =   0
         TabIndex        =   1
         Top             =   15
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   635
         TabFixedWidth   =   1766
         HotTracking     =   -1  'True
         Placement       =   1
         ImageList       =   "TabImage"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3480
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13115
            MinWidth        =   547
            Text            =   "Welcome to cEdit Final"
            TextSave        =   "Welcome to cEdit Final"
            Object.ToolTipText     =   "Shows Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Object.ToolTipText     =   "Shows Cursor Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Language"
            TextSave        =   "Language"
            Object.ToolTipText     =   "Shows Current Language"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img 
      Left            =   2880
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   47
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":247A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3030
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3582
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4026
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4138
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":424A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":435C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":446E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5024
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5576
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":601A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":656C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7010
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7562
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8006
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8558
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":954E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A544
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA96
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B53A
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BFDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C530
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CA82
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CFD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D526
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DA78
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DFCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F01C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F56E
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FAC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10012
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   635
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "img"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   47
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New Document"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open Document"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "close"
            Object.ToolTipText     =   "Close Open Document"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "save"
            Object.ToolTipText     =   "Save Document"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "saveall"
            Object.ToolTipText     =   "Save All Open Documents"
            ImageIndex      =   44
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "saveas"
            Object.ToolTipText     =   "Save Document As"
            ImageIndex      =   45
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "reload"
            Object.ToolTipText     =   "Reload Open Document"
            ImageIndex      =   46
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "print"
            Object.ToolTipText     =   "Print Document"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cut"
            Object.ToolTipText     =   "Cut Selected Text"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "copy"
            Object.ToolTipText     =   "Copy Selected Text"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "paste"
            Object.ToolTipText     =   "Paste Clipboard Contents"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Selected Text"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prop"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "findnext"
            Object.ToolTipText     =   "Find Next"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "findprev"
            Object.ToolTipText     =   "Find Previous"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tabl"
            Object.ToolTipText     =   "Tab Right"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tabr"
            Object.ToolTipText     =   "Tab Left"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cblock"
            Object.ToolTipText     =   "Comment Block"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ublock"
            Object.ToolTipText     =   "Uncomment Block"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "tbmark"
            Object.ToolTipText     =   "Toggle Bookmark"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "nbmark"
            Object.ToolTipText     =   "Next Bookmark"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pbmark"
            Object.ToolTipText     =   "Previous Bookmark"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cbmark"
            Object.ToolTipText     =   "Clear Bookmarks"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "pline"
            Object.ToolTipText     =   "Previous Line"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "nline"
            Object.ToolTipText     =   "Next Line"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ctag"
            Object.ToolTipText     =   "Custom Tag"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button42 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button43 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tilehor"
            Object.ToolTipText     =   "Tile Horizontaly"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button44 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tilever"
            Object.ToolTipText     =   "Tile Verticly"
            ImageIndex      =   29
         EndProperty
         BeginProperty Button45 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cascade"
            Object.ToolTipText     =   "Cascade"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button46 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button47 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   31
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMacro 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   635
      ButtonWidth     =   1931
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "img"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 1"
            Key             =   "mac1"
            Object.ToolTipText     =   "Macro 1"
            ImageIndex      =   33
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 2"
            Key             =   "mac2"
            Object.ToolTipText     =   "Play Macro 2"
            ImageIndex      =   34
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 3"
            Key             =   "mac3"
            Object.ToolTipText     =   "Play Macro 3"
            ImageIndex      =   35
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 4"
            Key             =   "mac4"
            Object.ToolTipText     =   "Play Macro 4"
            ImageIndex      =   36
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 5"
            Key             =   "mac5"
            Object.ToolTipText     =   "Play Macro 5"
            ImageIndex      =   37
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 6"
            Key             =   "mac6"
            Object.ToolTipText     =   "Play Macro 6"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 7"
            Key             =   "mac7"
            Object.ToolTipText     =   "Play Macro 7"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 8"
            Key             =   "mac8"
            Object.ToolTipText     =   "Play Macro 8"
            ImageIndex      =   40
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 9"
            Key             =   "mac9"
            Object.ToolTipText     =   "Play Macro 9"
            ImageIndex      =   41
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Macro 10"
            Key             =   "mac10"
            Object.ToolTipText     =   "Play Macro 10"
            ImageIndex      =   42
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Mac"
            Key             =   "cmac"
            Object.ToolTipText     =   "Create Macro"
            ImageIndex      =   43
         EndProperty
      EndProperty
   End
   Begin VB.Menu file
      Caption         =   "&File"
      Begin VB.Menu new
            Caption         =   "&New Document"
            Shortcut         =   ^N
      End
      Begin VB.Menu mnuNewTemplate
            Caption         =   "&New From Template"
         Begin VB.Menu mnuTemplate
                  Caption         =   ""
                  Index         =   0
         End
      End
      Begin VB.Menu mnuBar10
            Caption         =   "-"
      End
      Begin VB.Menu open
            Caption         =   "&Open Document"
            Shortcut         =   ^O
      End
      Begin VB.Menu close
            Caption         =   "&Close"
      End
      Begin VB.Menu bar0
            Caption         =   "-"
      End
      Begin VB.Menu save
            Caption         =   "&Save Document"
            Shortcut         =   ^S
      End
      Begin VB.Menu saveas
            Caption         =   "Save Document &As"
            Shortcut         =   {F12}
      End
      Begin VB.Menu saveall
            Caption         =   "Save A&ll"
      End
      Begin VB.Menu bar30
            Caption         =   "-"
      End
      Begin VB.Menu saveftp
            Caption         =   "&FTP"
         Begin VB.Menu openftp
                  Caption         =   "&Open From FTP"
         End
         Begin VB.Menu saveto
                  Caption         =   "&Save To FTP"
         End
      End
      Begin VB.Menu bar1
            Caption         =   "-"
      End
      Begin VB.Menu prints
            Caption         =   "&Print"
            Shortcut         =   ^P
      End
      Begin VB.Menu printsetup
            Caption         =   "Printer &Setup"
      End
      Begin VB.Menu bar2
            Caption         =   "-"
      End
      Begin VB.Menu properties
            Caption         =   "&Document Properties"
      End
      Begin VB.Menu bar3
            Caption         =   "-"
      End
      Begin VB.Menu mnuRecent
            Caption         =   "&Recent Files"
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   0
         End
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   1
                  Visible         =   0
         End
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   2
                  Visible         =   0
         End
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   3
                  Visible         =   0
         End
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   4
                  Visible         =   0
         End
         Begin VB.Menu mnuRec
                  Caption         =   ""
                  Index         =   5
                  Visible         =   0
         End
      End
      Begin VB.Menu mnuBar5
            Caption         =   "-"
      End
      Begin VB.Menu exit
            Caption         =   "&Exit"
      End
   End
   Begin VB.Menu edit
      Caption         =   "&Edit"
      Begin VB.Menu undo
            Caption         =   "&Undo"
            Shortcut         =   ^Z
      End
      Begin VB.Menu redo
            Caption         =   "&Redo"
            Shortcut         =   ^Y
      End
      Begin VB.Menu bar7
            Caption         =   "-"
      End
      Begin VB.Menu cut
            Caption         =   "&Cut"
            Shortcut         =   ^X
      End
      Begin VB.Menu copy
            Caption         =   "C&opy"
            Shortcut         =   ^C
      End
      Begin VB.Menu paste
            Caption         =   "&Paste"
            Shortcut         =   ^V
      End
      Begin VB.Menu delete
            Caption         =   "&Delete"
      End
      Begin VB.Menu bar4
            Caption         =   "-"
      End
      Begin VB.Menu mnuComment
            Caption         =   "Comment &Block"
      End
      Begin VB.Menu mnuUncomment
            Caption         =   "Uncomment B&lock"
      End
      Begin VB.Menu mnuBar4
            Caption         =   "-"
      End
      Begin VB.Menu selectall
            Caption         =   "&Select All"
            Shortcut         =   ^A
      End
      Begin VB.Menu selectline
            Caption         =   "Select &Line"
      End
      Begin VB.Menu bar5
            Caption         =   "-"
      End
      Begin VB.Menu datetime
            Caption         =   "Date/Time"
            Shortcut         =   {F7}
      End
   End
   Begin VB.Menu search
      Caption         =   "&Search"
      Begin VB.Menu find
            Caption         =   "&Find"
            Shortcut         =   ^F
      End
      Begin VB.Menu findnext
            Caption         =   "Find &Next"
            Shortcut         =   {F3}
      End
      Begin VB.Menu findprev
            Caption         =   "Find &Previous"
            Shortcut         =   ^{F3}
      End
      Begin VB.Menu mnuReplace
            Caption         =   "&Replace"
            Shortcut         =   ^H
      End
      Begin VB.Menu bar8
            Caption         =   "-"
      End
      Begin VB.Menu goto
            Caption         =   "&Goto Line..."
            Shortcut         =   ^G
      End
      Begin VB.Menu bar100
            Caption         =   "-"
      End
      Begin VB.Menu mnuToggle
            Caption         =   "&Toggle Bookmark"
      End
      Begin VB.Menu mnuNext
            Caption         =   "&Next Bookmark"
      End
      Begin VB.Menu mnuPrev
            Caption         =   "&Previous Bookmark"
      End
      Begin VB.Menu mnuClear
            Caption         =   "&Clear Bookmarks"
      End
      Begin VB.Menu bar101
            Caption         =   "-"
      End
      Begin VB.Menu mnuNLine
            Caption         =   "Next &Line"
      End
      Begin VB.Menu mnuLPrev
            Caption         =   "Previous L&ine"
      End
      Begin VB.Menu bar102
            Caption         =   "-"
      End
      Begin VB.Menu countall
            Caption         =   "Count &All"
            Shortcut         =   ^{F5}
      End
   End
   Begin VB.Menu language
      Caption         =   "&Language"
      Begin VB.Menu lang
            Caption         =   "Text"
            Index         =   0
      End
      Begin VB.Menu lang
            Caption         =   "C/C++"
            Index         =   1
      End
      Begin VB.Menu lang
            Caption         =   "Basic"
            Index         =   2
      End
      Begin VB.Menu lang
            Caption         =   "Java"
            Index         =   3
      End
      Begin VB.Menu lang
            Caption         =   "Pascal"
            Index         =   4
      End
      Begin VB.Menu lang
            Caption         =   "SQL"
            Index         =   5
      End
      Begin VB.Menu lang
            Caption         =   "HTML"
            Index         =   6
      End
      Begin VB.Menu lang
            Caption         =   "XML"
            Index         =   7
      End
      Begin VB.Menu lang
            Caption         =   "-"
            Index         =   8
      End
   End
   Begin VB.Menu view
      Caption         =   "&View"
      Begin VB.Menu editor
            Caption         =   "&Editor Options"
      End
      Begin VB.Menu fileassoc
            Caption         =   "&File Associations"
      End
      Begin VB.Menu bar11
            Caption         =   "-"
      End
      Begin VB.Menu template
            Caption         =   "&Template Editor"
      End
      Begin VB.Menu mnuBar15
            Caption         =   "-"
      End
      Begin VB.Menu toolbar
            Caption         =   "Standard &Toolbar"
            Checked         =   -1
      End
      Begin VB.Menu mnuMacBar
            Caption         =   "&Macro Toolbar"
      End
      Begin VB.Menu mnuBar3
            Caption         =   "-"
      End
      Begin VB.Menu statusbar2
            Caption         =   "MDI &Tab View"
            Checked         =   -1
      End
      Begin VB.Menu mnuBar6
            Caption         =   "-"
      End
      Begin VB.Menu quicknav
            Caption         =   "Quick Nav"
      End
      Begin VB.Menu MDebugOutput
            Caption         =   "Debug Output"
      End
      Begin VB.Menu bar20
            Caption         =   "-"
      End
      Begin VB.Menu hlline
            Caption         =   "Highlight Selected Line"
      End
      Begin VB.Menu whitespace
            Caption         =   "&White Spaces"
      End
   End
   Begin VB.Menu mnuBuild
      Caption         =   "&Compile"
      Begin VB.Menu mnuCompile
            Caption         =   "&Build/Compile"
            Shortcut         =   {F5}
      End
      Begin VB.Menu mnuBar7
            Caption         =   "-"
      End
      Begin VB.Menu mnuBuildConfig
            Caption         =   "&Configure Build Settings"
      End
   End
   Begin VB.Menu mnuMacro
      Caption         =   "&Macros"
      Begin VB.Menu mac
            Caption         =   "Macro 1"
            Index         =   1
      End
      Begin VB.Menu mac
            Caption         =   "Macro 2"
            Index         =   2
      End
      Begin VB.Menu mac
            Caption         =   "Macro 3"
            Index         =   3
      End
      Begin VB.Menu mac
            Caption         =   "Macro 4"
            Index         =   4
      End
      Begin VB.Menu mac
            Caption         =   "Macro 5"
            Index         =   5
      End
      Begin VB.Menu mac
            Caption         =   "Macro 6"
            Index         =   6
      End
      Begin VB.Menu mac
            Caption         =   "Macro 7"
            Index         =   7
      End
      Begin VB.Menu mac
            Caption         =   "Macro 8"
            Index         =   8
      End
      Begin VB.Menu mac
            Caption         =   "Macro 9"
            Index         =   9
      End
      Begin VB.Menu mac
            Caption         =   "Macro 10"
            Index         =   10
      End
      Begin VB.Menu mnuBar2
            Caption         =   "-"
      End
      Begin VB.Menu mnuSave
            Caption         =   ""&Save Macro"
      End
      Begin VB.Menu mnuBar1
            Caption         =   "-"
      End
      Begin VB.Menu mnuCreate
            Caption         =   "Create Macro"
      End
   End
   Begin VB.Menu mnuPlugins
      Caption         =   "&Plugins"
      Begin VB.Menu mnuPlugin
            Caption         =   "No Plugins Installed"
            Enabled         =   0
            Index         =   0
      End
   End
   Begin VB.Menu window
      Caption         =   "&Window"
      Begin VB.Menu tilehor
            Caption         =   "Tile Horrizontal"
      End
      Begin VB.Menu tilever
            Caption         =   "Tile Vertical"
      End
      Begin VB.Menu arrangeicons
            Caption         =   "Arrange Icons"
      End
      Begin VB.Menu cascade
            Caption         =   "&Casade"
      End
      Begin VB.Menu bar12
            Caption         =   "-"
      End
      Begin VB.Menu closeall
            Caption         =   "Close All Windows"
      End
      Begin VB.Menu bar13
            Caption         =   "-"
      End
      Begin VB.Menu inbrowser
            Caption         =   "Show File in Browser"
      End
      Begin VB.Menu wnlist
            Caption         =   "Window List"
            WindowList         =   -1
      End
   End
   Begin VB.Menu mnuLinks
      Caption         =   "&Links"
      Begin VB.Menu mnuPSC
            Caption         =   "&Planet Source Code"
      End
      Begin VB.Menu mnuFree
            Caption         =   "SourceCode For Free"
      End
      Begin VB.Menu mnuVB
            Caption         =   "&FreeVBCode"
      End
      Begin VB.Menu mnuVBA
            Caption         =   "&VisualBasic Accelerator"
      End
      Begin VB.Menu mnuBar16
            Caption         =   "-"
      End
      Begin VB.Menu mnucEdit
            Caption         =   "&cEdit Homepage"
      End
   End
   Begin VB.Menu help
      Caption         =   "&Help"
      Begin VB.Menu genhelp
            Caption         =   "General Help"
            Shortcut         =   {F1}
      End
      Begin VB.Menu online
            Caption         =   "Help Online"
            Shortcut         =   ^U
      End
      Begin VB.Menu bar22
            Caption         =   "-"
      End
      Begin VB.Menu readme
            Caption         =   "&Readme"
            Shortcut         =   {F8}
      End
      Begin VB.Menu bar14
            Caption         =   "-"
      End
      Begin VB.Menu acksoftsite
            Caption         =   "&cEdit Website"
            Shortcut         =   ^{F4}
      End
      Begin VB.Menu bar15
            Caption         =   "-"
      End
      Begin VB.Menu about
            Caption         =   "&About"
            Shortcut         =   {F4}
      End
   End
   Begin VB.Menu tabmenu
      Caption         =   "TabMenu"
      Visible         =   0
      Begin VB.Menu close2
            Caption         =   "&Close"
      End
      Begin VB.Menu bar40
            Caption         =   "-"
      End
      Begin VB.Menu save2
            Caption         =   "&Save"
      End
      Begin VB.Menu saveas2
            Caption         =   "Save &As"
      End
      Begin VB.Menu bar41
            Caption         =   "-"
      End
      Begin VB.Menu print2
            Caption         =   "&Print"
      End
   End

End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
Private Sub about_Click()
  On Error Resume Next
  frmAbout.Show vbModal, frmMain
End Sub

Private Sub acksoftsite_Click()
  On Error Resume Next
  OpenURL "http://cedit.sourceforge.net", Me.hwnd
End Sub

Private Sub arrangeicons_Click()
  On Error Resume Next
  Me.Arrange vbArrangeIcons
End Sub

Private Sub cascade_Click()
  On Error Resume Next
  Me.Arrange vbCascade
End Sub

Private Sub close_Click()
  On Error Resume Next
  'Document(dnum).Visible = False
  Unload Document(dnum)
End Sub

Private Sub close2_Click()
  On Error Resume Next
  Unload Document(dnum)
End Sub

Private Sub closeall_Click()
  On Error Resume Next
  CloseAllDoc
End Sub
Private Sub CloseAllDoc()
  On Error Resume Next
  LockWindowUpdate Me.hwnd
  Dim x As Integer
  For x = 1 To UBound(Document)
    'Document(X).Visible = False
    Unload Document(x)
    
    If StopClose = True Then Exit For
  Next
  LockWindowUpdate 0
End Sub

Private Sub copy_Click()
  On Error Resume Next
  Document(dnum).rt.copy
End Sub

Private Sub countall_Click()
  On Error Resume Next
  Dim ua2() As String, us As Integer, ut As Integer
  ua2 = Split(Document(dnum).rt.Text, " ")
  us = Len(Document(dnum).rt.Text)
  ut = Document(dnum).rt.LineCount
  MsgBox "Words: " & UBound(ua2) + 1 & Chr(10) & "Characters:" & us & Chr(10) & "Lines: " & ut, vbOKOnly + vbInformation, "Count All"
  Erase ua2
End Sub

Private Sub cut_Click()
  On Error Resume Next
  Document(dnum).rt.cut
End Sub

Private Sub datetime_Click()
  On Error Resume Next
  Dim timedate As String
  timedate = Date & "/" & Time
  InsertString Document(dnum).rt, timedate
End Sub

Private Sub delete_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdDelete
End Sub


Private Sub editor_Click()
  On Error Resume Next
  frmDoc.rt.ExecuteCmd cmCmdProperties
  WriteOptions
End Sub

Private Sub exit_Click()
  On Error Resume Next
  Unload Me
  Unload frmDoc
  Unload frmAbout
  End
End Sub



Private Sub fDock_FormHide(ByVal DockedForm As TabDock.TDockForm)
  On Error Resume Next
  Select Case DockedForm.Key
    Case "frmNav"
      quicknav.Checked = False
    Case "frmOutput"
      MDebugOutput.Checked = False
  End Select
End Sub

Private Sub fDock_FormShow(ByVal DockedForm As TabDock.TDockForm)
  On Error Resume Next
  Select Case DockedForm.Key
    Case "frmNav"
      quicknav.Checked = True
    Case "frmOutput"
      MDebugOutput.Checked = True
  End Select
End Sub

Private Sub fileassoc_Click()
  On Error Resume Next
  frmNew.Show vbModal, Me
End Sub

Private Sub find_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdFind
End Sub

Private Sub findnext_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdFindNext
End Sub

Private Sub findprev_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdFindPrev
End Sub

Private Sub genhelp_Click()
  On Error Resume Next
  HHShowContents Me.hwnd
End Sub

Private Sub goto_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdGotoLine, -1
End Sub

Private Sub hlline_Click()
  On Error Resume Next
  Dim x As Integer
  If hlline.Checked = False Then
    hlline.Checked = True
    HighLight = True
    For x = 1 To UBound(Document)
      Set Document(x).r = Document(x).rt.GetSel(True)
      Document(x).rt.HighlightedLine = Document(x).r.EndColNo
    Next
  Else
    hlline.Checked = False
    HighLight = False
    For x = 1 To UBound(Document)
      Document(x).rt.HighlightedLine = -1
    Next
  End If
  WriteInput
End Sub

Private Sub inbrowser_Click()
  On Error Resume Next
  ShowSite "about:" & Document(dnum).rt.Text
End Sub

Private Sub lang_Click(Index As Integer)
  On Error Resume Next
  Dim x As Integer
  Select Case Index
    Case 0
      Document(dnum).rt.language = ""
    Case 1
      Document(dnum).rt.language = "c/c++"
    Case 2
      Document(dnum).rt.language = "basic"
    Case 3
      Document(dnum).rt.language = "java"
    Case 4
      Document(dnum).rt.language = "perl"
    Case 5
      Document(dnum).rt.language = "pascal"
    Case 6
      Document(dnum).rt.language = "sql"
    Case 7
      Document(dnum).rt.language = "html"
    Case 8
      Document(dnum).rt.language = "xml"
    Case 9
      Document(dnum).rt.language = "css"
    Case Else
      Document(dnum).rt.language = lang(Index).Caption
  End Select
  For x = 0 To lang.Count - 1
    lang(x).Checked = False
  Next
  lang(Index).Checked = True
End Sub

Private Sub mac_Click(Index As Integer)
  On Error Resume Next
  Select Case Index
    Case 1
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro1
    Case 2
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro2
    Case 3
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro3
    Case 4
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro4
    Case 5
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro5
    Case 6
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro6
    Case 7
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro7
    Case 8
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro8
    Case 9
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro9
    Case 10
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro10
  End Select
End Sub


Private Sub MDIForm_Load()
  On Error Resume Next
  FlatBorder TB.hwnd
  LoadEditor
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
    Dim OLEFilename As String
    Dim I As Integer
    
    For I = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            OLEFilename = Data.Files(I)
        End If
        On Error GoTo errexit
        DoOpen OLEFilename
    Next I
errexit:
    Exit Sub
End Sub

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  On Error Resume Next
    If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If dnum = 0 Then Exit Sub
  CloseAllDoc
  If StopClose = True Then
    StopClose = False
    Cancel = 1
  End If
End Sub

Private Sub MDIForm_Resize()
  On Error Resume Next
  TB.Left = 0
  TB.Width = picBottom.ScaleWidth
  stBar.Panels(1).Width = (Me.Width - stBar.Panels(2).Width - stBar.Panels(3).Width - stBar.Panels(4).Width - 450)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  On Error Resume Next
  Dim x As Integer
  For x = 0 To 9
    SaveMacros App.path & "\macros\" & x & ".dem", x
  Next
  WriteData
  WriteInput
  
  'UnloadAll
End Sub

Private Sub mnuBuildConfig_Click()
  On Error Resume Next
  frmBuild.Show vbModal, Me
End Sub

Private Sub mnucEdit_Click()
  ShowSite "http://www.sourceforge.net/projects/cedit"
End Sub

Private Sub mnuClear_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdBookmarkClearAll
End Sub

Private Sub mnuComment_Click()
  On Error Resume Next
  Document(dnum).CommentBlock
End Sub

Private Sub mnuCompile_Click()
  On Error Resume Next
  Dim CaptureOut As String
  Dim s As String
  Dim lang As String, Exe As String, Comp As String, Variables As String
  Dim RunComp As String, InForOut As String, file As String, FileToCompile As String
  Dim Found As Boolean, VarRead As String
  'Dim dnum As Integer, Found As Boolean, VarRead As String, FileToCompile As String
  s = Dir(App.path & "\compile\")
  Found = False
  Do While s <> ""
    If Right(s, 3) = "cmp" Then
      file = App.path & "\compile\" & s
      lang = ReadINI("Compile", "Language", file)
      Exe = ReadINI("Compile", "Extension", file)
      If LCase(lang) = LCase(Document(dnum).rt.language) <> 0 Or GetExtension(Document(dnum).Caption) = LCase(Exe) Then
        Found = True
        Exit Do
      End If
    End If
    s = Dir
  Loop
  If Found = False Then
    MsgBox "No compiler found for this file type or language.", vbOKOnly + vbCritical, "Build"
    Exit Sub
  End If
  If Document(dnum).FTP = True Then
    Document(dnum).rt.SaveFile App.path & "\data\tmp." & GetExtension(Document(dnum).filename), False
    FileToCompile = App.path & "\data\tmp." & GetExtension(Document(dnum).filename)
  ElseIf Document(dnum).FTP = False And Document(dnum).filename <> "" Then
    doSave
    FileToCompile = Document(dnum).filename
    'Document(dnum).rt.SaveFile Document(dnum).filename, False
  Else
    FileToCompile = App.path & "\data\tmp." & Exe
    Document(dnum).rt.SaveFile App.path & "\data\tmp." & Exe, False
  End If
  Comp = ReadINI("Compile", "Compile", file)
  Variables = ReadINI("Compile", "Variables", file)
  RunComp = ReadINI("Compile", "RunWhenComplete", file)
  InForOut = ReadINI("Compile", "InputForOutput", file)
  Variables = Replace(Variables, "%s", StrWrap(FileToCompile))
  CaptureOut = ReadINI("Compile", "CaptureOutput", file)
  If InForOut = "on" Then
    VarRead = InputStr("Enter the filename you would like this outputed to. (IE: hello.exe)", "Write Name")
    Variables = Replace(Variables, "%e", VarRead)
  End If
  If Dir(Comp) = "" Then
    MsgBox "Compiler not found.", vbOKOnly + vbCritical, "Error"
    Exit Sub
  End If
  If CaptureOut = "on" Then
    fDock.FormShow ("frmOutput")
    frmOutput.txtOut.Text = "Compilation in progress..."
    DoEvents
    MDebugOutput.Checked = True
    ChDir Mid(FileToCompile, 1, InStrRev(FileToCompile, "\"))
    frmOutput.txtOut.Text = GetCommandOutput(StrWrap(Comp) & " " & Variables)
    frmOutput.txtOut.SelStart = Len(frmOutput.txtOut.Text)
  Else
    Shell StrWrap(Comp) & " " & Variables, vbNormalFocus
  End If
  If InForOut = "on" And RunComp = "on" Then
    Shell VarRead, vbNormalFocus
  End If
End Sub

Private Sub mnuCreate_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdRecordMacro
End Sub

Private Sub mnuFree_Click()
  ShowSite "http://www.sourcecode4free.com"
End Sub

Private Sub mnuLPrev_Click()
  On Error Resume Next
  Document(dnum).PrevLine
End Sub

Private Sub mnuMacBar_Click()
  On Error Resume Next
  If mnuMacBar.Checked = True Then
    tbMacro.Visible = False
    mnuMacBar.Checked = False
  Else
    tbMacro.Visible = True
    mnuMacBar.Checked = True
  End If
End Sub

Private Sub mnuNext_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdBookmarkNext
End Sub

Private Sub mnuNLine_Click()
  On Error Resume Next
  Document(dnum).NextLine
End Sub

Private Sub mnuPlugin_Click(Index As Integer)
  On Error Resume Next
  Call RunPlugin(mnuPlugin(Index).Tag, Me) ' Execute the plug-in
End Sub

Private Sub mnuPrev_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdBookmarkPrev
End Sub

Private Sub mnuPSC_Click()
  ShowSite "http://www.pscode.com"
End Sub

Private Sub mnuRec_Click(Index As Integer)
  On Error Resume Next
  DoOpen mnuRec(Index).Caption
End Sub

Private Sub mnuReplace_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdFindReplace
End Sub

Private Sub mnuSave_Click()
  On Error Resume Next
  Dim x As Integer
  For x = 0 To 9
    SaveMacros App.path & "\macros\" & x & ".dem", x
  Next
End Sub

Private Sub mnuTemplate_Click(Index As Integer)
  LoadTemplate mnuTemplate(Index).Tag
End Sub

Private Sub mnuToggle_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdBookmarkToggle
End Sub

Private Sub mnuUncomment_Click()
  On Error Resume Next
  Document(dnum).UncommentBlock
End Sub

Private Sub mnuVB_Click()
  ShowSite "http://www.freevbcode.com"
End Sub

Private Sub mnuVBA_Click()
  ShowSite "http://www.vbaccelerator.com"
End Sub

Private Sub new_Click()
  On Error Resume Next
  doNew ""
End Sub

Private Sub online_Click()
  On Error Resume Next
  OpenURL "http://cedit.sourceforge.net/doc/index.html", Me.hwnd
End Sub

Private Sub picBottom_Resize()
'  On Error Resume Next
  TB.Move 0, 0, picBottom.ScaleWidth, picBottom.ScaleHeight
End Sub

Private Sub print2_Click()
  On Error Resume Next
  Call Document(dnum).rt.PrintContents(0, cmPrnColor + cmPrnDefaultPrn + cmPrnRichFonts)

End Sub

Private Sub quicknav_Click()
  On Error Resume Next
  If quicknav.Checked = True Then
    fDock.FormHide ("frmNav")
    quicknav.Checked = False
  Else
    quicknav.Checked = True
    fDock.FormShow ("frmNav")
  End If
End Sub

Private Sub open_Click()
  On Error Resume Next
  cd.CancelError = True
  cd.DialogTitle = "Open a document..."
  cd.Filter = AllSupport & FilterB
  cd.ShowOpen
  If cd.filename = "" Then Exit Sub
  DoOpen cd.filename
  AddRecent cd.filename
End Sub

Private Sub openftp_Click()
  On Error Resume Next
  frmFTP.Caption = "Open Document"
  frmFTP.cmdOpen.Caption = "&Open"
  frmFTP.Show , Me
End Sub

Private Sub paste_Click()
  On Error Resume Next
  Document(dnum).rt.paste
End Sub

Private Sub Prints_Click()
  On Error Resume Next
  Call Document(dnum).rt.PrintContents(0, cmPrnColor + cmPrnDefaultPrn + cmPrnRichFonts)
End Sub

Private Sub printsetup_Click()
  On Error Resume Next
  Call Document(dnum).rt.PrintContents(0, cmPrnColor + cmPrnRichFonts)
End Sub

Private Sub properties_Click()
  On Error Resume Next
  Dim UA() As String, kB As Double
  kB = (Len(Document(dnum).rt.Text) / 1024)
  UA() = Split(Document(dnum).rt.Text, " ")
  With frmProperties
    .lblChar = "Characters: " & Len(Document(dnum).rt.Text)
    .lblLine = "Total Lines: " & Document(dnum).rt.LineCount
    .lblWord = "Word Count: " & UBound(UA) + 1
    If Left(Document(dnum).Caption, 12) = "New Document" Then
      .lblFile = "File Name: " & "New Document"
    Else
      .lblFile = "File Name: " & Document(dnum).Caption
    End If
    .lblSizeK = "File Size(K): " & kB & " KBytes"
    .lblSizeB = "File Size(B): " & Len(Document(dnum).rt.Text) & " Bytes"
    .lblData(0).Caption = Document(dnum).Caption
    .Show vbModal, frmMain
  End With
  Erase UA
End Sub

Private Sub readme_Click()
  On Error Resume Next
  DoOpen App.path & "\Readme.txt"
End Sub

Private Sub redo_Click()
  On Error Resume Next
  Document(dnum).rt.redo
End Sub


Private Sub save_Click()
  On Error Resume Next
  If Document(dnum).FTP = True And FState(dnum).Deleted = False Then
    frmUpload.cboAccount.Text = Document(dnum).FTPAccount
    frmUpload.cboAccount.Enabled = False
    DoEvents
    frmUpload.Show
    frmUpload.Refresh
    frmUpload.PutFile Document(dnum).filename, Document(dnum).rt.Text, Document(dnum).ftpDir
    Document(dnum).Changed = False
    Document(dnum).FTP = True
    Document(dnum).ftpDir = CurDir
    Document(dnum).DoAct
    
    Unload frmUpload
  Else
    doSave
  End If
End Sub

Private Sub save2_Click()
  On Error Resume Next
  If Document(dnum).FTP = True Then
      frmUpload.Show , frmMain
  Else
    doSave
  End If
End Sub

Private Sub saveall_Click()
  On Error Resume Next
  Dim x As Integer, y As Integer
  y = dnum
  For x = 1 To UBound(Document)
    Document(x).SetFocus
    doSave
  Next
  Document(y).SetFocus
End Sub

Private Sub saveas_Click()
  On Error Resume Next
  doSaveAs
  
End Sub

Private Sub saveas2_Click()
  On Error Resume Next
  doSaveAs
End Sub

Private Sub saveto_Click()
  On Error Resume Next
  frmFTP.Caption = "Save Document"
  frmFTP.cmdOpen.Caption = "&Save"
  frmFTP.SaveString = ActiveForm.rt.Text
  frmFTP.Show
End Sub

Private Sub selectall_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdSelectAll
End Sub

Private Sub selectline_Click()
  On Error Resume Next
  Document(dnum).rt.ExecuteCmd cmCmdSelectLine
End Sub


Private Sub statusbar2_Click()
  statusbar2.Checked = Not statusbar2.Checked
  picBottom.Visible = statusbar2.Checked
End Sub



Private Sub tb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  Dim dnum As String
  dnum = (Mid(TB.SelectedItem.Key, 4))
  Document(dnum).SetFocus
End Sub

Private Sub tb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  Dim dnum As String
  dnum = (Mid(TB.SelectedItem.Key, 4))
  Document(dnum).SetFocus
  If Button = vbRightButton Then
    Button = vbLeftButton
    PopupMenu tabmenu
  End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim QuickTag As String
  Select Case Button.Key
    Case "new"
      doNew ""
    Case "close"
      Unload Document(dnum)
    Case "prop"
      frmDoc.rt.ExecuteCmd cmCmdProperties
      WriteOptions
    Case "reload"
      If Document(dnum).IsFile = False Then Exit Sub
      Document(dnum).rt.OpenFile Document(dnum).Caption
    Case "find"
      Document(dnum).rt.ExecuteCmd cmCmdFind
    Case "findnext"
      Document(dnum).rt.ExecuteCmd cmCmdFindNext
    Case "findprev"
      Document(dnum).rt.ExecuteCmd cmCmdFindPrev
    Case "undo"
      Document(dnum).rt.undo
      SetDo
    Case "saveas"
      saveas_Click
    Case "saveall"
      saveall_Click
    Case "redo"
      Document(dnum).rt.redo
      SetDo
    Case "tilever"
      Me.Arrange vbTileVertical
    Case "tilehor"
      Me.Arrange vbTileHorizontal
    Case "cascade"
      Me.Arrange vbCascade
    Case "cut"
      Document(dnum).rt.cut
    Case "paste"
      Document(dnum).rt.paste
    Case "copy"
      Document(dnum).rt.copy
    Case "delete"
      Document(dnum).rt.ExecuteCmd cmCmdDelete
    Case "open"
      open_Click
    Case "print"
      Call Document(dnum).rt.PrintContents(0, cmPrnColor + cmPrnDefaultPrn + cmPrnRichFonts)
    Case "save"
      doSave
    Case "tabl"
      Document(dnum).rt.ExecuteCmd cmCmdIndentSelection
    Case "tabr"
      Document(dnum).rt.ExecuteCmd cmCmdUnindentSelection
    Case "cblock"
      Document(dnum).CommentBlock
    Case "ublock"
      Document(dnum).UncommentBlock
    Case "tbmark"
      Document(dnum).rt.ExecuteCmd cmCmdBookmarkToggle
    Case "nbmark"
      Document(dnum).rt.ExecuteCmd cmCmdBookmarkNext
    Case "pbmark"
      Document(dnum).rt.ExecuteCmd cmCmdBookmarkPrev
    Case "cbmark"
      Document(dnum).rt.ExecuteCmd cmCmdBookmarkClearAll
    Case "pline"
      Document(dnum).PrevLine
    Case "nline"
      Document(dnum).NextLine
    Case "ctag"
      QuickTag = InputStr("Enter the HTML tag to insert", "Quick Tag", "<>", 1)
      If QuickTag <> "" Then InsertString Document(dnum).rt, QuickTag
    Case "help"
      HHShowContents Me.hwnd
  End Select
End Sub

Private Sub tbMacro_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Select Case LCase(Button.Key)
    Case "mac1"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro1
    Case "mac2"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro2
    Case "mac3"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro3
    Case "mac4"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro4
    Case "mac5"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro5
    Case "mac6"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro6
    Case "mac7"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro7
    Case "mac8"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro8
    Case "mac9"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro9
    Case "mac10"
      Document(dnum).rt.ExecuteCmd cmCmdPlayMacro10
    Case "cmac"
      Document(dnum).rt.ExecuteCmd cmCmdRecordMacro
  End Select
End Sub



Private Sub tbQuick_ButtonClick(ByVal Button As MSComctlLib.Button)
  On Error Resume Next
  Dim l As String
  Open App.path & "\qbar\" & Button.Key For Input As #1
    l = Input(LOF(1), 1)
  Close 1
  InsertString Document(dnum).rt, l
End Sub

Private Sub template_Click()
  frmTemplate.Show vbModal, Me
End Sub

Private Sub tilehor_Click()
  On Error Resume Next
  Me.Arrange vbTileHorizontal
End Sub

Private Sub tilever_Click()
  On Error Resume Next
  Me.Arrange vbTileVertical
End Sub


Private Sub toolbar_Click()
'  On Error Resume Next
'  If toolbar.Checked = True Then
'    toolbar.Checked = False
'    tBar.Visible = False
'  Else
'    toolbar.Checked = True
'    tBar.Visible = True
'  End If
  toolbar.Checked = Not toolbar.Checked
  tBar.Visible = toolbar.Checked
End Sub

Private Sub undo_Click()
  On Error Resume Next
  Document(dnum).rt.undo
End Sub

Private Sub WhiteSpace_Click()
  On Error Resume Next
  Dim x As Integer
  If whitespace.Checked = False Then
    For x = 1 To UBound(Document)
      Document(x).rt.DisplayWhitespace = True
    Next
    whitespace.Checked = True
  Else
    For x = 1 To UBound(Document)
      Document(x).rt.DisplayWhitespace = False
    Next
    whitespace.Checked = False
  End If
  WriteInput
End Sub

Private Function SaveMacros(ByVal sFileName As String, ByVal nMacroNum As Long) As Boolean
  On Error Resume Next
    Dim bArr() As Byte
    Dim hFile As Integer
    Dim g As CodeMaxCtl.globals
    Set g = New CodeMaxCtl.globals
    g.GetMacro nMacroNum, bArr
    If UBound(bArr) >= 0 Then
        hFile = FreeFile
        On Error Resume Next
        Open sFileName For Binary Access Write As #hFile
          Put #hFile, , bArr
        Close #hFile
        If Err.Number Then
            Exit Function
        End If
        SaveMacros = True
    End If
End Function

Private Sub LoadMacros()
  On Error Resume Next
  Dim s As String
  s = Dir(App.path & "\macros\")
  Do Until s = ""
    If Right(s, 3) = "dem" Then
      AddMacro App.path & "\macros\" & s, Left(s, InStr(1, s, ".") - 1)
    End If
    s = Dir
  Loop
End Sub

Private Sub AddMacro(file As String, macNum As Long)
  On Error Resume Next
  Dim p As CodeMaxCtl.globals
  Set p = New CodeMaxCtl.globals
  Dim fFile As Integer, bBar() As Byte
  fFile = FreeFile()
  Open file For Binary Access Read As #fFile
    ReDim bBar(0 To LOF(fFile))
    Get fFile, , bBar
  Close #fFile
  p.SetMacro macNum, bBar
End Sub

Private Sub SetDo()
  On Error Resume Next
  If Document(dnum).rt.CanUndo Then
    tBar.Buttons("undo").Enabled = True
  Else
    tBar.Buttons("undo").Enabled = False
  End If
  If Document(dnum).rt.CanRedo Then
    tBar.Buttons("redo").Enabled = True
  Else
    tBar.Buttons("redo").Enabled = False
  End If
  
End Sub

Private Sub LoadEditor()
'  On Error Resume Next
  Dim hk As CodeMaxCtl.HotKey, hk_index As Integer
  Dim num_hk As Long, cmGlobals As CodeMaxCtl.globals
  Dim cmd(7) As CodeMaxCtl.cmCommand, cmd_index As Integer
  Set cmGlobals = New CodeMaxCtl.globals
  Set hk = New CodeMaxCtl.HotKey
  fDock.GrabMain Me.hwnd
  
  fDock.AddForm frmNav, tdDocked, tdAlignLeft, "frmNav", tdDockLeft Or tdDockFloat Or tdDockRight
  fDock.AddForm frmOutput, tdDocked, tdAlignBottom, "frmOutput", tdDockBottom Or tdDockFloat
  addTags
  fDock.Show
  fDock.FormHide frmOutput
  TB.Tabs.Remove 1
  LoadMacros
  ReadData
  ReadOptions frmDoc.rt
  ReadInput
  'setup the default color settings (Used to set highlight language based on extension
  ClrString = "c:c/c++ cpp:c/c++ h:c/c++ java:java asp:html sql:sql bas:basic cls:basic xml:xml htm:html pas:pascal frm:basic vbp:basic ctl:basic html:html java:java"
  
  'Setup the first chunk for the filters on the dialogs
  AllSupport = "All Files|*.*|All Supported Files|*txt;*.htm;*.cls;*.sql;*.html;*css;*.js;*.c;*.cpp;*.h;*.pl;*.cgi;*.xml;*.pas;*.bas;*.frm;*.vbp"
  
  'Setup the second chunk for the filters on the dialogs
  FilterB = "|Text Files|*.txt|Html Files|*.html;*.htm|Java Script Files|*.js|Style Sheets|*.cs|C/C++ Files|*.c;*.cpp;*.h|Perl Files|*.pl|CGI/Perl Files|*.cgi|XML Files|*.xml|Pascal Files|*.pas|Basic Files|*.bas;*.cls;*.frm;*.vbp|SQL Files|*.sql"
  Langs = ""
  RegisterAll
  LoadTemplates
  'Unregister a few of the hotkeys in the codemax control
  cmd(1) = cmCmdCut
  cmd(2) = cmCmdPaste
  cmd(3) = cmCmdCopy
  cmd(4) = cmCmdLineCut
  cmd(7) = cmCmdLineDelete
  cmd(5) = cmCmdUndo
  cmd(6) = cmCmdRedo
  LoadRecent
  
  AddPlugins Me
  
  For cmd_index = 1 To 7
     num_hk = cmGlobals.GetNumHotKeysForCmd(cmd(cmd_index))
     For hk_index = num_hk - 1 To 0 Step -1
       Set hk = cmGlobals.GetHotKeyForCmd(cmd(cmd_index), hk_index)
       Call cmGlobals.UnregisterHotKey(hk)
     Next hk_index
  Next cmd_index
  If Command = "" Then
    doNew ""
  Else
    Dim OpnStr As String
    OpnStr = Command
    If Left$(Command, 1) = Chr$(34) Then
      OpnStr = Right$(OpnStr, Len(OpnStr) - 1)
    End If
    If Right$(Command, 1) = Chr$(34) Then
      OpnStr = Left$(OpnStr, Len(OpnStr) - 1)
    End If
    DoOpen OpnStr
  End If
  StopClose = False
  
  
  
End Sub
Private Sub MDebugOutput_Click()
    On Error Resume Next
    If MDebugOutput.Checked = True Then
        fDock.FormHide ("frmOutput")
        MDebugOutput.Checked = False
    Else
        MDebugOutput.Checked = True
        fDock.FormShow ("frmOutput")
    End If

End Sub

'**************************************************************
'* The following functions are for use with the plugin code   *
'**************************************************************

Public Sub AddText(str As String)
  If dnum = 0 Then Exit Sub
  InsertString Document(dnum).rt, str
End Sub

Public Sub MessageBox(Optional msgStr As String, Optional msgStyle As VbMsgBoxStyle, Optional msgTitle As String)
  MsgBox msgStr, msgStyle, msgTitle
End Sub


