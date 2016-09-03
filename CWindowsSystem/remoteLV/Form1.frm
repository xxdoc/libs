VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ilLarge 
      Left            =   870
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilSmall 
      Left            =   1650
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilHeader 
      Left            =   2520
      Top             =   5430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvWinList 
      Height          =   2595
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvDupe 
      Height          =   2595
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   4577
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dz note: dep stops the image list code from working apparently..

'========================================================
'Form code
'Add 2 ListView (lvWinList and lvDupe) and 3 ImageLists (ilLarge, ilSmall, ilHeader)
'=========================================================
Option Explicit

Private Sub Form_Load()
   Caption = "List View duplicate demo"
   lvWinList.View = lvwReport
   GetWindowList lvWinList, "SysListView32"
   If lvWinList.ListItems.Count Then
      lvWinList_ItemClick lvWinList.ListItems(1)
   End If
End Sub

Private Sub Form_Resize()
   If WindowState = vbMinimized Then Exit Sub
   lvWinList.Move 0, 0, lvWinList.Width, ScaleHeight
   lvDupe.Move lvWinList.Width + 30, 0, ScaleWidth - lvWinList.Width - 60, ScaleHeight
End Sub

Private Sub lvWinList_ItemClick(ByVal Item As MSComctlLib.ListItem)
   MousePointer = vbHourglass
   LV_Duplicate Item.Tag, lvDupe ', ilLarge, ilSmall, ilHeader
   MousePointer = vbDefault
End Sub


'=================================================
'Enjoy :)
'=================================================
