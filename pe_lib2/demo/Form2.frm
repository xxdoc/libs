VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2595
      Left            =   2820
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   180
      Width           =   8835
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   10927
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ShowResources(pee As CPEEditor)
    Set pe = pee
    
    Dim r As CResourceEntry
    Dim li As ListItem
    
    For Each r In pe.Resources.Entries
        Set li = lv.ListItems.Add(, , Hex(r.Size))
        Set li.Tag = r
        li.SubItems(1) = r.path
        'Debug.Print r.Report
    Next
    
    Me.Show 1
    
End Sub
 
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim r As CResourceEntry
    
    Set r = Item.Tag
    Text1 = r.Report
    
End Sub

