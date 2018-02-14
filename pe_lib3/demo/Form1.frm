VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A143468-B450-48DD-930D-925078198E4D}#1.1#0"; "hexed.ocx"
Begin VB.Form frmResList 
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   14340
   StartUpPosition =   3  'Windows Default
   Begin rhexed.HexEd he 
      Height          =   4755
      Left            =   3420
      TabIndex        =   2
      Top             =   2040
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8387
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   3420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   10755
   End
   Begin MSComctlLib.ListView lv 
      Height          =   6675
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   11774
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
   End
End
Attribute VB_Name = "frmResList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pe As New CPEEditor
Dim selRes As CResData

Sub ShowResources(p As CPEEditor)
    
    Set pe = p
    
    Dim r As CResData
    Dim li As ListItem
    
    For Each r In pe.Resources.Entries
        Set li = lv.ListItems.add(, , r.path)
        Set li.Tag = r
    Next
    
    If lv.ListItems.Count > 0 Then lv_ItemClick lv.ListItems(1)
        
    
End Sub

Private Sub Form_Load()
    mnuPopup.Visible = False
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim r As CResData
    Dim offset As Long
    
    Set r = Item.Tag
    Set selRes = r
    
    offset = pe.RvaToOffset(r.OffsetToDataRVA)
    he.LoadFile pe.LoadedFile
    he.scrollTo offset - &H20
    he.SelStart = offset
    he.SelLength = r.size - 1
    
    Text1 = "FileOffset: " & Hex(offset) & vbCrLf & r.Report
    
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuSaveAs_Click()
    
    Dim b() As Byte, r As String, fp As String, f As Long
    
    fp = "C:\res.bin"
    If pe.Resources.GetResourceData(selRes.path, b) Then
        MsgBox "Loaded byte array Size: " & UBound(b) & vbCrLf & HexDump2(b)
    Else
        MsgBox "failed to getresourcedata by path? : " & selRes.path
    End If
    
    If pe.Resources.SaveResource(fp, selRes.path) Then
        MsgBox "Saved as " & fp
    Else
        MsgBox "Save to file failed?"
    End If

End Sub
