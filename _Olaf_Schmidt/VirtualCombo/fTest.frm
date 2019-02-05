VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Virtual (OwnerDrawn) ComboBox"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   3  'Windows-Standard
   Begin VirtualCombo.ucVirtualCombo cboVSimple 
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
   End
   Begin VirtualCombo.ucVirtualCombo cboV 
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   661
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tCountries
  Nam As String
  Chk As Boolean
End Type
Private CLst() As tCountries 'this external DataSource-Container could also be a Recordset (e.g. to support easy Sorting)

Private Sub Form_Load()
  'prepare an example for an external Data-Set (here in a UDT-Arr)
  Dim i As Long, F As String
  F = Dir(App.Path & "\Res\*.gif")
  Do While Len(F)
    ReDim Preserve CLst(i): CLst(i).Nam = Left(F, Len(F) - 4): i = i + 1
    F = Dir
  Loop
  
  'now setup the "MultiSelect-Combo"
  cboV.ListCount = i
  cboV.ItemHeight = 22
  cboV.MinVisibleItems = 15
  cboV.MultiSelect = True
  
  'finally the setup for the "Simple-Combo"
  cboVSimple.ListCount = 32 '<- setting a ListCount is all what's needed here
End Sub

'small helper, to join the currently selected Countries
Private Function GetCheckedCountries(Optional Delimiter As String = ", ") As String
  Dim i As Long, j As Long, Arr() As String
  For i = 0 To UBound(CLst)
    If CLst(i).Chk Then ReDim Preserve Arr(j): Arr(j) = CLst(i).Nam: j = j + 1
  Next
  GetCheckedCountries = Join(Arr, Delimiter)
End Function
 
'***** 2 Event-Handlers for the "Multi-Flag-Combo"-Scenario
Private Sub cboV_ListMultiClick()
  CLst(cboV.ListIndex).Chk = Not CLst(cboV.ListIndex).Chk
End Sub
Private Sub cboV_OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)
  With Canvas 'all Drawings happen "Item-wise" on a Canvas-PicBox, which is passed from inside the Virtual-ComboControl
    .FontName = "Arial": .FontSize = 10
     Canvas.Line (0, 0)-(dx, dy), IIf(IsSelected, RGB(205, 230, 255), Canvas.BackColor), BF
    
    If Index = -1 Then 'here we choose, to draw the "checked accumulation" when Index= -1 comes in
      cboV.TextOut 32, 3, IIf(Len(GetCheckedCountries), GetCheckedCountries, "<Select multiple Countries>")
      .PaintPicture LoadPicture(App.Path & "\Res\" & CLst(1).Nam & ".gif"), 0, 0, 20, 15
      .PaintPicture LoadPicture(App.Path & "\Res\" & CLst(4).Nam & ".gif"), 3, 3, 20, 15
      .PaintPicture LoadPicture(App.Path & "\Res\" & CLst(7).Nam & ".gif"), 6, 6, 20, 15
    Else
      cboV.TextOut 58, 2, CLst(Index).Nam  'Print the Country-Name
      Canvas.Line (4, 4)-(dy - 5, dy - 5), vbBlack, B 'draw an empty Rectangle for the CheckBox
      If CLst(Index).Chk Then .FontName = "WebDings": .FontSize = 13: cboV.TextOut 3, 0, "a" 'Print the Checked-State
      .PaintPicture LoadPicture(App.Path & "\Res\" & CLst(Index).Nam & ".gif"), 23, 1 'draw the Flag
    End If
  End With
End Sub

'***** and finally the two Event-Handlers for the Simple-Combo-scenario (which has no external DataSource, but renders its ListIndexes instead)
Private Sub cboVSimple_Click()
  Debug.Print "cboVSimple_Click", cboVSimple.ListIndex
End Sub
Private Sub cboVSimple_OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)
  Canvas.Line (0, 0)-(dx, dy), IIf(IsSelected, RGB(205, 230, 255), Canvas.BackColor), BF
  Canvas.FontName = "Arial": Canvas.FontSize = 10
  cboVSimple.TextOut 1, 1, Index
End Sub

