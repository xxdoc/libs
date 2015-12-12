VERSION 5.00
Begin VB.Form frmMenuEditor 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Menu Builder"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6600
   Icon            =   "frmMenuEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6600
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   5340
      TabIndex        =   14
      Top             =   1080
      Width           =   1155
   End
   Begin VB.ListBox lstMenu 
      Height          =   2400
      Left            =   120
      TabIndex        =   35
      Top             =   3360
      Width           =   6435
   End
   Begin VB.PictureBox picForm 
      BackColor       =   &H80000011&
      BorderStyle     =   0  '없음
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   6555
      TabIndex        =   26
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtFormCaption 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   360
         Width           =   4995
      End
      Begin VB.TextBox txtFormTag 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   270
         Left            =   3960
         TabIndex        =   32
         Top             =   60
         Width           =   2475
      End
      Begin VB.TextBox txtFormName 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   60
         Width           =   1935
      End
      Begin VB.TextBox txtFormPath 
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   660
         Width           =   4995
      End
      Begin VB.Label lblFormCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Form Caption:"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label lblFormTag 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Tag:"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   3480
         TabIndex        =   31
         Top             =   105
         Width           =   390
      End
      Begin VB.Label lblFormName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Form Name:"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label lblFormPath 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Path:"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   705
         Width           =   435
      End
   End
   Begin VB.ComboBox cboNegotiationPosition 
      Height          =   300
      ItemData        =   "frmMenuEditor.frx":08CA
      Left            =   4380
      List            =   "frmMenuEditor.frx":08DA
      Style           =   2  '드롭다운 목록
      TabIndex        =   23
      Top             =   2160
      Width           =   2115
   End
   Begin VB.TextBox txtHelpContextID 
      Height          =   270
      Left            =   1440
      TabIndex        =   22
      Text            =   "0"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox chkWindowList 
      Caption         =   "Window List"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  '없음
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   135
      TabIndex        =   20
      Top             =   3360
      Width           =   135
   End
   Begin VB.ComboBox cboShortCut 
      Height          =   300
      ItemData        =   "frmMenuEditor.frx":0909
      Left            =   4380
      List            =   "frmMenuEditor.frx":09FD
      Style           =   2  '드롭다운 목록
      TabIndex        =   6
      Top             =   1800
      Width           =   2115
   End
   Begin VB.TextBox txtIndex 
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   330
      Left            =   4020
      TabIndex        =   12
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   330
      Left            =   2700
      TabIndex        =   11
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "Visible"
      Height          =   255
      Left            =   3420
      TabIndex        =   5
      Top             =   2520
      Value           =   1  '확인
      Width           =   975
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Value           =   1  '확인
      Width           =   1095
   End
   Begin VB.CheckBox chkChecked 
      Caption         =   "Checked"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   5340
      TabIndex        =   15
      Top             =   1440
      Width           =   1155
   End
   Begin VB.TextBox txtName 
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   1425
      Width           =   4395
   End
   Begin VB.TextBox txtCaption 
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   4395
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   3
      Left            =   1560
      Picture         =   "frmMenuEditor.frx":0CDC
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   2940
      Width           =   420
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   2
      Left            =   1080
      Picture         =   "frmMenuEditor.frx":100A
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   2940
      Width           =   420
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   1
      Left            =   600
      Picture         =   "frmMenuEditor.frx":1338
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   2940
      Width           =   420
   End
   Begin VB.CommandButton cmdPos 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "frmMenuEditor.frx":1666
      Style           =   1  '그래픽
      TabIndex        =   7
      Top             =   2940
      Width           =   420
   End
   Begin VB.Label lblNegotiatePosition 
      AutoSize        =   -1  'True
      Caption         =   "Negotiate Position:"
      Height          =   180
      Left            =   2640
      TabIndex        =   25
      Top             =   2220
      Width           =   1650
   End
   Begin VB.Label lblHelpcontextID 
      Caption         =   "HelpcontextID:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6480
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6480
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Label lblShortcut 
      AutoSize        =   -1  'True
      Caption         =   "Shortcut:"
      Height          =   180
      Left            =   3480
      TabIndex        =   19
      Top             =   1860
      Width           =   750
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1860
      Width           =   615
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   1455
      Width           =   570
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Caption :"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   1110
      Width           =   765
   End
   Begin VB.Menu mnuEdit
      Caption         =   "&Edit"
      Index         =   0
      NegotiatePosition         =   3   '오른쪽
      Begin VB.Menu mnuEditClear
            Caption         =   "Clear Menu"
      End
      Begin VB.Menu mnuEditTemplate
            Caption         =   "Replace with Template"
      End
   End
   Begin VB.Menu mnuFile
      Caption         =   "&File"
      NegotiatePosition         =   1   '왼쪽
      Begin VB.Menu mnuFileNew
            Caption         =   "&New Menu"
            Shortcut         =   ^N
            HelpContextID         =   1234
      End
      Begin VB.Menu mnuFileOpen
            Caption         =   "&Open Form"
            Shortcut         =   ^O
      End
      Begin VB.Menu mnuFileOpenTemplate
            Caption         =   "Open &Template"
            Shortcut         =   ^T
      End
      Begin VB.Menu mnuFileSpace
            Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveForm
            Caption         =   "Save Form"
            Shortcut         =   ^S
            Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveMenuAsForm
            Caption         =   "Save Menu As New Form"
            Shortcut         =   ^{F12}
      End
      Begin VB.Menu mnuFileSaveMenuAsTemplate
            Caption         =   "Save Menu As Template"
            Shortcut         =   +{F12}
            Checked         =   -1   'True
      End
      Begin VB.Menu mnuFileSpace1
            Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit
            Caption         =   "Exit"
            Shortcut         =   ^{F4}
      End
   End
   Begin VB.Menu mnuView
      Caption         =   "&View"
      Begin VB.Menu mnuViewInfo
            Caption         =   "View Menu &Info"
      End
      Begin VB.Menu mnuViewWalk
            Caption         =   "&Walk Menu"
      End
   End


End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Enum eOpenFileType
   OFT_FILE
   OFT_TEMPLATE
   OFT_CMD_TEMPLATE
End Enum

Private Enum eSaveFileType
   SFT_SAVE_FORM
   SFT_SAVEAS_FORM
   SFT_SAVEAS_TEMPLATE
End Enum

Private m_VBForm As New VBForm
Private m_VBMenus As New VBMenus
Private m_SeledtedMenu As VBMenu

Private Sub Form_Load()
   ListSetTabStop lstMenu.HWnd, 150
   OpenFile OFT_CMD_TEMPLATE, Command()
   'Associate App.Path + "\BoboMenuBuilder.exe", ".bmu"
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub mnuFile_Click()
   mnuFileSaveForm.Enabled = FileExists(m_VBForm.FileName)
End Sub

'========== Save File/ Save Template ===================
Private Sub cmdOK_Click()
   SaveFile SFT_SAVE_FORM
End Sub
Private Sub mnuFileSaveForm_Click()
   SaveFile SFT_SAVE_FORM
End Sub
Private Sub mnuFileSaveMenuAsForm_Click()
   SaveFile SFT_SAVEAS_FORM
End Sub
Private Sub mnuFileSaveMenuAsTemplate_Click()
   SaveFile SFT_SAVEAS_TEMPLATE
End Sub

Private Sub SaveFile(Index As eSaveFileType)

   On Error GoTo Ooops

   Const DLGTITLE_SAVEAS_FORM As String = "메뉴 저장 폼"
   Const DLGTITLE_SAVEAS_TEMPLATE As String = "메뉴 템플레이트 저장"
   Const FF_VBF_FILE As String = "Visual Basic 폼/컨트롤 (*.frm;*.ctl;*.pag)" & _
                                                vbNullChar & "*.frm;*.ctl;*.pag" & vbNullChar
   Const FF_VBM_FILE As String = "Visual Basic 메뉴 템플레이트 (*.vbm)" & _
                                                vbNullChar & "*.vbm" & vbNullChar

   Dim sfile As String
   Dim sDialogTitle As String
   Dim sFilter As String

   Select Case Index
   Case SFT_SAVE_FORM
      If LenB(m_VBForm.FileName) = 0 Then
         SaveFile SFT_SAVEAS_FORM
         Exit Sub
      End If
   Case SFT_SAVEAS_FORM
      sDialogTitle = DLGTITLE_SAVEAS_FORM
      sFilter = FF_VBF_FILE
   Case SFT_SAVEAS_TEMPLATE
      sDialogTitle = DLGTITLE_SAVEAS_TEMPLATE
      sFilter = FF_VBM_FILE
   Case Else
      Exit Sub
   End Select

   If Index <> SFT_SAVE_FORM Then
      sfile = SelectSaveFile(, Me.HWnd, sDialogTitle, , , , sFilter & FF_ALL_FILE)
      If LenB(sfile) = 0 Then
         Exit Sub '>---> Bottom
      End If
      m_VBForm.FileName = sfile
      txtFormPath.Text = sfile
      
      With m_VBForm
         Select Case Index
         Case SFT_SAVE_FORM, SFT_SAVEAS_FORM
            lblFormName.Caption = "Form Name:"
         Case SFT_SAVEAS_TEMPLATE
            lblFormName.Caption = "Template:"
            .Name = GetFileName(.FileName, efpBaseName)
         End Select
         
         txtFormCaption.Visible = Not (Index = SFT_SAVEAS_TEMPLATE)
         lblFormTag.Visible = Not (Index = SFT_SAVEAS_TEMPLATE)
         txtFormTag.Visible = Not (Index = SFT_SAVEAS_TEMPLATE)
         
         txtFormPath.Refresh
         txtFormName.Text = .Name
         txtFormCaption.Refresh
         lblFormTag.Refresh
         txtFormTag.Refresh
      End With
   End If
   
   Dim bSuccess As Boolean
   Screen.MousePointer = vbHourglass
   If Not CheckMenu() Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   bSuccess = m_VBForm.SaveMenus(False, True, True)
   Screen.MousePointer = vbDefault
   If Not bSuccess Then
      MsgBox "Error occurred during saving the menus."
      Exit Sub
   End If

Ooops:

End Sub

Private Function CheckMenu() As Boolean
   Dim ErrorItem As VBMenu
   Dim ErrorItemIndex As Long
   Dim ErrorDescription As String
   
   Debug.Assert m_VBMenus.Count
   If m_VBMenus.Count = 0 Then
      MsgBox "Error! No menus created."
   End If
   If m_VBMenus.Validate(ErrorItem, ErrorItemIndex, ErrorDescription) <> VBM_ERR_NONE Then
      lstMenu.ListIndex = ErrorItemIndex - 1
      Screen.MousePointer = vbDefault
      MsgBox ErrorDescription, vbCritical
      Exit Function '>---> Bottom
   End If
   CheckMenu = True
End Function

'========== New/ Open File/ Open Template ===================
Private Sub mnuFileNew_Click()

   Set m_VBForm = New VBForm
   With m_VBForm
      'Set reference for later use
      Set m_VBMenus = .Menus

      .FileName = vbNullString
      .Name = "Form1"
      .Caption = "Form1"
      txtFormPath.Text = .FileName
      txtFormName.Text = .Name
      txtFormCaption.Text = .Caption
      txtFormTag.Text = .Tag
      lblFormName.Caption = "Form Name:"
   End With 'M_VBFORM

   m_VBMenus.Clear
   With m_VBMenus.Add
      .Level = 1
   End With
   lstMenu.Clear
   lstMenu.AddItem vbNullString
   lstMenu.ListIndex = 0
End Sub

Private Sub mnuFileOpen_Click()
   OpenFile OFT_FILE
End Sub

Private Sub mnuFileOpenTemplate_Click()
   OpenFile OFT_TEMPLATE
End Sub

Private Sub OpenFile(Index As eOpenFileType, Optional CmdLineFileName As String)

'Open file

   'On Error GoTo Ooops

   Const DIALOG_FORM As String = "Visual Basic 폼 파일 열기"
   Const DIALOG_TEMPLATE As String = "Visual Basic 메뉴 템플레이트 파일 열기"
   Const FF_VBF_FILE As String = "Visual Basic 폼/컨트롤 (*.frm;*.ctl;*.pag)" & _
         vbNullChar & "*.frm;*.ctl;*.pag" & vbNullChar
   Const FF_VBM_FILE As String = "Visual Basic 메뉴 템플레이트 (*.vbm)" & _
         vbNullChar & "*.vbm" & vbNullChar

   Dim sfile As String
   Dim sDialogTitle As String
   Dim sFilter As String

   Select Case Index
   Case OFT_FILE
      sDialogTitle = DIALOG_FORM
      sFilter = FF_VBF_FILE
   Case OFT_TEMPLATE
      sDialogTitle = DIALOG_TEMPLATE
      sFilter = FF_VBM_FILE
   Case OFT_CMD_TEMPLATE
      sfile = CmdLineFileName
   End Select

   If Index <> OFT_CMD_TEMPLATE Then
      sfile = SelectFile(, Me.HWnd, sDialogTitle, , , , sFilter & FF_ALL_FILE)
      DoEvents
   End If
   If LenB(sfile) = 0 Or Not FileExists(sfile) Then
      Exit Sub '>---> Bottom
   End If

   Screen.MousePointer = vbHourglass
   Set m_VBForm = New VBForm
   With m_VBForm
      'Set reference for later use
      Set m_VBMenus = .Menus
      .FileName = sfile
      txtFormPath.Text = sfile
      .ParseModuleDeclare
      Set m_VBMenus = .Menus
      
      Select Case Index
      Case OFT_FILE
         lblFormName.Caption = "Form Name:"
         txtFormCaption.Text = .Caption
         txtFormTag.Text = .Tag
      Case OFT_TEMPLATE, OFT_CMD_TEMPLATE
         lblFormName.Caption = "Template:"
         .Name = GetFileName(.FileName, efpBaseName)
      End Select

      txtFormCaption.Visible = (Index = OFT_FILE)
      lblFormTag.Visible = (Index = OFT_FILE)
      txtFormTag.Visible = (Index = OFT_FILE)
      txtFormName.Text = .Name
   End With 'M_VBFORM

   'Load the parsed menus.
   LoadMenus

Ooops:
   If lstMenu.ListCount = 0 Then
      mnuEditClear_Click
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub LoadMenus()

   'Pase source text and load menus to the list box
   lstMenu.Clear
   
   Dim i As Long
   With m_VBMenus
      For i = 1 To .Count
         With .Item(i)
            lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc
         End With '.ITEM(I)
      Next i
   End With 'M_VBMENUS

   'Add a new item for inserting purpose
   If lstMenu.ListCount > 0 Then
      lstMenu.ListIndex = 0
   End If

End Sub

'=========== Edit Menu =============================
Private Sub mnuEditClear_Click()

   With m_VBMenus
      .Clear
      .Add
   End With 'M_VBMENUS
   With lstMenu
      .Clear
      .AddItem vbNullString
      .ListIndex = 0
   End With 'LSTMENU

End Sub

Private Sub mnuEditTemplate_Click()

   OpenFile OFT_TEMPLATE

End Sub

'============ View Menu ===========================
Private Sub mnuViewInfo_Click()

   Dim strMsg As String

   strMsg = "List:" & lstMenu.List(lstMenu.ListIndex) & vbCrLf & _
            "m_VBMenus:" & m_VBMenus(lstMenu.ListIndex + 1).Caption & "::" & m_VBMenus(lstMenu.ListIndex + 1).Key
   MsgBox strMsg

End Sub

Private Sub mnuViewWalk_Click()

   Debug.Print m_VBMenus.GetMenuText(False, False)

End Sub

'======= Menu Related Command Handler ==============
Private Sub cmdNext_Click()

   With lstMenu
      If .ListIndex < .ListCount - 1 Then
         .ListIndex = .ListIndex + 1
      Else 'NOT .LISTINDEX...
         .AddItem vbNullString
         With m_VBMenus.Add
            .Level = 1
         End With
         .ListIndex = .ListIndex + 1
      End If
   End With 'LSTMENU

End Sub

Private Sub cmdInsert_Click()

   With lstMenu
      .AddItem vbNullString, IIf(.ListIndex > -1, .ListIndex, 0)
      With m_VBMenus.Add
         .Level = 1
      End With
      .ListIndex = IIf(.ListIndex > -1, .ListIndex - 1, 0)
   End With 'LSTMENU

End Sub

Private Sub cmdDelete_Click()

   With lstMenu
      If .ListCount > 1 Then
         If .ListIndex > 0 Then
            .ListIndex = .ListIndex - 1
            .RemoveItem .ListIndex + 1
            m_VBMenus.Remove .ListIndex + 2
         Else 'NOT .LISTINDEX...
            .ListIndex = .ListIndex + 1
            .RemoveItem .ListIndex - 1
            m_VBMenus.Remove .ListIndex + 1
         End If
      Else 'NOT .LISTCOUNT...
         .List(0) = vbNullString
         m_VBMenus.Clear
         m_VBMenus.Add
         .ListIndex = 0
      End If
   End With 'LSTMENU

End Sub

Private Sub cmdPos_Click(Index As Integer)

   With m_SeledtedMenu
      Select Case Index
      Case 0 'Left
         If .Level > 1 Then
            .Level = .Level - 1
            lstMenu.List(lstMenu.ListIndex) = Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCut
         End If
      Case 1 'Right
         .Level = .Level + 1
         lstMenu.List(lstMenu.ListIndex) = Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCut
      Case 2 'Up
         With lstMenu
            If .ListIndex > 0 Then
               m_VBMenus.Swap .ListIndex, .ListIndex + 1
               Call ListMoveUp(lstMenu)
            End If
         End With 'LSTMENU
      Case 3 'Down
         With lstMenu
            If .ListIndex < .ListCount - 1 Then
               m_VBMenus.Swap .ListIndex + 1, .ListIndex + 2
               ListMoveDown lstMenu
            End If
         End With 'LSTMENU
      End Select
   End With 'M_SELEDTEDMENU

End Sub

'======= Menu Related Event Handler==============
Private Sub lstMenu_Click()

   Set m_SeledtedMenu = m_VBMenus(lstMenu.ListIndex + 1)

   With m_SeledtedMenu
      txtCaption.Text = .Caption
      txtName.Text = .Name
      chkChecked.Value = Abs(.Checked)
      chkEnabled.Value = Abs(.Enabled)
      chkVisible.Value = Abs(.Visible)
      txtIndex.Text = IIf(.Index = -1, vbNullString, .Index)
      chkWindowList.Value = .WindowList
      txtHelpContextID.Text = .HelpContextID
      cboShortCut.ListIndex = .ShortcutIndex
      cboNegotiationPosition.ListIndex = .NegotiatePosition
   End With 'M_SELEDTEDMENU

End Sub

Private Sub lstMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      PopupMenu mnuView
   End If
End Sub

Private Sub pUpdateListItem()
   With m_SeledtedMenu
      .Caption = txtCaption.Text
      .ShortcutIndex = cboShortCut.ListIndex
      lstMenu.List(lstMenu.ListIndex) = Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc
   End With 'M_SELEDTEDMENU
End Sub

Private Sub cboNegotiationPosition_Click()
   m_SeledtedMenu.NegotiatePosition = cboNegotiationPosition.ListIndex
End Sub
Private Sub chkChecked_Click()
   m_SeledtedMenu.Checked = chkChecked.Value
End Sub
Private Sub chkEnabled_Click()
   m_SeledtedMenu.Enabled = chkEnabled.Value
End Sub
Private Sub chkVisible_Click()
   m_SeledtedMenu.Visible = chkVisible.Value
End Sub
Private Sub chkWindowList_Click()
   m_SeledtedMenu.WindowList = chkWindowList.Value
End Sub
Private Sub cboShortCut_Click()
   pUpdateListItem
End Sub
Private Sub txtCaption_KeyUp(KeyCode As Integer, Shift As Integer)
   pUpdateListItem
End Sub
Private Sub txtCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   pUpdateListItem
End Sub

Private Sub txtIndex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_SeledtedMenu.Index = Val(txtIndex.Text)
End Sub

Private Sub txtIndex_KeyUp(KeyCode As Integer, Shift As Integer)
   m_SeledtedMenu.Index = Val(txtIndex.Text)
End Sub

Private Sub txtHelpContextID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_SeledtedMenu.HelpContextID = Val(txtHelpContextID.Text)
End Sub

Private Sub txtHelpContextID_KeyUp(KeyCode As Integer, Shift As Integer)
   m_SeledtedMenu.HelpContextID = Val(txtHelpContextID.Text)
End Sub

'======= Form Related ==============
Private Sub txtFormName_KeyUp(KeyCode As Integer, Shift As Integer)
   m_VBForm.Name = txtFormName.Text
End Sub

Private Sub txtFormName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   m_VBForm.Name = txtFormName.Name
End Sub

Private Sub txtFormPath_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txtName_Change()
   If lstMenu.ListIndex > -1 Then
      m_SeledtedMenu.Name = txtName.Text
   End If
End Sub


