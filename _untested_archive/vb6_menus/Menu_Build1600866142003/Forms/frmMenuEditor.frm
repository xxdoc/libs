VERSION 5.00
Begin VB.Form frmMenuEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Builder"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6600
   Icon            =   "frmMenuEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTag 
      Height          =   270
      Left            =   960
      TabIndex        =   27
      Top             =   780
      Width           =   4275
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "P&review"
      Height          =   315
      Left            =   5340
      TabIndex        =   26
      Top             =   780
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Height          =   300
      Left            =   5340
      TabIndex        =   14
      Top             =   60
      Width           =   1155
   End
   Begin VB.ListBox lstMenu 
      BackColor       =   &H00C0FFFF&
      DragIcon        =   "frmMenuEditor.frx":058A
      Height          =   2400
      ItemData        =   "frmMenuEditor.frx":0894
      Left            =   180
      List            =   "frmMenuEditor.frx":0896
      MouseIcon       =   "frmMenuEditor.frx":0898
      TabIndex        =   25
      Top             =   2700
      Width           =   6255
   End
   Begin VB.ComboBox cboNegotiationPosition 
      Height          =   300
      ItemData        =   "frmMenuEditor.frx":0BA2
      Left            =   4380
      List            =   "frmMenuEditor.frx":0BB2
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1500
      Width           =   2115
   End
   Begin VB.TextBox txtHelpContextID 
      Height          =   270
      Left            =   1440
      TabIndex        =   21
      Text            =   "0"
      Top             =   1500
      Width           =   855
   End
   Begin VB.CheckBox chkWindowList 
      Caption         =   "&Window List"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1860
      Width           =   1455
   End
   Begin VB.ComboBox cboShortCut 
      Height          =   300
      ItemData        =   "frmMenuEditor.frx":0BE1
      Left            =   4380
      List            =   "frmMenuEditor.frx":0CD5
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1140
      Width           =   2115
   End
   Begin VB.TextBox txtIndex 
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Top             =   1140
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   330
      Left            =   4140
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "&Visible"
      Height          =   255
      Left            =   3300
      TabIndex        =   5
      Top             =   1860
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "&Enabled"
      Height          =   255
      Left            =   1740
      TabIndex        =   4
      Top             =   1860
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkChecked 
      Caption         =   "&Checked"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cl&ose"
      Height          =   300
      Left            =   5340
      TabIndex        =   15
      Top             =   420
      Width           =   1155
   End
   Begin VB.TextBox txtName 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   405
      Width           =   4275
   End
   Begin VB.TextBox txtCaption 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   60
      Width           =   4275
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   3
      Left            =   1620
      Picture         =   "frmMenuEditor.frx":0FB4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   420
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   2
      Left            =   1140
      Picture         =   "frmMenuEditor.frx":12E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   420
   End
   Begin VB.CommandButton cmdPos 
      Height          =   330
      Index           =   1
      Left            =   660
      Picture         =   "frmMenuEditor.frx":1610
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
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
      Left            =   180
      Picture         =   "frmMenuEditor.frx":193E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   420
   End
   Begin VB.Label lblTag 
      AutoSize        =   -1  'True
      Caption         =   "&Tag:"
      Height          =   180
      Left            =   180
      TabIndex        =   28
      Top             =   810
      Width           =   390
   End
   Begin VB.Label lblNegotiatePosition 
      AutoSize        =   -1  'True
      Caption         =   "Neg&otiate Position:"
      Height          =   180
      Left            =   2640
      TabIndex        =   24
      Top             =   1560
      Width           =   1590
   End
   Begin VB.Label lblHelpcontextID 
      Caption         =   "&HelpcontextID:"
      Height          =   255
      Left            =   180
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6480
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6480
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label lblShortcut 
      AutoSize        =   -1  'True
      Caption         =   "&Shortcut:"
      Height          =   180
      Left            =   3480
      TabIndex        =   19
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label lblIndex 
      Caption         =   "Inde&x :"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Na&me:"
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   435
      Width           =   570
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Ca&ption :"
      Height          =   180
      Left            =   180
      TabIndex        =   16
      Top             =   90
      Width           =   765
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Template"
         HelpContextID   =   1234
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenTemplate 
         Caption         =   "&Open Template"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open &Form"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAsTemplate 
         Caption         =   "Save As Template"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFileSaveAsForm 
         Caption         =   "Save As Form"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      NegotiatePosition=   3  'Right
      Tag             =   "Edit Menu"
      Begin VB.Menu mnuEditPreviewMenuBar 
         Caption         =   "Preview as &Menubar"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditPreviewPopup 
         Caption         =   "Preview as Pop&up menu"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuEditSep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreviewItem 
         Caption         =   "&Preview item result"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuEditPreview 
         Caption         =   "Preview &result"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEditSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGenMenuBarCode 
         Caption         =   "Generate API Menu&Bar Code"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEditGenPopupCode 
         Caption         =   "&Generate API Popup Code"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditCopyParent 
         Caption         =   "Copy &with children"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCutParent 
         Caption         =   "C&ut with children"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditDeleteParent 
         Caption         =   "De&lete with children"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "Cle&ar"
      End
      Begin VB.Menu mnuEditSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditImportTemplate 
         Caption         =   "&Import Template"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditImportFile 
         Caption         =   "Import &Form"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditSep6 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "&Option"
      Begin VB.Menu mnuOptionBackup 
         Caption         =   "&Backup before save"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
' Created & released by KSY, 06/14/2003
'
#Const USE_ASSOCIATE = 1
#Const USE_SUBCLASS = 0

Public Enum eOpenFileType
   OFT_FILE
   OFT_TEMPLATE
   OFT_CMD_TEMPLATE
   OFT_IMPORT_FILE
   OFT_IMPORT_TEMPLATE
   OFT_DRAGDROP
End Enum

Public Enum eSaveFileType
   SFT_SAVE = 0
   SFT_SAVEAS_FORM = 1
   SFT_SAVEAS_TEMPLATE = 2
End Enum
   
Private m_strFileName As String 'Current File Name
Private m_VBMenus As New VBMenus 'Menus collection
Private m_SeledtedMenu As New VBMenu 'Current selected menu
Private m_bChanged As Boolean 'Indicator for changes


Private Sub DoCaption(bChanged As Boolean)
   'If changed, displays *.
   m_bChanged = bChanged
   If m_bChanged Then
      Me.Caption = "Menu Builder [* " & m_strFileName & "]"
   Else
      Me.Caption = "Menu Builder [" & m_strFileName & "]"
   End If
End Sub

Private Sub cmdPreview_Click()
   On Error Resume Next
   'Remove the last blank items.
   Call pRemoveLastBlankItems
   m_VBMenus.CreateShowAPIPopupPreview
   Call pAddLastBlankItem
   On Error GoTo 0
End Sub


Private Sub Form_Load()
   
   'Set tab stop for the ListBox to display caption and shortcut on the different columns.
   ListSetTabStop lstMenu.hWnd, 150
   
   'Process command line.
   'Read menus if an associated file is open or a file is dropped to the app.
   OpenFile OFT_CMD_TEMPLATE, VBA.Command$()
   
   'If there is no menu items read, prepare a new item
   If lstMenu.ListCount = 0 Then
      mnuEditClear_Click
   End If
   
   'Associate this app with "vbm" files.
   #If USE_ASSOCIATE Then
      With App
         AssociateFileType .Path, .EXEName, "Visual Basic Menu Template File", "vbm", False
      End With
   #End If
   
   'Start subclassing for file drag & drop, and ListBox dragging.
   'For debugging in IDE, set USE_SUBCLASS to 0 for safe.
   #If USE_SUBCLASS Then
      Call BeginSubclassing
   #End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Stop subclassing. Required.
   #If USE_SUBCLASS Then
      Call EndSubclassing
   #End If
End Sub


Private Sub mnuEdit_Click()
   'Update Edit menu items UIs.
   Call pUpdateEditMenu
End Sub

Private Sub pUpdateEditMenu()
   'Update Edit menu items UIs.
   With lstMenu
      If .ListIndex > -1 Then
         mnuEditCopy.Enabled = LenB(TrimCrLfTab(.List(.ListIndex))) > 0
      Else
         mnuEditCopy.Enabled = False
      End If
      mnuEditCut.Enabled = mnuEditCopy.Enabled
      mnuEditDelete.Enabled = .ListIndex > -1
   End With
   mnuEditCopyParent.Enabled = mnuEditCopy.Enabled And m_SeledtedMenu.IsParent
   mnuEditCutParent.Enabled = mnuEditCopyParent.Enabled
   mnuEditDeleteParent.Enabled = mnuEditDelete.Enabled And m_SeledtedMenu.IsParent
   'Ask to the VBMenus class whether there is any available menu on the clipboard.
   mnuEditPaste.Enabled = m_VBMenus.CanPaste
End Sub

Private Sub mnuEditCut_Click()
   'Copies the selected item to clipboard and delete from the collection.
   mnuEditCopy_Click
   mnuEditDelete_Click
End Sub

Private Sub mnuEditCutParent_Click()
   'Copies the selected item with children to clipboard and delete from the collection.
   mnuEditCopyParent_Click
   mnuEditDeleteParent_Click
End Sub

Private Sub mnuEditDelete_Click()
    cmdDelete_Click
End Sub

Private Sub mnuEditDeleteParent_Click()
   'Deletes a parent menu with its children.
   Dim nLast As Long, i As Long, nSelLevel As Long
   With lstMenu
      If .ListCount = 0 Or .ListIndex = -1 Then
         Exit Sub
      End If
      
      If .ListCount > 1 Then
         'Find the last child.
         nSelLevel = m_SeledtedMenu.Level
         For i = .ListIndex + 1 To .ListCount - 1
            If m_VBMenus(i + 1).Level <= nSelLevel Then
               nLast = i - 1
               Exit For
            End If
         Next
         
         'Removes the menu items from ListBox and Menus collection.
         'Process differently according that ListIndex is -1 or not.
         If .ListIndex > 0 Then
            .ListIndex = .ListIndex - 1
            For i = nLast To .ListIndex + 1 Step -1
               .RemoveItem i
               m_VBMenus.Remove i + 1
            Next i
            If .ListCount = 0 Then
               mnuEditClear_Click
            End If
         Else
            For i = nLast To .ListIndex Step -1
               .RemoveItem i
               m_VBMenus.Remove i + 1
            Next i
            If .ListCount > 0 Then
               .ListIndex = 0
            Else
               mnuEditClear_Click
            End If
         End If
      Else
         mnuEditClear_Click
      End If
   End With
   
   'Change caption & set popup(IsParent) poperty for menu items.
   'Popup property is needed for copy & paste, and moving, etc.
   DoCaption True
   m_VBMenus.SetPopupProperties
End Sub

Private Sub mnuEditGenMenuBarCode_Click()
   Dim frm As frmPreview
   Set frm = New frmPreview
   
   Call pRemoveLastBlankItems
   frm.txtResult.ColourEntireRTB m_VBMenus.CreateAPIDrawMenuBarCode(), True
   
   Call pAddLastBlankItem
   frm.Show vbModal, Me
   
   Set frm = Nothing
End Sub

Private Sub mnuEditGenPopupCode_Click()
   Dim frm As frmPreview
   Set frm = New frmPreview
   
   Call pRemoveLastBlankItems
   frm.txtResult.ColourEntireRTB m_VBMenus.CreateAPIPopupMenuCode(), True
   
   Call pAddLastBlankItem
   frm.Show vbModal, Me
   
   Set frm = Nothing
End Sub

Private Sub mnuEditImportFile_Click()
   'Import (i.e., add) menus to the list from a form or control file.
   OpenFile OFT_IMPORT_FILE
End Sub

Private Sub mnuEditImportTemplate_Click()
   'Import (i.e., add) menus to the list from a template file.
   'Template file is a  simplified form. Just has different extension (.vbm).
   OpenFile OFT_IMPORT_TEMPLATE
End Sub

Private Sub mnuEditPreviewMenuBar_Click()
   On Error Resume Next
   'Preview the whole menu text to created, without validation
   pRemoveLastBlankItems
   m_VBMenus.CreateShowAPIMenuBarPreview Me
   pAddLastBlankItem
   On Error GoTo 0
End Sub

Private Sub mnuEditPreviewPopup_Click()
   Call cmdPreview_Click
End Sub

Private Sub mnuFileExit_Click()
   'Currently I didn't added "Ask Save Changes." ^.^;
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   'Currently I didn't added "Ask Save Changes." ^.^;
   Unload Me
End Sub

'========== Save File/ Save Template ===================
Private Sub cmdOK_Click()
   SaveFile SFT_SAVE
End Sub
Private Sub mnuFileSave_Click()
   SaveFile SFT_SAVE
End Sub
Private Sub mnuFileSaveAsForm_Click()
   SaveFile SFT_SAVEAS_FORM
End Sub
Private Sub mnuFileSaveAsTemplate_Click()
   SaveFile SFT_SAVEAS_TEMPLATE
End Sub

Private Sub SaveFile(Index As eSaveFileType)

   On Error GoTo Ooops
   
   Dim sFilename As String
   Dim sDialogTitle As String
   Dim sFilter As String

   'Prepare dialoag title.
   Select Case Index
   Case SFT_SAVE
      'If no filename is not yet specified, just call save as.
      If LenB(m_strFileName) = 0 Then
         SaveFile SFT_SAVEAS_TEMPLATE
         Exit Sub
      End If
   Case SFT_SAVEAS_FORM
      sDialogTitle = DLGTITLE_SAVEAS_FORM
      sFilter = FF_VBF_FILE & FF_VBM_FILE
   Case SFT_SAVEAS_TEMPLATE
      sDialogTitle = DLGTITLE_SAVEAS_TEMPLATE
      sFilter = FF_VBM_FILE & FF_VBF_FILE
   Case Else
      Exit Sub
   End Select

   'If not just save, show save dialog.
   If Index <> SFT_SAVE Then
      sFilename = SelectSaveFile(Me.hWnd, sDialogTitle, sFilter & FF_ALL_FILE, GetFileName(m_strFileName, efpPath))
      If LenB(sFilename) = 0 Then
         Exit Sub '>---> Bottom
      End If
      m_strFileName = sFilename
   End If
   
   Dim bSuccess As Boolean
   Screen.MousePointer = vbHourglass
   
   'Remove the last blank items.
   Call pRemoveLastBlankItems
   
   'Validate menus.
   If Not CheckMenu() Then
      pAddLastBlankItem
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
      
  'Finally save the (changed) menu.
   bSuccess = m_VBMenus.SaveMenus(m_strFileName, False, mnuOptionBackup.Checked, True)
   Call DoCaption(Not bSuccess)
   pAddLastBlankItem
   Screen.MousePointer = vbDefault
   If Not bSuccess Then
      MsgBox "Error occurred during saving the menus."
   End If
Ooops:
End Sub

Private Sub pRemoveLastBlankItems()
   Dim i As Long
   
   'Remove the last blank items.
   With lstMenu
      For i = .ListCount - 1 To 0 Step -1
         If LenB(TrimCrLfTab(.List(i))) = 0 Then
            .RemoveItem i
            m_VBMenus.Remove i + 1
         Else
            Exit For
         End If
      Next i
   End With
End Sub
Private Sub pAddLastBlankItem()
   Dim i As Long
   
   'Remove the last blank items.
   With lstMenu
      If .ListCount = 0 Then
         .AddItem vbNullString
         m_VBMenus.Add.Level = 1
         .ListIndex = 0
      ElseIf LenB(TrimCrLfTab(.List(.ListCount))) <> 0 Then
         .AddItem vbNullString
         m_VBMenus.Add.Level = 1
      End If
      'If listindex is -1, select the last item to keep m_SeledtedMenu alive.
      If .ListIndex = -1 Then
         .ListIndex = .ListCount - 1
      End If
   End With
End Sub

Private Function CheckMenu() As Boolean
   'Validate menus.
   
   Dim ErrorItem As VBMenu
   Dim ErrorItemIndex As Long
   Dim ErrorDescription As String
   Dim lPrevListIndex As Long
  
   If m_VBMenus.Count = 0 Then
      MsgBox "There is no menu item created. Create menu items."
      mnuEditClear_Click
      Exit Function
   End If
   
   'Validate menus. If there is an error, show the error message.
   If m_VBMenus.Validate(ErrorItem, ErrorItemIndex, ErrorDescription) <> VBM_ERR_NONE Then
      lstMenu.ListIndex = ErrorItemIndex - 1
      Screen.MousePointer = vbDefault
      MsgBox ErrorDescription, vbCritical
      CheckMenu = False
   Else
      CheckMenu = True
   End If
   
End Function

'========== New/ Open File/ Open Template ===================
Private Sub mnuFileNew_Click()

   'Create new menu.
   'Currently I didn't added "Ask Save Changes." ^.^;
   
   m_strFileName = vbNullString
   DoCaption False

   'Clear menus collection. And set the level to 1. (First item's level should be 1.)
   With m_VBMenus
      .Clear
      .Add.Level = 1
   End With
   
   'Clear listbox.
   With lstMenu
      .Clear
      .AddItem vbNullString
      .ListIndex = 0
   End With
   
   DoCaption False
End Sub

Private Sub mnuFileOpen_Click()
   OpenFile OFT_FILE
End Sub

Private Sub mnuFileOpenTemplate_Click()
   OpenFile OFT_TEMPLATE
End Sub

Public Sub OpenFile(Index As eOpenFileType, Optional CmdLineFileName As String)
'Open/Import file

   Dim sFilename As String
   Dim sDialogTitle As String
   Dim sFilter As String
   
   On Error Resume Next

   'Prepare dialog title & file filter.
   Select Case Index
   Case OFT_FILE
      sDialogTitle = DIALOG_FORM
      sFilter = FF_VBF_FILE & FF_VBM_FILE
   Case OFT_IMPORT_FILE
      sDialogTitle = DIALOG_IMPORTFORM
      sFilter = FF_VBF_FILE & FF_VBM_FILE
   Case OFT_TEMPLATE
      sDialogTitle = DIALOG_TEMPLATE
      sFilter = FF_VBM_FILE & FF_VBF_FILE
   Case OFT_IMPORT_TEMPLATE
      sDialogTitle = DIALOG_IMPORTTEMPLATE
      sFilter = FF_VBM_FILE & FF_VBF_FILE
   Case OFT_CMD_TEMPLATE
      sFilename = GetLongFileName(CmdLineFileName)
   Case OFT_DRAGDROP
      sFilename = CmdLineFileName
   End Select

   'Show open file dialog, if this procedure is called from other reason
   'than command line or file drag & drop.
   If Index <> OFT_CMD_TEMPLATE And Index <> OFT_DRAGDROP Then
      sFilename = SelectFile(Me.hWnd, sDialogTitle, sFilter & FF_ALL_FILE, GetFileName(m_strFileName, efpPath))
      DoEvents
   End If
   If LenB(sFilename) = 0 Or Not FileExists(sFilename) Then
      Exit Sub '>---> Bottom
   End If

   Screen.MousePointer = vbHourglass
   
   Dim i As Long
   
   Select Case Index
   'if importing or file drag & drop, add items.
   Case OFT_IMPORT_FILE, OFT_IMPORT_TEMPLATE, OFT_DRAGDROP
      Dim oMenus As VBMenus
      Set oMenus = New VBMenus
      oMenus.ParseFile sFilename
      pAddMenus oMenus
      Set oMenus = Nothing
      
   'else, read items from the given file.
   Case Else
      m_strFileName = sFilename
      DoCaption False
      
      'Load the parsed menus.
      lstMenu.Clear
      With m_VBMenus
         On Error GoTo Bye
         .ParseFile sFilename
         For i = 1 To .Count
            With .Item(i)
               lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc
            End With
         Next i
      End With
      
      If lstMenu.ListCount > 0 Then
         lstMenu.ListIndex = 0
         m_VBMenus.Add.Level = 1
         lstMenu.AddItem vbNullString
      Else 'If there is no item, create an item by default.
         mnuEditClear_Click
      End If
   End Select

   Screen.MousePointer = vbDefault
   On Error GoTo 0
   Exit Sub
Bye:
   Screen.MousePointer = vbDefault
   'MsgBox "OpenFile::" & Err.Description
   On Error GoTo 0
End Sub

'=========== Edit Menu =============================
Private Sub mnuEditCopy_Click()
   'Copies the selected item to clipboard.
   'We use a custom clipboard format (ksCF_VBMENU) for menu text copy.
   With lstMenu
      If .ListIndex > -1 Then
         m_VBMenus.CopyToClipboard .ListIndex + 1, False
      End If
   End With
End Sub

Private Sub mnuEditCopyParent_Click()
   'Copies the selected item with its childeren to clipboard.
   With lstMenu
      If .ListIndex > -1 Then
         m_VBMenus.CopyToClipboard .ListIndex + 1, True
      End If
   End With
End Sub

Private Sub mnuEditPaste_Click()
   pAddMenus m_VBMenus.GetFromClipboard
End Sub

Private Sub pAddMenus(ByVal oMenus As VBMenus)
   'Add a new menu collection to the current collection.
   'and displays them on the listbox.
   Dim oMenu As VBMenu
   Dim i As Long, iBefore As Long
   
   With oMenus
      If .Count Then
         With lstMenu
            iBefore = IIf(.ListIndex > -1, .ListIndex, 0)
         End With
         For i = .Count To 1 Step -1
            Set oMenu = .Item(i)
            With oMenu
               m_VBMenus.Add oMenu, , iBefore + 1
               lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc, iBefore
            End With
         Next i
         DoCaption True
         lstMenu.ListIndex = iBefore
      End If
   End With
   
   'Set the popup properties
   m_VBMenus.SetPopupProperties
End Sub

Private Sub mnuEditClear_Click()
   'Create a new item to keep m_SeledtedMenu alive.
   With m_VBMenus
      .Clear
      .Add.Level = 1
      .SetPopupProperties 'Set the popup properties
   End With 'M_VBMENUS
   With lstMenu
      .Clear
      .AddItem vbNullString
      .ListIndex = 0
   End With 'LSTMENU

   'DoCaption True
End Sub

'============ Preview ===========================
Private Sub mnuEditPreview_Click()
   'Preview the whole menu text to created, without validation
   Dim frm As frmPreview
   Set frm = New frmPreview
   pRemoveLastBlankItems
   frm.txtResult.ColourEntireRTB m_VBMenus.GetMenuText(False, False), True
   pAddLastBlankItem
   frm.Show vbModal, Me
   Set frm = Nothing
End Sub

Private Sub mnuEditPreviewItem_Click()
   'Preview the text for an item to created, without validation.
   Dim frm As frmPreview
   Set frm = New frmPreview
   frm.txtResult.ColourEntireRTB m_SeledtedMenu.GetMenuText, True
   frm.Show vbModal, Me
   Set frm = Nothing
End Sub


'======= Menu Related Command Handler ==============
Private Sub cmdNext_Click()
   'Insert a new blank item to the last of the list.
   
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
   
   'Set the popup properties
   m_VBMenus.SetPopupProperties
   
   DoCaption True
End Sub

Private Sub cmdInsert_Click()
   'Insert a new blank item after the currently selected item.
   
   Dim iBefore As Long
   With lstMenu
      iBefore = IIf(.ListIndex > -1, .ListIndex, 0)
      .AddItem vbNullString, iBefore
      If iBefore + 2 > m_VBMenus.Count Then
         m_VBMenus.Add.Level = 1
      Else
        m_VBMenus.Add(, , iBefore + 1).Level = 1
      End If
      .ListIndex = iBefore
   End With 'LSTMENU
   
   'Set the popup properties
   m_VBMenus.SetPopupProperties
   
   DoCaption True
End Sub

Private Sub cmdDelete_Click()
   'Delete the selected item.
   
   With lstMenu
      If .ListCount = 0 Or .ListIndex = -1 Then
         Exit Sub
      End If
      
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
         m_VBMenus.Add.Level = 1
         .ListIndex = 0
      End If
   End With 'LSTMENU

   DoCaption True
   'Set the popup properties
   m_VBMenus.SetPopupProperties
End Sub

Private Sub cmdPos_Click(Index As Integer)

   On Error GoTo Bye:
   
   'Changes level and position of a menu item.
   With m_SeledtedMenu
      Select Case Index
      Case 0 'Left - Decrease the level. (Higher level means child.)
         If .Level > 1 Then
            .Level = .Level - 1
            lstMenu.List(lstMenu.ListIndex) = Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCut
            DoCaption True
         End If
      Case 1 'Right - Increase the level.
         .Level = .Level + 1
         lstMenu.List(lstMenu.ListIndex) = Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCut
         DoCaption True
      Case 2 'Up - Move the postion up by one.
         With lstMenu
            If .ListIndex > 0 Then
               ''Debug.Assert m_VBMenus.Count = .ListCount
               m_VBMenus.Swap .ListIndex, .ListIndex + 1
               Call ListMoveUp(lstMenu)
               'WalkMenus
               'Debug.Print m_VBMenus(.ListIndex) & "::" & m_VBMenus(.ListIndex + 1)
               DoCaption True
            End If
         End With 'LSTMENU
      Case 3 'Down - Move the position down by one.
         With lstMenu
            If .ListIndex < .ListCount - 1 Then
               m_VBMenus.Swap .ListIndex + 1, .ListIndex + 2
               ListMoveDown lstMenu
               'WalkMenus
               DoCaption True
            End If
         End With 'LSTMENU
       End Select
   End With 'M_SELEDTEDMENU

   'Set the popup properties
   m_VBMenus.SetPopupProperties
   Exit Sub
Bye:
   MsgBox "cmdPos(" & Index & ")::" & Err.Description
End Sub

Public Function MoveNodes(ByVal FromIdx As Long, ByVal ToIdx As Long) As String
     
   Dim i As Long
   Dim colClipMenus As New VBMenus
   Dim oMenu As VBMenu
   Dim nCount As Long
   Dim bCtrlKeyPressed As Boolean, bShiftKeyPressed As Boolean
   
   On Error GoTo Bye
   
   bCtrlKeyPressed = GetAsyncKeyState(VK_CONTROL) 'Copy indicator
   bShiftKeyPressed = GetAsyncKeyState(VK_SHIFT) 'Parent move/copy indicator
        
   'Validate count
   nCount = lstMenu.ListCount
   If nCount = 0 Or ToIdx < 0 Or ToIdx > nCount - 1 _
      Or FromIdx < 0 Or FromIdx > nCount - 1 Or FromIdx = ToIdx Then
      Exit Function
   End If

   'Set the popup properties
   m_VBMenus.SetPopupProperties
      
   'Crete temp collection
   Set colClipMenus = New VBMenus
   
   'Get a temporary reference.
   Set oMenu = m_VBMenus.Item(FromIdx + 1)
   
   'If parent, copy its children
   colClipMenus.Add oMenu
   
   If bShiftKeyPressed And oMenu.IsParent Then
      With m_VBMenus
         For i = FromIdx + 2 To nCount
            If .Item(i).Level > oMenu.Level Then
               colClipMenus.Add .Item(i)
            Else
               Exit For
            End If
         Next
      End With
   End If
   
   With colClipMenus
      nCount = .Count
      
      'If destination index is less than source index
      If ToIdx < FromIdx Then
         
         'First add items to the destination index.
         For i = nCount To 1 Step -1
            Set oMenu = .Item(i)
            With oMenu
               m_VBMenus.Items.Add oMenu, , ToIdx + 1
               lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc, ToIdx
            End With
         Next i
         
         'If just moving, remove the added items.
         If Not bCtrlKeyPressed Then
            For i = nCount To 1 Step -1
               m_VBMenus.Remove FromIdx + nCount + 1
               lstMenu.RemoveItem FromIdx + nCount
            Next i
         End If
         lstMenu.ListIndex = ToIdx
         
      Else
         If ToIdx + 1 - nCount > 0 Then
         
            If bCtrlKeyPressed Then 'Copy
               
               'Just add items.
               For i = nCount To 1 Step -1
                  Set oMenu = .Item(i)
                  With oMenu
                     m_VBMenus.Items.Add oMenu, , , ToIdx + 1 - nCount
                     lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc, ToIdx - nCount + 1
                  End With
               Next i
               
               lstMenu.ListIndex = IIf(ToIdx = 0, 0, ToIdx)

            Else 'Move
               
               'First remove items starting from the source index
               For i = nCount To 1 Step -1
                  m_VBMenus.Remove FromIdx + 1
                  lstMenu.RemoveItem FromIdx
               Next i
            
               'Then, add the tempoary copied items before the destination index
               'Some, calculated destination index is"ToIdx - nCount".
               For i = nCount To 1 Step -1
                  Set oMenu = .Item(i)
                  With oMenu
                     m_VBMenus.Items.Add oMenu, , ToIdx + 1 - nCount
                     lstMenu.AddItem Repeat((.Level - 1) * 4, ".") & .Caption & vbTab & .ShortCutDesc, ToIdx - nCount
                  End With
               Next i
               
               lstMenu.ListIndex = IIf(ToIdx = 0, 0, ToIdx - 1)
               
            End If
         Else
            lstMenu.ListIndex = FromIdx
         End If
      End If
   End With

   'Set the popup properties
   m_VBMenus.SetPopupProperties
   Exit Function
Bye:
   MsgBox "MoveNodes Error!: " & Err.Description
End Function


'Used for debugging purpose.
'Comment out this procedure when releasing.
Private Sub WalkMenus()
   Dim i As Long
   With m_VBMenus
   'Debug.Print "Count=" & .Count
   Debug.Print vbLine
   For i = 1 To .Count
      Debug.Print i & "::" & .Item(i).Caption
   Next i
   Debug.Print vbLine
   End With
End Sub

'======= Menu Related Event Handler==============
Private Sub lstMenu_Click()

   On Error GoTo Bye
   
   'Set the current menu item.
   Set m_SeledtedMenu = m_VBMenus(lstMenu.ListIndex + 1)
   
   'Changes the UI elements with the current menu item's properties.
   With m_SeledtedMenu
      txtCaption.Text = .Caption
      txtName.Text = .Name
      txtTag.Text = .Tag
      chkChecked.Value = Abs(.Checked)
      chkEnabled.Value = Abs(.Enabled)
      chkVisible.Value = Abs(.Visible)
      txtIndex.Text = IIf(.Index = -1, vbNullString, .Index)
      chkWindowList.Value = Abs(.WindowList)
      txtHelpContextID.Text = .HelpContextID
      cboShortCut.ListIndex = .ShortcutIndex
      cboNegotiationPosition.ListIndex = .NegotiatePosition
   End With 'M_SELEDTEDMENU
   Exit Sub
Bye:
   MsgBox "lstMenu_Click::" & Err.Description
End Sub

Private Sub lstMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Change the Edit menu UI elements and popup it.
   If Button = vbRightButton Then
      Call pUpdateEditMenu
      PopupMenu mnuEdit
   End If
End Sub

Private Sub pUpdateListItem()
   'Updates the properties of the the selected menu item.
   With m_SeledtedMenu
      .Caption = txtCaption.Text
      .Tag = txtTag.Text
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

Private Sub mnuOptionBackup_Click()
   mnuOptionBackup.Checked = Not mnuOptionBackup.Checked
End Sub

Private Sub txtCaption_KeyUp(KeyCode As Integer, Shift As Integer)
   pUpdateListItem
End Sub
Private Sub txtCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   pUpdateListItem
End Sub
Private Sub txtIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_SeledtedMenu.Index = Val(txtIndex.Text)
End Sub
Private Sub txtIndex_KeyUp(KeyCode As Integer, Shift As Integer)
   If LenB(txtIndex.Text) Then
      m_SeledtedMenu.Index = Val(txtIndex.Text)
   Else
      m_SeledtedMenu.Index = -1
   End If
End Sub
Private Sub txtHelpContextID_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_SeledtedMenu.HelpContextID = Val(txtHelpContextID.Text)
End Sub
Private Sub txtHelpContextID_KeyUp(KeyCode As Integer, Shift As Integer)
   m_SeledtedMenu.HelpContextID = Val(txtHelpContextID.Text)
End Sub

Private Sub txtName_Change()
   If lstMenu.ListIndex > -1 Then
      m_SeledtedMenu.Name = txtName.Text
   End If
End Sub

Private Sub txtTag_Change()
   If lstMenu.ListIndex > -1 Then
      m_SeledtedMenu.Tag = txtTag.Text
   End If
End Sub

Private Sub txtTag_KeyUp(KeyCode As Integer, Shift As Integer)
   If lstMenu.ListIndex > -1 Then
      m_SeledtedMenu.Tag = txtTag.Text
   End If
End Sub
