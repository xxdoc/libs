VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "TrickVB6Installer"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   683
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   2
      Left            =   540
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   15
      Top             =   6180
      Visible         =   0   'False
      Width           =   9960
      Begin VB.CheckBox chkIncludeManifest 
         Caption         =   "Include manifest"
         Height          =   270
         Left            =   75
         TabIndex        =   17
         ToolTipText     =   "Use manifest"
         Top             =   2160
         Width           =   2865
      End
      Begin VB.TextBox txtManifest 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   150
         MaxLength       =   32767
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "frmMain.frx":000C
         Top             =   75
         Width           =   9660
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Index           =   1
      Left            =   630
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   9960
      Begin VB.CheckBox chkExeIgnoreError 
         Caption         =   "Ignore error"
         Height          =   300
         Left            =   75
         TabIndex        =   14
         ToolTipText     =   "Ignore errors"
         Top             =   2085
         Width           =   1470
      End
      Begin VB.TextBox txtExePath 
         Height          =   345
         Left            =   45
         MaxLength       =   260
         TabIndex        =   12
         ToolTipText     =   "Executable path"
         Top             =   1695
         Width           =   4665
      End
      Begin VB.TextBox txtParameters 
         Height          =   345
         Left            =   4800
         MaxLength       =   32768
         TabIndex        =   11
         ToolTipText     =   "Parameters"
         Top             =   1695
         Width           =   5040
      End
      Begin MSComctlLib.Toolbar tbrExecute 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "iglIcon"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "add"
               Object.ToolTipText     =   "Add execute path"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "remove"
               Object.ToolTipText     =   "Remove"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "up"
               Object.ToolTipText     =   "Move up"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "down"
               Object.ToolTipText     =   "Move down"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwExecute 
         Height          =   1215
         Left            =   30
         TabIndex        =   9
         Top             =   450
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Path"
            Object.Width           =   12241
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Parameters"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Flags"
            Object.Width           =   11642
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList iglBigIcon 
      Left            =   1125
      Top             =   9465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0282
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0994
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ECA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglBigIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New project"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save project"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "compile"
            Object.ToolTipText     =   "Compile"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   0
      Left            =   630
      ScaleHeight     =   171
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   664
      TabIndex        =   1
      Top             =   1020
      Width           =   9960
      Begin VB.CheckBox chkStrIgnoreError 
         Caption         =   "Ignore error"
         Height          =   300
         Left            =   3405
         TabIndex        =   13
         ToolTipText     =   "Ignore errors"
         Top             =   2160
         Width           =   1470
      End
      Begin VB.CheckBox chkMainExe 
         Caption         =   "Main executable"
         Height          =   300
         Left            =   1830
         TabIndex        =   10
         ToolTipText     =   "Startup after install"
         Top             =   2160
         Width           =   1470
      End
      Begin VB.CheckBox chkReplaceIfExist 
         Caption         =   "Replace if exists"
         Height          =   300
         Left            =   75
         TabIndex        =   5
         ToolTipText     =   "Replace destination file if exist"
         Top             =   2160
         Width           =   1650
      End
      Begin VB.TextBox txtStorageDst 
         Height          =   345
         Left            =   45
         MaxLength       =   260
         TabIndex        =   3
         ToolTipText     =   "Destination path"
         Top             =   1740
         Width           =   9825
      End
      Begin MSComctlLib.Toolbar tbrStorage 
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "iglIcon"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "add"
               Object.ToolTipText     =   "Add to storage"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "remove"
               Object.ToolTipText     =   "Remove from storage"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "up"
               Object.ToolTipText     =   "Move up"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "down"
               Object.ToolTipText     =   "Move down"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvwStorage 
         Height          =   1215
         Left            =   30
         TabIndex        =   4
         Top             =   450
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Destination path"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Source file"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Options"
            Object.Width           =   11642
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList iglIcon 
      Left            =   510
      Top             =   9465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":292E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3324
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3676
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabCategory 
      Height          =   9540
      Left            =   90
      TabIndex        =   0
      Top             =   570
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   16828
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Placement       =   2
      ImageList       =   "iglIcon"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Storage"
            Key             =   "storage"
            Object.ToolTipText     =   "Files storage."
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Execute"
            Key             =   "execute"
            Object.ToolTipText     =   "Execute after install."
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Manifest"
            Key             =   "manifest"
            Object.ToolTipText     =   "Include manifest xml-file"
            ImageVarType    =   2
            ImageIndex      =   7
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu nuSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuStorage 
         Caption         =   "&Storage"
         Begin VB.Menu mnuStorAdd 
            Caption         =   "&Add files"
            Shortcut        =   +{INSERT}
         End
         Begin VB.Menu mnuStorRemove 
            Caption         =   "&Remove files"
            Shortcut        =   +{DEL}
         End
      End
      Begin VB.Menu mnuExe 
         Caption         =   "&Execute"
         Begin VB.Menu mnuExeAdd 
            Caption         =   "&Add command"
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu mnuExeRemove 
            Caption         =   "&Remove command"
            Shortcut        =   ^R
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // frmMain.frm - main form of TrickVB6Installer application
' // © Krivous Anatoly Anatolevich (The trick), 2014

Option Explicit

Private Const CtlSpc = 5            ' // Space between controls

Private Project As clsProject       ' // Current project
Private Storage As clsStorage       ' // Current storage
Private Execute As clsExecute       ' // Current execute list
Private PrevTab As Long

Dim WithEvents FrmSubclass As clsTrickSubclass2
Attribute FrmSubclass.VB_VarHelpID = -1

' // +---------------------------------------------------------------------------------------+
' // |                                    Menu handlers                                      |
' // +---------------------------------------------------------------------------------------+

' // Menu "File"
Private Sub mnuNew_Click()
    
    ' // Suggest to save changes, if any
    If Not Project Is Nothing Then
        If Not QueryChanges Then Exit Sub
    End If
    
    ' // Create new project
    Set Project = New clsProject
    Set Storage = Project.Storage
    Set Execute = Project.Execute
    
    ' // Clear items
    lvwStorage.ListItems.Clear
    
    ' // Update manifest
    Project.Manifest = txtManifest.Text
    Project.UseManifest = chkIncludeManifest.Value = vbChecked
    Project.Modify = False
    
    ' // Set form caption
    Me.Caption = "Untitled - " & App.ProductName
    
End Sub

' // Manu "Open"
Private Sub mnuOpen_Click()
    Dim name    As Collection:  Dim i       As Long
    
    ' // Suggest to save changes, if any
    If Not QueryChanges Then Exit Sub
    
    ' // Show open dialog and obtain selected file name
    Set name = GetOpenFile(Me.hWnd, "Open project", "Trick setup project files" & vbNullChar & "*.ti" & vbNullChar & vbNullChar, False)
    
    If Not name Is Nothing Then
        
        ' // Try to load project from file
        If Not Project.Load(name(1) & "\" & name(2)) Then
            MsgBox "Error loading project", vbCritical
            Exit Sub
        End If
        
        ' // Update links
        Set Storage = Project.Storage
        Set Execute = Project.Execute
        
        ' // Update lists
        lvwStorage.ListItems.Clear
        lvwExecute.ListItems.Clear
        
        For i = 0 To Project.Storage.Count - 1
            lvwStorage.ListItems.Add
            UpdateStorageList i
        Next
        
        For i = 0 To Project.Execute.Count - 1
            lvwExecute.ListItems.Add
            UpdateExeList i
        Next
        
        Call lvwStorage_Click
        
        ' // Update manifest
        txtManifest.Text = Project.Manifest
        chkIncludeManifest.Value = Project.UseManifest And vbChecked
        
        ' // Update form caption
        Me.Caption = GetFileTitle(Project.FileName) & " - " & App.ProductName
        
    End If
    
End Sub

' // Manu "Save"
Private Sub mnuSave_Click()
    
    ' // Determine if project was saved before
    If Len(Project.FileName) Then
    
        ' // Save with old file name
        If Not Project.Save(Project.FileName) Then
            MsgBox "Error saving project", vbCritical
        Else
            ' // Update form caption
            Me.Caption = GetFileTitle(Project.FileName) & " - " & App.ProductName
        End If
        
    Else
    
        ' // Call "Save as..." handler
        Call mnuSaveAs_Click
        
    End If
    
End Sub

' // Menu "Save as..."
Private Sub mnuSaveAs_Click()
    Dim name As String
    
    ' // Show save dialog and get selected file name
    name = GetSaveFile(Me.hWnd, "Save project", "Trick setup project files" & vbNullChar & "*.ti" & vbNullChar & vbNullChar, "*.ti")
    
    If Len(name) Then
        
        ' // Try to save project
        If Not Project.Save(name) Then
            MsgBox "Error saving project", vbCritical
        Else
            ' // Update form caption
            Me.Caption = GetFileTitle(Project.FileName) & " - " & App.ProductName
        End If
        
    End If
    
End Sub

' // Menu "Compile..."
Private Sub mnuCompile_Click()
    Dim name As String
    
    ' // Show save dialog and get selected file name
    name = GetSaveFile(Me.hWnd, "Compilie project", "Executable file" & vbNullChar & "*.exe" & vbNullChar & "Binary file" & vbNullChar & "*.bin" & vbNullChar & vbNullChar, "*.exe")
    
    If Len(name) Then
    
        ' // Try to compile project
        If Not Project.Compile(name) Then
            MsgBox "Error compile", vbCritical
        End If
        
    End If
    
End Sub

' // Menu "Exit"
Private Sub mnuExit_Click()
    Unload Me
End Sub

' // Menu "Storage->Add files"
Private Sub mnuStorAdd_Click()
    Dim files As Collection
    
    ' // Show open dialog and get selected files
    Set files = GetOpenFile(Me.hWnd, "Add files", "All file types" & vbNullChar & "*.*" & vbNullChar & vbNullChar, True)
    
    ' // Add selected files to storage
    If Not files Is Nothing Then AddToStorage files
    
End Sub

' // Menu "Storage->Remove files"
Private Sub mnuStorRemove_Click()
    Dim index   As Long:        Dim itm     As ListItem
    
    ' // Check selected items
    If lvwStorage.SelectedItem Is Nothing Then Exit Sub
    
    ' // Show notification
    If MsgBox("Do you really want to delete?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
    
    ' // Go thru all items
    Do While index < lvwStorage.ListItems.Count
        
        ' // Get item
        Set itm = lvwStorage.ListItems(index + 1)
        
        ' // Remove selected item from storage
        If Not itm.Selected Then
            index = index + 1
        Else
            lvwStorage.ListItems.Remove (index + 1)
            Storage.Remove index
        End If
        
    Loop
    
    ' // Update checkboxes
    chkMainExe.Value = vbUnchecked
    chkReplaceIfExist.Value = vbUnchecked
    
End Sub

' // Menu "Execute->Add command"
Private Sub mnuExeAdd_Click()
    Dim cmd As String
    
    ' // Ask for command
    cmd = InputBox("Enter command")
    
    If Len(cmd) Then
        
        ' // Add to execution list
        Execute.Add cmd, vbNullString
        
        ' // Unselect
        If Not lvwExecute.SelectedItem Is Nothing Then lvwExecute.SelectedItem.Selected = False
        
        ' // Select new item
        lvwExecute.ListItems.Add.Selected = True
        
        ' // Update execution list
        UpdateExeList Execute.Count - 1
        
        ' // Click on execution list
        Call lvwExecute_Click
        
    End If
    
End Sub

' // Menu "Execute->Remove command"
Private Sub mnuExeRemove_Click()
    Dim index   As Long:        Dim itm     As ListItem
    
    ' // Check selected items
    If lvwExecute.SelectedItem Is Nothing Then Exit Sub
    
    ' // Show notification
    If MsgBox("Do you really want to delete?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
    
    ' // Go thru all items
    Do While index < lvwExecute.ListItems.Count
    
        Set itm = lvwExecute.ListItems(index + 1)
        
        ' // Remove selected item
        If Not itm.Selected Then
            index = index + 1
        Else
            lvwExecute.ListItems.Remove (index + 1)
            Execute.Remove index
        End If
        
    Loop
    
End Sub

' // Menu "About"
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

' // +---------------------------------------------------------------------------------------+
' // |                              Form and controls handlers                               |
' // +---------------------------------------------------------------------------------------+

' // From
Private Sub Form_Load()
    
    ' // Set icon
    SetWindowIcon Me.hWnd
    
    ' // Start subclassing
    Set FrmSubclass = New clsTrickSubclass2
    FrmSubclass.Hook Me.hWnd
    
    ' // Create new project
    Call mnuNew_Click
    
End Sub

' // Unloading
Private Sub Form_QueryUnload( _
            ByRef Cancel As Integer, _
            ByRef UnloadMode As Integer)
    ' // Suggest to save changes
    Cancel = Not QueryChanges
End Sub

' // Resizing
Private Sub Form_Resize()
    Dim i As Long
    
    If WindowState = vbMinimized Then Exit Sub
    
    ' // Move controls
    tabCategory.Move CtlSpc, tbrMain.Height + CtlSpc, ScaleWidth - CtlSpc * 2, ScaleHeight - tbrMain.Height - CtlSpc * 2
    i = tabCategory.SelectedItem.index - 1
    picTab(i).Move tabCategory.ClientLeft + CtlSpc, _
                   tabCategory.ClientTop + CtlSpc, _
                   tabCategory.ClientWidth - CtlSpc * 2, _
                   tabCategory.ClientHeight - CtlSpc * 2
                   
    Select Case i
    Case 0
    
        tbrStorage.Move CtlSpc, CtlSpc
        lvwStorage.Move CtlSpc, tbrStorage.Top + tbrStorage.Height + CtlSpc, _
                        picTab(i).ScaleWidth - CtlSpc * 2, picTab(i).ScaleHeight - tbrStorage.Height - CtlSpc * 5 - _
                        txtStorageDst.Height - chkReplaceIfExist.Height
        txtStorageDst.Move CtlSpc, lvwStorage.Top + lvwStorage.Height + CtlSpc, lvwStorage.Width
        chkReplaceIfExist.Top = txtStorageDst.Top + txtStorageDst.Height + CtlSpc
        chkMainExe.Top = chkReplaceIfExist.Top
        chkStrIgnoreError.Top = chkReplaceIfExist.Top
        
    Case 1
    
        tbrExecute.Move CtlSpc, CtlSpc
        lvwExecute.Move CtlSpc, tbrStorage.Top + tbrStorage.Height + CtlSpc, _
                        picTab(i).ScaleWidth - CtlSpc * 2, picTab(i).ScaleHeight - tbrStorage.Height - CtlSpc * 5 - _
                        txtParameters.Height - chkExeIgnoreError.Height
        txtExePath.Move CtlSpc, lvwExecute.Top + lvwExecute.Height + CtlSpc, lvwExecute.Width \ 2 - CtlSpc
        txtParameters.Move txtExePath.Left + txtExePath.Width + CtlSpc, txtExePath.Top, txtExePath.Width
        chkExeIgnoreError.Top = txtParameters.Top + txtParameters.Height + CtlSpc
        
    Case 2
    
        txtManifest.Move CtlSpc, CtlSpc, picTab(i).ScaleWidth - CtlSpc * 2, _
                         picTab(i).ScaleHeight - CtlSpc * 3 - chkIncludeManifest.Height
        chkIncludeManifest.Top = txtManifest.Top + txtManifest.Height + CtlSpc
        
    End Select
    
End Sub

' // Main toolbar
Private Sub tbrMain_ButtonClick( _
            ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "new":     Call mnuNew_Click       ' // Menu "New" handler
    Case "save":    Call mnuSave_Click      ' // Menu "Save" handler
    Case "open":    Call mnuOpen_Click      ' // Menu "Open" handler
    Case "compile": Call mnuCompile_Click   ' // Menu "Compile" handler
    End Select
    
End Sub

'// Tab selection
Private Sub tabCategory_BeforeClick( _
            ByRef Cancel As Integer)
    ' // Save previous tab
    PrevTab = tabCategory.SelectedItem.index - 1
End Sub

' // Tqab mouse up
Private Sub tabCategory_MouseUp( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
    
    ' // Hide previous tab container
    picTab(PrevTab).Visible = False
    ' // Show new tab container
    picTab(tabCategory.SelectedItem.index - 1).Visible = True
    ' // Resize
    Call Form_Resize
    
End Sub

' // Toolbar "storage"
Private Sub tbrStorage_ButtonClick( _
            ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "add":     Call mnuStorAdd_Click               ' // Menu "Storage->Add file" handler
    Case "remove":  Call mnuStorRemove_Click            ' // Menu "Storage->Remove file" handler
    Case "up"
        If lvwStorage.ListItems.Count = 0 Then Exit Sub ' // Move up file in storage
        MoveFile -1
    Case "down"
        If lvwStorage.ListItems.Count = 0 Then Exit Sub ' // Move down file in storage
        MoveFile 1
    End Select
    
End Sub

' // Toolbar "Execute"
Private Sub tbrExecute_ButtonClick( _
            ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "add":     Call mnuExeAdd_Click                ' // Menu "Execute->Add command" handler
    Case "remove":  Call mnuExeRemove_Click             ' // Menu "Execute->Remove command" handler
    Case "up"
        If lvwExecute.ListItems.Count = 0 Then Exit Sub ' // Move up command in execution list
        MoveCommand -1
    Case "down"
        If lvwExecute.ListItems.Count = 0 Then Exit Sub ' // Move down command in execution list
        MoveCommand 1
    End Select
    
End Sub

' // Click on list storage
Private Sub lvwStorage_Click()
    Dim ReplFlag    As Long:            Dim IgnoreFlag  As Long
    Dim MainFlag    As Long:            Dim Count       As Long
    Dim Item        As ListItem:        Dim fle         As clsStorageItem
    Dim Dst         As String
    
    ' // Go thru all item and get identical properties
    For Each Item In lvwStorage.ListItems
        
        If Item.Selected Then
            
            ' // Get storage item
            Set fle = Storage(Item.index - 1)
            
            ' // Check flags
            If fle.Flags And FF_REPLACEONEXISTS Then ReplFlag = ReplFlag + 1
            If fle.Flags And FF_IGNOREERROR Then IgnoreFlag = IgnoreFlag + 1
            
            ' // Check main executable
            If Storage.MainExecutable = Item.index - 1 Then MainFlag = 1
            
            ' // If it is same path
            If Count Then
                If StrComp(Dst, fle.DestinationPath, vbTextCompare) Then Dst = vbNullString
            Else: Dst = fle.DestinationPath
            End If
            
            Count = Count + 1
            
        End If
        
    Next
    
    ' // Update controls
    chkReplaceIfExist.Value = IIf(ReplFlag, IIf(ReplFlag = Count, vbChecked, vbGrayed), vbUnchecked)
    chkStrIgnoreError.Value = IIf(IgnoreFlag, IIf(IgnoreFlag = Count, vbChecked, vbGrayed), vbUnchecked)
    chkMainExe.Value = IIf(MainFlag, IIf(Count = 1, vbChecked, vbGrayed), vbUnchecked)
    txtStorageDst.Text = Dst
    
End Sub

' // Drop the files to storage list
Private Sub lvwStorage_OLEDragDrop( _
            ByRef Data As MSComctlLib.DataObject, _
            ByRef Effect As Long, _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
            
    If Data.GetFormat(vbCFFiles) = True Then
        Dim Col As Collection:  Dim fle As Variant
        
        Set Col = New Collection
        
        ' // Add dropped files to storage
        For Each fle In Data.files
        
            Col.Add GetFilePath(CStr(fle))
            Col.Add GetFileTitle(CStr(fle), True)
            AddToStorage Col
            Col.Remove 1: Col.Remove 1
            
        Next
        
    End If
    
End Sub

' // Right click on storage list
Private Sub lvwStorage_MouseUp( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)
            
    If Button = vbRightButton Then
        ' // Show storage popup menu
        PopupMenu mnuStorage
    End If
    
End Sub

' // Click on execution list
Private Sub lvwExecute_Click()
    Dim Count       As Long:            Dim Item        As ListItem
    Dim exe         As clsExecuteItem:  Dim fle         As String
    Dim cmd         As String:          Dim IgnoreFlag  As Long
    
    ' // Go thru all items and get identical properties
    For Each Item In lvwExecute.ListItems
    
        If Item.Selected Then
            
            ' // Get item
            Set exe = Execute(Item.index - 1)
            
            ' // Get identical flag
            If exe.Flags And EF_IGNOREERROR Then IgnoreFlag = IgnoreFlag + 1
            
            If Count Then
                ' // Get identical parameters
                If StrComp(cmd, exe.Parameters, vbTextCompare) Then cmd = vbNullString
                If StrComp(fle, exe.FileName, vbTextCompare) Then fle = vbNullString
                If Len(cmd) = 0 And Len(fle) = 0 Then Exit For
            Else
                fle = exe.FileName
                cmd = exe.Parameters
            End If
            
            Count = Count + 1
            
        End If
        
    Next
    
    ' // Update controls
    chkExeIgnoreError.Value = IIf(IgnoreFlag, IIf(IgnoreFlag = Count, vbChecked, vbGrayed), vbUnchecked)
    txtExePath.Text = fle
    txtParameters.Text = cmd
    
End Sub

' // Double click on execution list
Private Sub lvwExecute_DblClick()
    Dim cmd As String
    
    ' // Check selected item
    If lvwExecute.SelectedItem Is Nothing Then Exit Sub
    
    ' // Edit selected command
    cmd = InputBox("Edit command", , lvwExecute.SelectedItem.Text)
    
    If Len(cmd) Then
        
        ' // Update selected command
        With Execute.Item(lvwExecute.SelectedItem.index - 1)
            .FileName = cmd
        End With
        
        UpdateExeList lvwExecute.SelectedItem.index - 1
        
    End If
    
End Sub

' // Right click on execution list
Private Sub lvwExecute_MouseUp( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef x As Single, _
            ByRef y As Single)

    If Button = vbRightButton Then
        ' // Show execution popup menu
        PopupMenu mnuExe
    End If
    
End Sub

' // Execution caommand textbox
Private Sub txtExePath_KeyPress( _
            ByRef KeyAscii As Integer)
    
    ' // If field is empty then do nothing
    If Len(Trim(txtExePath.Text)) = 0 Then Exit Sub
    
    ' // If user press enter
    If KeyAscii = vbKeyReturn Then
        Dim itm As ListItem
        
        For Each itm In lvwExecute.ListItems
            
            ' // Update selected commands
            If itm.Selected Then
                Execute(itm.index - 1).FileName = txtExePath.Text
                UpdateExeList itm.index - 1
            End If
            
        Next
        
    End If
    
End Sub

' // Execution parameters textbox
Private Sub txtParameters_KeyPress( _
            ByRef KeyAscii As Integer)
    
    ' // If user press enter
    If KeyAscii = vbKeyReturn Then
        Dim itm As ListItem
        
        For Each itm In lvwExecute.ListItems
            
            ' // Update selected commands
            If itm.Selected Then
                Execute(itm.index - 1).Parameters = txtParameters.Text
                UpdateExeList itm.index - 1
            End If
            
        Next
        
    End If
    
End Sub

' // Manifest textbox
Private Sub txtManifest_Change()
    If Not Me.ActiveControl Is txtManifest Then Exit Sub
    Project.Manifest = txtManifest.Text
End Sub

' // Storage destiantion path textbox
Private Sub txtStorageDst_KeyPress( _
            ByRef KeyAscii As Integer)

    ' // If user press enter
    If KeyAscii = vbKeyReturn Then
        Dim itm As ListItem
        
        For Each itm In lvwStorage.ListItems
            
            ' // Update selected items
            If itm.Selected Then
                Storage(itm.index - 1).DestinationPath = txtStorageDst.Text
                UpdateStorageList itm.index - 1
            End If
            
        Next
        
    End If
    
End Sub

' // Main exe chaeckbox
Private Sub chkMainExe_Click()
    Dim i As Long
    
    If Not Me.ActiveControl Is chkMainExe Then Exit Sub
    
    ' // Get previous executable
    i = Storage.MainExecutable
    
    If chkMainExe.Value = vbUnchecked Then
        ' // Unset main executable
        Storage.MainExecutable = -1
        UpdateStorageList i
    Else
    
        If Not lvwStorage.SelectedItem Is Nothing Then
            ' // Set main executable
            Storage.MainExecutable = lvwStorage.SelectedItem.index - 1
            UpdateStorageList i
            UpdateStorageList lvwStorage.SelectedItem.index - 1
            
        End If
        
    End If
    
End Sub

' // "Replace if exists" checkbox
Private Sub chkReplaceIfExist_Click()
    Dim itm As ListItem:        Dim fle As clsStorageItem
    
    If Not Me.ActiveControl Is chkReplaceIfExist Then Exit Sub
    
    For Each itm In lvwStorage.ListItems
    
        ' // Update selected items
        If itm.Selected Then
        
            Set fle = Storage(itm.index - 1)
            
            If chkReplaceIfExist.Value = vbChecked Then
                fle.Flags = fle.Flags Or FF_REPLACEONEXISTS
            Else
                fle.Flags = fle.Flags And (Not FF_REPLACEONEXISTS)
            End If
            
            UpdateStorageList itm.index - 1
            
        End If
        
    Next
    
End Sub

' // "Ignore execution errors" checkbox
Private Sub chkStrIgnoreError_Click()
    Dim itm As ListItem:        Dim fle As clsStorageItem
    
    If Not Me.ActiveControl Is chkStrIgnoreError Then Exit Sub
    
    For Each itm In lvwStorage.ListItems
    
        ' // Update selected items
        If itm.Selected Then
        
            Set fle = Storage(itm.index - 1)
            
            If chkStrIgnoreError.Value = vbChecked Then
                fle.Flags = fle.Flags Or FF_IGNOREERROR
            Else
                fle.Flags = fle.Flags And (Not FF_IGNOREERROR)
            End If
            
            UpdateStorageList itm.index - 1
            
        End If
        
    Next
    
End Sub

' // "Ignore storage errors" checkbox
Private Sub chkExeIgnoreError_Click()
    Dim itm As ListItem:        Dim exe As clsExecuteItem
    
    If Not Me.ActiveControl Is chkExeIgnoreError Then Exit Sub
    
    For Each itm In lvwExecute.ListItems
        
        ' // Update selected items
        If itm.Selected Then
        
            Set exe = Execute(itm.index - 1)
            
            If chkExeIgnoreError.Value = vbChecked Then
                exe.Flags = exe.Flags Or EF_IGNOREERROR
            Else
                exe.Flags = exe.Flags And (Not EF_IGNOREERROR)
            End If
            
            UpdateExeList itm.index - 1
            
        End If
        
    Next
    
End Sub

' // "Manifest usage" checkbox
Private Sub chkIncludeManifest_Click()
    If Not Me.ActiveControl Is chkIncludeManifest Then Exit Sub
    Project.UseManifest = chkIncludeManifest.Value = vbChecked
End Sub

' // Subclassing proc
Private Sub FrmSubclass_WndProc( _
            ByVal hWnd As Long, _
            ByVal Msg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long, _
            ByRef Ret As Long, _
            ByRef DefCall As Boolean)
    
    Select Case Msg
    Case WM_GETMINMAXINFO   ' // Intercept minimum window size
    
        Dim inf As MINMAXINFO
        
        CopyMemory inf, ByVal lParam, Len(inf)
        inf.ptMinTrackSize.x = 350
        inf.ptMinTrackSize.y = 300
        CopyMemory ByVal lParam, inf, Len(inf)
        
        DefCall = False
        
    Case Else: DefCall = True
    End Select
    
End Sub

' // +---------------------------------------------------------------------------------------+
' // |                                   Helper functions                                    |
' // +---------------------------------------------------------------------------------------+

' // Query save changes (if any) and return True if project has been saved/ has no changes
Private Function QueryChanges() As Boolean

    If Project.Modify Then
    
        Select Case MsgBox("Save changes?", vbYesNoCancel Or vbQuestion)
        Case vbCancel: Exit Function
        Case vbYes
            Call mnuSave_Click
            If Project.Modify Then Exit Function
        End Select
        
    End If
    
    QueryChanges = True
    
End Function

' // Update data in execution list according clsExecute object
Private Sub UpdateExeList( _
            ByVal index As Long)
            
    Dim itm As ListItem:        Dim fle As clsExecuteItem
    
    If index < 0 Then Exit Sub
    
    Set itm = lvwExecute.ListItems(index + 1)
    Set fle = Execute(index)
    
    itm.Text = fle.FileName
    itm.SubItems(1) = fle.Parameters
    itm.SubItems(2) = vbNullString
    
    If fle.Flags And EF_IGNOREERROR Then itm.SubItems(2) = itm.SubItems(2) & " IGNORE"
    
End Sub

' // Move items in the execution list and clsExecute object
Private Sub MoveCommand( _
            ByVal Dir As Long)
    Dim i   As Long:        Dim s   As Long
    Dim d   As Long:        Dim c   As Long
    Dim itm As ListItem
    
    Dir = Sgn(Dir)
    
    If Dir > 0 Then
        d = 1: s = lvwExecute.ListItems.Count
    Else: s = 1: d = lvwExecute.ListItems.Count
    End If
    
    If lvwExecute.ListItems(s).Selected Then Exit Sub
    
    For c = s To d Step -Dir
    
        Set itm = lvwExecute.ListItems(c)
        
        If itm.Selected Then
        
            i = itm.index - 1
            If Execute.Swap(i, i + Dir) Then
            
                itm.Selected = False
                lvwExecute.ListItems(i + Dir + 1).Selected = True
                UpdateExeList i
                UpdateExeList i + Dir
                
            End If
            
        End If
        
    Next
    
End Sub

' // Add collection of files to storage. First item in collection is directory other items is file names
Private Sub AddToStorage( _
            ByRef files As Collection)
    Dim fle     As Variant: Dim sPath  As String
    Dim index   As Long:    Dim name    As String
    
    index = Storage.Count
    
    For Each fle In files
    
        If Len(sPath) Then
        
            name = sPath & "\" & fle
            Storage.Add name, "<app>", 0
            lvwStorage.ListItems.Add.Selected = True
            UpdateStorageList index
            index = index + 1
            
        Else: sPath = fle
        End If
        
    Next
    
    Call lvwStorage_Click
    
End Sub

' // Update data in storage list according clsStorage object
Private Sub UpdateStorageList( _
            ByVal index As Long)
    Dim itm As ListItem:        Dim fle As clsStorageItem
    
    If index < 0 Then Exit Sub
    
    Set itm = lvwStorage.ListItems(index + 1)
    Set fle = Storage(index)
    
    itm.Text = GetFileTitle(fle.FileName, True)
    itm.SubItems(1) = fle.DestinationPath
    itm.SubItems(2) = Project.ToAbsolute(fle.FileName, IIf(Len(Project.FileName), GetFilePath(Project.FileName), App.Path))
    itm.SubItems(3) = vbNullString
    
    If fle.Flags And FF_REPLACEONEXISTS Then itm.SubItems(3) = itm.SubItems(3) & " REPLACE"
    If fle.Flags And FF_IGNOREERROR Then itm.SubItems(3) = itm.SubItems(3) & " IGNORE"
    If index = Storage.MainExecutable Then itm.SubItems(3) = itm.SubItems(3) & " MAIN"
    
End Sub

' // Move items in the storage list and clsStorage object
Private Sub MoveFile( _
            ByVal Dir As Long)
    Dim i   As Long:        Dim s   As Long
    Dim d   As Long:        Dim c   As Long
    Dim itm As ListItem
    
    Dir = Sgn(Dir)
    
    If Dir > 0 Then
        d = 1: s = lvwStorage.ListItems.Count
    Else: s = 1: d = lvwStorage.ListItems.Count
    End If
    
    If lvwStorage.ListItems(s).Selected Then Exit Sub
    
    For c = s To d Step -Dir
    
        Set itm = lvwStorage.ListItems(c)
        
        If itm.Selected Then
        
            i = itm.index - 1
            
            If Storage.Swap(i, i + Dir) Then
            
                itm.Selected = False
                lvwStorage.ListItems(i + Dir + 1).Selected = True
                UpdateStorageList i
                UpdateStorageList i + Dir
                
            End If
            
        End If
        
    Next
    
End Sub
