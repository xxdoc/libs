VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoComplete Demo"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextCustomSource 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox TextCustomSource 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox TextAll 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   4245
   End
   Begin VB.TextBox TextMRU 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   4245
   End
   Begin VB.TextBox TextFileSystem 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   4245
   End
   Begin VB.TextBox TextHistory 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   4245
   End
   Begin VB.Label LabelCustomSource 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom Source 2:"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Label LabelAll 
      Alignment       =   1  'Right Justify
      Caption         =   "All:"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label LabelCustomSource 
      Alignment       =   1  'Right Justify
      Caption         =   "Custom Source 1:"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label LabelMRU 
      Alignment       =   1  'Right Justify
      Caption         =   "MRU:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label LabelFileSystem 
      Alignment       =   1  'Right Justify
      Caption         =   "File System:"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label LabelHistory 
      Alignment       =   1  'Right Justify
      Caption         =   "History:"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1425
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HistoryAutoComplete As AutoComplete
Private FileSystemAutoComplete As AutoComplete
Attribute FileSystemAutoComplete.VB_VarHelpID = -1
Private MRUAutoComplete As AutoComplete
Private AllAutoComplete As AutoComplete
Attribute AllAutoComplete.VB_VarHelpID = -1
Private CustomAutoComplete(0 To 1) As AutoComplete

Private Sub Form_Load()
Set HistoryAutoComplete = New AutoComplete
Set FileSystemAutoComplete = New AutoComplete
Set MRUAutoComplete = New AutoComplete
Set AllAutoComplete = New AutoComplete
Set CustomAutoComplete(0) = New AutoComplete
Set CustomAutoComplete(1) = New AutoComplete
' Creation of the custom autocomplete sources.
Dim i As Long, StringArray() As String
ReDim StringArray(10) As String
For i = 0 To 10
    StringArray(i) = "Test A " & CStr(i)
Next i
CustomAutoComplete(0).CustomSource = StringArray()
For i = 0 To 10
    StringArray(i) = "Test B " & CStr(i)
Next i
CustomAutoComplete(1).CustomSource = StringArray()
' Set options and init the autocomplete objects.
HistoryAutoComplete.Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
HistoryAutoComplete.Init TextHistory.hWnd, AutoCompleteSourceHistory
FileSystemAutoComplete.Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
FileSystemAutoComplete.FileSystemOptions = AutoCompleteFileSystemOptionMyComputer + AutoCompleteFileSystemOptionFileSysDirs
FileSystemAutoComplete.Init TextFileSystem.hWnd, AutoCompleteSourceFileSystem
MRUAutoComplete.Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
MRUAutoComplete.Init TextMRU.hWnd, AutoCompleteSourceMRU
AllAutoComplete.Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
AllAutoComplete.FileSystemOptions = AutoCompleteFileSystemOptionMyComputer + AutoCompleteFileSystemOptionFileSysDirs
AllAutoComplete.Init TextAll.hWnd, AutoCompleteSourceAll
CustomAutoComplete(0).Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
CustomAutoComplete(0).Init TextCustomSource(0).hWnd, AutoCompleteSourceCustomSource
CustomAutoComplete(1).Options = AutoCompleteOptionSuggestAppend + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
CustomAutoComplete(1).Init TextCustomSource(1).hWnd, AutoCompleteSourceCustomSource
End Sub

' The option flag 'AutoCompleteOptionUseTab' is not working when using the intrinsic VB TextBox control.
' But it will work when using the TextBoxW control from my common controls project.
' -> http://www.vbforums.com/showthread.php?698563-CommonControls-(Replacement-of-the-MS-common-controls)
' In order to get it work it is necessary to handle the 'PreviewKeyDown' event of the TextBoxW control.
' The Tab key will be treated as a input key but only when the drop-down list of the autocomplete object is dropped down.
' Example: (replace "TestAutoComplete" to your valid autocomplete object)

'Private Sub TextBoxW1_PreviewKeyDown(ByVal KeyCode As Integer, IsInputKey As Boolean)
'If KeyCode = vbKeyTab Then IsInputKey = TestAutoComplete.DroppedDown()
'End Sub
