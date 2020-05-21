VERSION 5.00
Begin VB.UserControl PDFReader 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   KeyPreview      =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   7425
   ToolboxBitmap   =   "PDFReader.ctx":0000
   Begin VB.Frame TextPanel 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   210
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   5205
      Begin VB.TextBox PDFText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   30
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   330
         Width           =   3330
      End
      Begin VB.PictureBox ZoomPanel2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   510
         ScaleHeight     =   300
         ScaleWidth      =   4665
         TabIndex        =   16
         Top             =   0
         Width           =   4695
         Begin VB.Image btnClose 
            Height          =   240
            Left            =   4380
            Picture         =   "PDFReader.ctx":0312
            ToolTipText     =   "Close the text box"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image btnSave 
            Height          =   240
            Left            =   4080
            Picture         =   "PDFReader.ctx":069C
            ToolTipText     =   "Save the text"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image btnOCR 
            Height          =   240
            Left            =   3780
            Picture         =   "PDFReader.ctx":0A26
            ToolTipText     =   "Request an OCR on the PDF File"
            Top             =   30
            Width           =   240
         End
         Begin VB.Label lblText 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Extracting text ..."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   30
            Width           =   3600
         End
      End
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   355
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7395
      TabIndex        =   9
      Top             =   6525
      Width           =   7425
      Begin VB.PictureBox PagePanel 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2655
         ScaleHeight     =   300
         ScaleWidth      =   2145
         TabIndex        =   10
         Top             =   0
         Width           =   2175
         Begin VB.TextBox TxPageInView 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   855
            TabIndex        =   11
            Text            =   "0"
            Top             =   30
            Width           =   465
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   90
            Picture         =   "PDFReader.ctx":0DB0
            Top             =   30
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   1830
            Picture         =   "PDFReader.ctx":113A
            Top             =   30
            Width           =   240
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Page"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   14
            Top             =   30
            Width           =   420
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1335
            TabIndex        =   13
            Top             =   30
            Width           =   75
         End
         Begin VB.Label lblNbPages 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1410
            TabIndex        =   12
            Top             =   30
            UseMnemonic     =   0   'False
            Width           =   420
         End
      End
   End
   Begin VB.PictureBox ToolBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   355
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7395
      TabIndex        =   3
      Top             =   0
      Width           =   7425
      Begin VB.PictureBox ZoomPanel 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   324
         Left            =   2115
         ScaleHeight     =   300
         ScaleWidth      =   5070
         TabIndex        =   5
         Top             =   0
         Width           =   5100
         Begin VB.HScrollBar ZoomScroll 
            Height          =   195
            LargeChange     =   10
            Left            =   630
            Max             =   999
            Min             =   1
            TabIndex        =   6
            Top             =   60
            Value           =   1
            Width           =   2670
         End
         Begin VB.Image btnGetText 
            Height          =   240
            Left            =   4635
            Picture         =   "PDFReader.ctx":14C4
            ToolTipText     =   "Displays the textual content of the PDF file"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image btnFit 
            Height          =   240
            Left            =   4320
            Picture         =   "PDFReader.ctx":184E
            ToolTipText     =   "Adapts the display to the size of the control"
            Top             =   30
            Width           =   240
         End
         Begin VB.Image Btn100 
            Height          =   240
            Left            =   4005
            Picture         =   "PDFReader.ctx":1BD8
            ToolTipText     =   "Displays the document at its original size (Zoom 100%)"
            Top             =   30
            Width           =   240
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Zoom"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   30
            Width           =   420
         End
         Begin VB.Label lblZoomFactor 
            BackStyle       =   0  'Transparent
            Caption         =   "100.00%"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   7
            Top             =   30
            Width           =   555
         End
      End
      Begin VB.CommandButton PDFLoadButton 
         Height          =   250
         Left            =   90
         Picture         =   "PDFReader.ctx":1F62
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Open a PDF File"
         Top             =   45
         Width           =   250
      End
   End
   Begin VB.PictureBox PictureOCR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   90
      ScaleHeight     =   1455
      ScaleWidth      =   2370
      TabIndex        =   2
      Top             =   4980
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.VScrollBar PageScroll 
      Enabled         =   0   'False
      Height          =   3165
      Left            =   4815
      TabIndex        =   1
      Top             =   720
      Width           =   240
   End
   Begin VB.PictureBox PDFView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   360
      MousePointer    =   15  'Size All
      ScaleHeight     =   3465
      ScaleWidth      =   2280
      TabIndex        =   0
      Top             =   945
      Width           =   2310
   End
End
Attribute VB_Name = "PDFReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Déclarations d'événements:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Picture1,Picture1,-1,MouseMove
Attribute MouseMove.VB_Description = "Se produit lorsque l'utilisateur bouge la souris au dessus du contrôle "
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Se produit lors du premier affichage d'une feuille ou lorsque la taille d'un objet change."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Se produit lorsque l'utilisateur appuie sur un bouton de la souris puis le relâche au-dessus d'un objet."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Se produit lorsque l'utilisateur appuie sur un bouton de la souris et le relâche puis appuie à nouveau dessus avant de le relâcher au-dessus d'un objet."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Se produit lorsque l'utilisateur appuie sur une touche alors qu'un objet a le focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Se produit lorsque l'utilisateur appuie sur une touche ANSI puis la relâche ."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Se produit lorsque l'utilisateur relâche une touche alors qu'un objet a le focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Se produit lorsque l'utilisateur appuie sur le bouton de la souris alors qu'un objet a le focus."
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Se produit lorsque l'utilisateur relâche le bouton de la souris alors qu'un objet a le focus."
Event PDFLoaded(FileName As String, FilePath As String) 'MappingInfo=UserControl,UserControl,-1,PDFLoaded
Attribute PDFLoaded.VB_Description = "Se produit après le chargement réussi d'un document PDF"
Event PageChanged(PageViewed As Integer) 'MappingInfo=UserControl,UserControl,-1,PageChanged
Attribute PageChanged.VB_Description = "Se produit quand l'utilisateur ou l'application change la page du document en cours d'affichage"
'Valeurs de propriétés par défaut:
Const m_def_Align = 0
Const m_def_ToolTipText = ""
Const m_def_IsToolbarVisible = True
Const m_def_IsStatusBarVisible = True
Const m_def_IsPDFButtonVisible = True
Const ConvertPointToTwips As Double = 20.00376

'Variables de propriétés:
Dim m_Align As Integer
Dim m_ToolTipText As String
Dim m_IsToolbarVisible As Boolean
Dim m_IsStatusBarVisible As Boolean
Dim m_IsPDFButtonVisible As Boolean
Private ZoomFactor As Double
Private FactorFormat As Double

Private PageInView As Integer
Private PageHeight As Double
Private PageWith As Double
Private NoZoomChange As Boolean
Public Enum BorderTypes
    [None] = 0
    [Fixed Single] = 1
End Enum
Public Enum oPrinterOrientation
    [oAuto] = 0
    [oPortrait] = 1
    [oLandscape] = 2
End Enum
Private WithEvents cBrowser As CmnDialogEx   ' must use WithEvents keyword if events are desired
Attribute cBrowser.VB_VarHelpID = -1

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private PDFLib As Long

Public Enum ePDFPageRotation
  Rot_0
  Rot_90
  Rot_180
  Rot_270
End Enum

Private Declare Sub FPDF_InitLibrary Lib "PDFium" Alias "_FPDF_InitLibrary@0" ()
Private Declare Sub FPDF_DestroyLibary Lib "PDFium" Alias "_FPDF_DestroyLibrary@0" ()
Private Declare Function FPDF_LoadMemDocument Lib "PDFium" Alias "_FPDF_LoadMemDocument@12" (ByVal pData As Long, ByVal DataLen As Long, ByVal Password As String) As Long
Private Declare Sub FPDF_CloseDocument Lib "PDFium" Alias "_FPDF_CloseDocument@4" (ByVal hDoc As Long)
Private Declare Function FPDF_GetPageCount Lib "PDFium" Alias "_FPDF_GetPageCount@4" (ByVal hDoc As Long) As Long
Private Declare Function FPDF_LoadPage Lib "PDFium" Alias "_FPDF_LoadPage@8" (ByVal hDoc As Long, ByVal PageIdx As Long) As Long
Private Declare Sub FPDF_ClosePage Lib "PDFium" Alias "_FPDF_ClosePage@4" (ByVal hPage As Long)
Private Declare Sub FPDF_RenderPage Lib "PDFium" Alias "_FPDF_RenderPage@32" (ByVal hDC&, ByVal hPage&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal Rotation As ePDFPageRotation, ByVal Flags&)
Private Declare Function FPDFBitmap_Create Lib "PDFium" Alias "_FPDFBitmap_Create@12" (ByVal dx As Long, ByVal dy As Long, ByVal Alpha As Long) As Long
Private Declare Sub FPDF_RenderPageBitmap Lib "PDFium" Alias "_FPDF_RenderPageBitmap@32" (ByVal hBM&, ByVal hPage&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal Rotation As ePDFPageRotation, ByVal Flags&)
Private Declare Function FPDFBitmap_GetBuffer Lib "PDFium" Alias "_FPDFBitmap_GetBuffer@4" (ByVal hBM As Long) As Long
Private Declare Sub FPDFBitmap_Destroy Lib "PDFium" Alias "_FPDFBitmap_Destroy@4" (ByVal hBM As Long)
Private Declare Function FPDF_GetPageWidth Lib "PDFium" Alias "_FPDF_GetPageWidth@4" (ByVal hPage As Long) As Double
Private Declare Function FPDF_GetPageHeight Lib "PDFium" Alias "_FPDF_GetPageHeight@4" (ByVal hPage As Long) As Double
Private Declare Function FPDFLink_CountWebLinks Lib "PDFium" Alias "_FPDFLink_CountWebLinks@4" (ByVal hPage As Long) As Double
Private Declare Function FPDFText_LoadPage Lib "PDFium" Alias "_FPDFText_LoadPage@4" (ByVal hPage As Long) As Long
Private Declare Function FPDFPage_CountObjects Lib "PDFium" Alias "_FPDFPage_CountObjects@4" (ByVal hPage As Long) As Long
Private Declare Function FPDFPage_GetObject Lib "PDFium" Alias "_FPDFPage_GetObject@8" (ByVal hPage As Long, ByVal ObjIdx As Long) As Long
Private Declare Sub FPDFText_ClosePage Lib "PDFium" Alias "_FPDFText_ClosePage@4" (ByVal hPageText As Long)
Private Declare Function FPDFText_CountChars Lib "PDFium" Alias "_FPDFText_CountChars@4" (ByVal hPageText As Long) As Long
Private Declare Function FPDFText_GetText Lib "PDFium" Alias "_FPDFText_GetText@16" (ByVal hPageText As Long, ByVal StartIdx As Long, ByVal CharsCount As Long, ByVal pStrW As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function OpenProcess Lib "kernel32" _
       (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
       (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const CP_UTF8 As Long = 65001
Private Const CP_ACP As Long = 0

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long


Private mPageText As Long, ToShow As Boolean
Private PDFLoadedFileName As String
Private OCRPath As String
Private Language As String
Private hDoc As Long, Pages() As Long, Content() As Byte 'as long as a document is open, the Content-buffer should not be touched or changed
Private NbPage As Integer
Private pdfX, pdfY, pdfH, pdfW, MouseX, MouseY, pdfPosX, pdfPosY, pdfInit As Boolean, ToMove As Boolean, Xmem, Ymem

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Type tPOINT
    X As Long
    Y As Long
End Type
Private MouseCoords As tPOINT

Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPOINT) As Long

Private PDFView_hWnd As Long
Private Const WM_MOUSEWHEEL         As Long = &H20A ' window message for mouse wheel
Private MouseWheelUp As Boolean     ' true if mouse wheel up, false if down
Private OldProc1 As Long             ' Holds the old TWndProc for form 1
Private oldproc2 As Long             ' Holds the old TWndProc for form 2
Dim WithEvents FormHook  As clsTrickSubclass2
Attribute FormHook.VB_VarHelpID = -1



Private Function TWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Direction As Integer, ConnectedTo As Long, Shift As Integer

    If wMsg = WM_MOUSEWHEEL Then ' Have we got a mouse wheel message
        If NbPage > 0 Then 'A PDF is loaded
            If wParam > 0 Then MouseWheelUp = True Else MouseWheelUp = False
            Select Case MouseWheelUp
                Case True ' mouse up value is found
                    Direction = -1
                Case False ' mouse value down is found
                    Direction = 1
            End Select
            GetCursorPos MouseCoords 'get current mouse location
            ConnectedTo = WindowFromPoint(MouseCoords.X, MouseCoords.Y)
            Shift = wParam And &HFFFF&
            WheelScroll ConnectedTo, Direction, Shift
        End If
    End If

  TWndProc = CallWindowProc(OldProc1, hWnd, wMsg, wParam, lParam)

End Function

Public Function GetPDFText() As String
Attribute GetPDFText.VB_Description = "Returns a string containing the textual data present within the PDF document"
    Dim F As Long, tmpText As String, CharsCount As Long, resultText As String
    If NbPage > 0 Then
        For F = 0 To NbPage - 1
            mPageText = FPDFText_LoadPage(Pages(F))
            CharsCount = FPDFText_CountChars(mPageText)
            tmpText = Space$(CharsCount)
            If Len(tmpText) Then FPDFText_GetText mPageText, 0, Len(tmpText) + 1, StrPtr(tmpText)
            resultText = resultText & tmpText & vbCrLf
            FPDFText_ClosePage mPageText
        Next F
    End If
    GetPDFText = resultText
End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the background color used to display the text and graphics of an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If BackColor <> New_BackColor Then
        UserControl.BackColor() = New_BackColor
        Me.BackColor = New_BackColor
        ToolBar.BackColor = New_BackColor
        StatusBar.BackColor = New_BackColor
        PropertyChanged "BackColor"
    End If
End Property

Public Property Get Zoom() As Double
Attribute Zoom.VB_Description = "Returns or sets the value of the Zoom. Can only be edited when a PDF file is loaded in the control."
    Zoom = ZoomFactor
    'ForeColor = UserControl.ForeColor
End Property

Public Property Let Zoom(ByVal New_zoom As Double)
    If NbPage > 0 Then
        ZoomFactor = New_zoom
        If ZoomFactor < 0.1 Then ZoomFactor = 0.1
        ZoomScroll.Value = IIf(ZoomFactor < ZoomScroll.Max, ZoomFactor, ZoomScroll.Max)
        ZoomScroll_Change
    Else
        ZoomFactor = 100
    End If
    PropertyChanged "Zoom"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns or sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'
Public Property Get TesseractPath() As String
    TesseractPath = OCRPath
End Property

Public Property Let TesseractPath(ByVal New_Path As String)
    OCRPath = New_Path
    PropertyChanged "TesseractPath"
End Property
'
Public Property Get OCRLanguage() As String
    OCRLanguage = Language
End Property

Public Property Let OCRLanguage(ByVal New_Language As String)
    Language = New_Language
    PropertyChanged "OCRLanguage"
End Property


Public Property Get BorderStyle() As BorderTypes
Attribute BorderStyle.VB_Description = "Returns or sets the border style of an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderTypes)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


Public Sub Refresh()
Attribute Refresh.VB_Description = "Force a new complete display of an object."
    UserControl.Refresh
End Sub



Private Sub Btn100_Click()
    On Error Resume Next
    ZoomScroll.Value = 100
End Sub

Private Sub btnClose_Click()
    TextPanel.Visible = False
    PDFText.Text = ""
End Sub

Private Sub btnFit_Click()
    UserControl_Resize
    UserControl_Paint
End Sub

Private Sub ButtonOrdoConcept1_Click()

End Sub



Private Sub btnGetText_Click()
    lblText.Caption = "Extracting text ..."
    TextPanel.Visible = True
    DoEvents
    PDFText.Text = GetPDFText
    lblText.Caption = "Text content of PDF file"
    Dim strTMP$
    strTMP = Replace(PDFText.Text, " ", "")
    strTMP = Replace(strTMP, Chr$(13), "")
    strTMP = Replace(strTMP, Chr$(10), "")
    If Trim$(strTMP) = "" And OCRPath > "" Then
        If MsgBox("Your file does not appear to contain clear text." & vbNewLine & "Do you want to try a Character Recognition (OCR) ?", vbQuestion Or vbYesNo, "No text found...") = vbYes Then
            MakeOCR 'PDFLoadedFileName
        End If
    End If
'    Debug.Print """" & Trim$(PDFText.Text) & """"
End Sub



Private Sub btnOCR_Click()
    If OCRPath > "" Then
        If MsgBox("Do you want to try Character Recognition (OCR) on this file ?", vbQuestion Or vbYesNo, "Characters recognition...") = vbYes Then
            MakeOCR 'OnFile PDFLoadedFileName
        End If
    Else
        MsgBox "The character recognition (OCR) function is not available on your machine.", vbExclamation, "Error"
    End If
End Sub

Private Sub btnSave_Click()
    Dim bReturn As Boolean, eEvents As EventTypeEnum, FileName As String, File As Integer
    If PDFText.Text = "" Then
        Beep
        Exit Sub
    End If
    With cBrowser
        .Clear
        .DefaultExt = "txt"
        .Tag = "SaveDialog" ' flag used below
        .DialogTitle = "Save text..."
        .FlagsDialog = DLG__BaseSaveDialogFlags
        .Filter = "Text Files|*.txt" ' vbNullChar ' flag to hide the filter
        bReturn = .ShowSave(Me.hWnd, eEvents)
        
    End With
    If bReturn Then
        FileName = cBrowser.FileName
        File = FreeFile
        Open FileName For Output As #File
            Print #File, PDFText.Text
        Close #File
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub FormHook_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
    Dim Direction As Integer, ConnectedTo As Long, Shift As Integer
    DefCall = False
    Select Case Msg
        Case WM_MOUSEWHEEL
            If NbPage > 0 Then 'A PDF is loaded
                If wParam > 0 Then MouseWheelUp = True Else MouseWheelUp = False
                Select Case MouseWheelUp
                    Case True ' mouse up value is found
                        Direction = -1
                    Case False ' mouse value down is found
                        Direction = 1
                End Select
                GetCursorPos MouseCoords 'get current mouse location
                ConnectedTo = WindowFromPoint(MouseCoords.X, MouseCoords.Y)
                Shift = wParam And &HFFFF&
                WheelScroll ConnectedTo, Direction, Shift
            End If
            
        Case Else
            DefCall = True
    End Select
End Sub

Private Sub Image1_Click() ' Previous page
    If NbPage = 0 Then Exit Sub
    PageInView = PageInView - 1
    If PageInView < 1 Then PageInView = 1
    TxPageInView.Text = PageInView
    PDFView.Cls
    If ToShow Then PDFView.top = m_IsToolbarVisible * ToolBar.Height * -1: DoEvents
    If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
    If ToShow Then RaiseEvent PageChanged(PageInView)
    PageScroll.Value = PageInView
End Sub

Private Sub Image2_Click() ' Next page
    If NbPage = 0 Then Exit Sub
    PageInView = PageInView + 1
    If PageInView > NbPage Then PageInView = NbPage
    TxPageInView.Text = PageInView
    PDFView.Cls
    If ToShow Then PDFView.top = m_IsToolbarVisible * ToolBar.Height * -1: DoEvents
    If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
    If ToShow Then RaiseEvent PageChanged(PageInView)
    PageScroll.Value = PageInView
End Sub

Private Sub WheelScroll(ByVal ConnectedTo As Long, ByVal Direction As Long, ByVal Shift As Integer)
Attribute WheelScroll.VB_Description = "Not used"
    If NbPage = 0 Or PDFLoadedFileName = "" Then Exit Sub
    On Error Resume Next
    Debug.Print Shift
    Select Case ConnectedTo
        Case PDFView.hWnd
            If Shift = 0 Then
                On Error Resume Next
                PageScroll.Value = PageScroll.Value + Direction
            ElseIf Shift = 4 Then
                PDFView.top = PDFView.top - (Direction * 500)
                UserControl_Paint
            ElseIf Shift = 6 Or Shift = 2 Then
                PDFView.left = PDFView.left - (Direction * 500)
                UserControl_Paint
            ElseIf Shift = 8 Then
                ZoomScroll.Value = ZoomScroll.Value - (Direction * 10)
            Else
                '
            End If
        Case Else
            On Error Resume Next
            If Shift = 8 Then
                ZoomScroll.Value = ZoomScroll.Value - (Direction * 10)
            Else
                PageScroll.Value = PageScroll.Value + Direction
            End If
    End Select
End Sub


Private Sub PageScroll_Change()
    Dim NewPage As Integer
    NewPage = PageScroll.Value
    If NewPage = 0 Then NewPage = 1
    If NewPage > NbPage Then NewPage = NbPage
    If PageInView <> NewPage Then
        PageInView = NewPage
        PDFView.Cls
        If ToShow Then PDFView.top = m_IsToolbarVisible * ToolBar.Height * -1: DoEvents
        If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
        If ToShow Then RaiseEvent PageChanged(PageInView)
    End If
    TxPageInView.Text = CStr(PageInView)
End Sub

Private Sub PDFLoadButton_Click()
    SelectPDFFile
End Sub

Private Sub PDFView_DblClick()
    UserControl_Resize
    UserControl_Paint
End Sub

Private Sub PDFView_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print KeyCode
    If NbPage > 0 And ToShow = True Then
        Select Case KeyCode
            Case 33
                Image1_Click
            Case 34
                Image2_Click
            Case 37
               PDFView.left = PDFView.left + 100
            Case 38
               PDFView.top = PDFView.top + 100
            Case 39
               PDFView.left = PDFView.left - 100
            Case 40
               PDFView.top = PDFView.top - 100
            Case Else
                UserControl_KeyDown KeyCode, Shift
        End Select
    Else
        UserControl_KeyDown KeyCode, Shift
    End If
End Sub

Private Sub PDFView_KeyPress(KeyAscii As Integer)
    UserControl_KeyPress KeyAscii
End Sub

Private Sub PDFView_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 To 40
            UserControl_Paint
        Case 33 To 34
        Case Else
            UserControl_KeyUp KeyCode, Shift
    End Select
End Sub

Private Sub PDFView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then '
        UserControl.AutoRedraw = True
        PDFView.AutoRedraw = True
        MouseX = X
        MouseY = Y
        pdfPosX = PDFView.left
        pdfPosY = PDFView.top
        ToMove = True
        Xmem = pdfPosX
        Ymem = pdfPosY
    Else
        UserControl_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub PDFView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xpos, Ypos
    If Button = 1 Then '
        Xpos = X - MouseX
        Ypos = Y - MouseY
        If Abs(Xmem - Xpos) > 60 Or Abs(Ymem - Ypos) > 60 Then
            Xmem = Xpos
            Ymem = Ypos
            PDFView.Move pdfPosX + Xpos, pdfPosY + Ypos, PDFView.Width, PDFView.Height
        End If
    Else
        UserControl_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub PDFView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToMove = False
    PDFView.AutoRedraw = False
    UserControl.AutoRedraw = False
    UserControl_Paint
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub TxPageInView_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
        Case 13
            KeyAscii = 0
            PDFView.SetFocus
        Case Else
            Beep
    End Select
End Sub
Private Sub TxPageInView_LostFocus()
    Dim NewPage As Integer
    NewPage = Val(TxPageInView.Text)
    If NewPage = 0 Then NewPage = 1
    If NewPage > NbPage Then NewPage = NbPage
    If PageInView <> NewPage Then
        PageInView = NewPage
        PDFView.Cls
        If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
        If ToShow Then RaiseEvent PageChanged(PageInView)
        PageScroll.Value = PageInView
    End If
    TxPageInView.Text = CStr(PageInView)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    PDFView.SetFocus
End Sub
Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
    Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
    Set Extender.Container = Value
End Property

Private Sub UserControl_Initialize()
    PDFView_hWnd = PDFView.hWnd
    '*** It is possible to load the library upon initialization, or only when loading a PDF file for the first time ****
    'The second option decreases the initialization time of the control and its memory footprint but makes the first loading of a PDF longer
'    PDFLib = LoadLibraryW(StrPtr(App.Path & "\PDFium.dll"))
'    FPDF_InitLibrary
    '---
    Set cBrowser = New CmnDialogEx
    FactorFormat = 29.7 / 21 'A4 size
    Language = "eng" 'Default OCR language
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 67 Then
        CopyPageToClipboard
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Do nothing for the moment
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Do nothing for the moment
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Cls()
Attribute Cls.VB_Description = "Clears and closes the PDF document being displayed."
    PDFView.Cls
    CloseDocument
End Sub

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a descriptor (from Microsoft Windows) to the context of the object's device."
    hDC = PDFView.hDC
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a descriptor (from Microsoft Windows) to an object window."
    hWnd = UserControl.hWnd
End Property
'

Private Sub UserControl_InitProperties() 'Initialize properties for user control
    m_IsToolbarVisible = True
    m_IsStatusBarVisible = True
    m_IsPDFButtonVisible = True
    UserControl_Resize
End Sub

Private Sub UserControl_Paint()
    If NbPage > 0 And PageInView > 0 And ToMove = False Then PDFView.Cls: If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Align = PropBag.ReadProperty("Align", m_def_Align)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_IsToolbarVisible = PropBag.ReadProperty("IsToolbarVisible", m_def_IsToolbarVisible)
    m_IsStatusBarVisible = PropBag.ReadProperty("IsStatusBarVisible", m_def_IsStatusBarVisible)
    m_IsPDFButtonVisible = PropBag.ReadProperty("IsPDFButtonVisible", m_def_IsPDFButtonVisible)
    OCRPath = PropBag.ReadProperty("TesseractPath", "")
    Language = PropBag.ReadProperty("OCRLanguage", "")
    If NbPage > 0 Then
        ZoomFactor = PropBag.ReadProperty("Zoom", 100)
        If ZoomFactor < 0.1 Then ZoomFactor = 0.1
        ZoomScroll.Value = IIf(ZoomFactor < ZoomScroll.Max, ZoomFactor, ZoomScroll.Max)
    Else
        ZoomFactor = 100
    End If
    PDFView.ToolTipText = m_ToolTipText
    PDFLoadButton.Visible = m_IsPDFButtonVisible
    StatusBar.Visible = m_IsStatusBarVisible
    ToolBar.Visible = m_IsToolbarVisible
    ToolBar.BackColor = UserControl.BackColor
    StatusBar.FillColor = UserControl.BackColor
End Sub

Private Sub UserControl_Terminate()
    CloseDocument
    'FPDF_DestroyLibary
    'FreeLibrary PDFLib
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Zoom", ZoomFactor, 100)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Align", m_Align, m_def_Align)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("IsToolbarVisible", m_IsToolbarVisible, m_def_IsToolbarVisible)
    Call PropBag.WriteProperty("IsStatusBarVisible", m_IsStatusBarVisible, m_def_IsStatusBarVisible)
    Call PropBag.WriteProperty("IsPDFButtonVisible", m_IsPDFButtonVisible, m_def_IsPDFButtonVisible)
    Call PropBag.WriteProperty("TesseractPath", OCRPath, "")
    Call PropBag.WriteProperty("OCRLanguage", Language, "")
    Me.Align = m_Align
End Sub

Private Sub UserControl_Resize()
    Dim UsercontrolHeight As Long
    Dim UsercontrolWidth As Long
    Dim UsercontrolTop As Long
    Dim UsercontrolLeft As Long
    UsercontrolHeight = UserControl.ScaleHeight + (m_IsToolbarVisible * ToolBar.Height) + (m_IsStatusBarVisible * StatusBar.Height)
    UsercontrolWidth = UserControl.ScaleWidth - PageScroll.Width
    UsercontrolTop = m_IsToolbarVisible * ToolBar.Height * -1
    If UsercontrolHeight < 100 Or UsercontrolWidth < 100 Then Exit Sub
    If FactorFormat > 1 Then
        PDFView.Height = UsercontrolHeight
        PDFView.Width = PDFView.Height / FactorFormat
        If PDFView.Width > UsercontrolWidth Then
            PDFView.Width = UsercontrolWidth
            PDFView.Height = PDFView.Width * FactorFormat
            PDFView.left = (UsercontrolWidth - PDFView.Width) / 2
            PDFView.top = UsercontrolTop + (UsercontrolHeight - PDFView.Height) / 2
        Else
            PDFView.left = (UsercontrolWidth - PDFView.Width) / 2
            PDFView.top = UsercontrolTop
        End If
    Else
        PDFView.Width = UsercontrolWidth
        PDFView.Height = PDFView.Width * FactorFormat
        If PDFView.Height > UsercontrolHeight Then
            PDFView.Height = UsercontrolHeight
            PDFView.Width = PDFView.Height / FactorFormat
            PDFView.left = (UsercontrolWidth - PDFView.Width) / 2
            PDFView.top = UsercontrolTop
        Else
            PDFView.top = ((UsercontrolHeight - PDFView.Height) / 2) + UsercontrolTop
            PDFView.left = 0
        End If
    End If
    TextPanel.top = 0
    TextPanel.Height = UserControl.ScaleHeight
    TextPanel.Width = UserControl.ScaleWidth
    TextPanel.left = 0
    PDFText.Width = TextPanel.Width - 150
    PDFText.Height = TextPanel.Height - PDFText.top - 75
    
On Error Resume Next
    PageScroll.top = UsercontrolTop
    PageScroll.left = UserControl.ScaleWidth - PageScroll.Width
    PageScroll.Height = UsercontrolHeight
    PagePanel.left = (StatusBar.Width - PagePanel.Width - PageScroll.Width) / 2
    ZoomPanel.left = (ToolBar.Width - ZoomPanel.Width - PageScroll.Width) / 2
    ZoomPanel2.left = ZoomPanel.left '(ToolBar.Width - ZoomPanel.Width) / 2
    If PageWith = 0 Then
        ZoomFactor = 100
    Else
        ZoomFactor = PDFView.Width / PageWith * 100
    End If
    lblZoomFactor.Caption = Format$(ZoomFactor / 100, "##0.00%")
    NoZoomChange = True
    ZoomScroll.Value = ZoomFactor ' (1 / ZoomFactor) * 100
    NoZoomChange = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
'
Public Property Get Align() As AlignConstants
Attribute Align.VB_Description = "Returns or sets a value that determines whether an object is displayed on a sheet."
    Align = m_Align
End Property

Public Property Let Align(ByVal New_Align As AlignConstants)
    If New_Align <> m_Align Then
        m_Align = New_Align
        PropertyChanged "Align"
        Me.Align = m_Align
    End If
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns or sets the text for the tooltip that appears when the mouse is over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PDFView.ToolTipText = m_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Function SelectPDFFile(Optional PDFPath As String = "", Optional ShowPDF As Boolean = True) As String
Attribute SelectPDFFile.VB_Description = "Open a file selector and load the selected document PDF. Displays it if ShowPDF = True. Returns the name of the loaded file."
    Dim bReturn As Boolean, eEvents As EventTypeEnum
    SelectPDFFile = ""
    cBrowser.Clear
    If PDFPath > "" Then cBrowser.InitDir = AddDirSep(PDFPath) Else cBrowser.InitDir = cBrowser.GetKnownFolderGUID(eFOLDERID_Documents)
    cBrowser.DialogTitle = "Open a PDF file..."
    cBrowser.FlagsDialog = DLG__BaseOpenDialogFlags
    cBrowser.Filter = "PDF Files|*.pdf" ' vbNullChar ' flag to hide the filter
    bReturn = cBrowser.ShowOpen(Me.hWnd, eEvents)
    If bReturn Then ' process dialog selected item(s)
        Call Load(cBrowser.FileName, ShowPDF)
        SelectPDFFile = cBrowser.FileName
        DoEvents
        UserControl_Resize
    End If
End Function
Public Function SelectPDFFileForOCR(Optional PDFPath As String = "") As String
Attribute SelectPDFFileForOCR.VB_Description = "Open a file selector and performs OCR on the PDF document"
    Dim bReturn As Boolean, eEvents As EventTypeEnum
    SelectPDFFileForOCR = ""
    cBrowser.Clear
    If PDFPath > "" Then cBrowser.InitDir = AddDirSep(PDFPath) Else cBrowser.InitDir = cBrowser.GetKnownFolderGUID(eFOLDERID_Documents)
    cBrowser.DialogTitle = "Ouvrir un fichier PDF..."
    cBrowser.FlagsDialog = DLG__BaseOpenDialogFlags
    cBrowser.Filter = "Fichiers PDF|*.pdf" ' vbNullChar ' flag to hide the filter
    bReturn = cBrowser.ShowOpen(Me.hWnd, eEvents)
    If bReturn Then ' process dialog selected item(s)
        SelectPDFFileForOCR = MakeOCROnFile(cBrowser.FileName)
        DoEvents
    End If
End Function

Public Function Load(ByVal FileName As String, Optional ShowPDF As Boolean = True) As Boolean
Attribute Load.VB_Description = "Loads a PDF file and displays the first page in the control if ShowPDF = True. Returns True if the file loaded correctly"
    Dim F As Integer, B() As Byte
    If Exist(FileName) Then
        If PDFLib = 0 Then 'If the PDFium library is not loaded, we load it and initialize it
            PDFLib = LoadLibraryW(StrPtr(App.Path & "\PDFium.dll"))
            FPDF_InitLibrary
            Set FormHook = New clsTrickSubclass2
            FormHook.Hook Me.hWnd
        End If
        Dim file_length As Long
        Dim fnum As Integer
        Dim bytes() As Byte
        Dim txt As String
        Dim i As Integer
        PDFLoadedFileName = ""
        NbPage = 0
        PageInView = 0
        ZoomFactor = 1
        PDFView.Cls
        TextPanel.Visible = False
        ToShow = False
        file_length = FileLen(FileName)
        fnum = FreeFile
        ReDim B(1 To file_length + 1)
    
        Open FileName For Binary As #fnum
        Get #fnum, 1, B
        Close fnum

        SetPDFByteContent B
        NbPage = PageCount
        
        If NbPage > 0 Then 'a PDF file is loaded
            PDFLoadedFileName = FileName
            TextPanel.Visible = False
            Load = True
            Debug.Print FileName & " *********************************************"
            PageHeight = PageHeightPoints(0) * ConvertPointToTwips
            PageWith = PageWidthPoints(0) * ConvertPointToTwips
            FactorFormat = PageHeight / PageWith
            Debug.Print FactorFormat
            If NbPage > 1 Then
                PageScroll.Max = NbPage
                PageScroll.Min = 1
                PageScroll.Enabled = True
            Else
                PageScroll.Enabled = False
            End If
            ToShow = ShowPDF
            UserControl_Resize
            ZoomFactor = PDFView.Width / PageWith * 100
            Debug.Print ZoomFactor
            PageInView = 1
            PageScroll.Value = PageInView
            
            PDFView.Cls
            If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
            Me.Refresh
            lblNbPages.Caption = NbPage
            TxPageInView.Text = PageInView
            lblZoomFactor.Caption = Format$(ZoomFactor / 100, "##0.00%")
            NoZoomChange = True
            ZoomScroll.Max = (Screen.Width / PageWith) * 100 '(1 / ZoomFactor) * 1000
            ZoomScroll.Value = ZoomFactor ' (1 / ZoomFactor) * 100
            Debug.Print (1 / ZoomFactor) * 100
            ZoomPanel.Enabled = True
            PagePanel.Enabled = True
            If NbPage > 0 Then
                RaiseEvent PDFLoaded(GetFileName(FileName), GetFilePath(FileName))
                If ToShow Then RaiseEvent PageChanged(1)
                DoEvents
                PDFView.SetFocus
            End If
            NoZoomChange = False
        End If

    End If
End Function
Private Sub SavePageForOCR(Page As Integer, strPath As String)
    Dim Page_Height As Double, Page_With As Double
    PictureOCR.Cls
    PictureOCR.Picture = Nothing
    Page_Height = PageHeightPoints(Page - 1) * ConvertPointToTwips
    Page_With = PageWidthPoints(Page - 1) * ConvertPointToTwips
    PictureOCR.Height = Page_Height
    PictureOCR.Width = Page_With
    DoEvents
    RenderPageToDC PictureOCR.hDC, Page - 1, 0, 0, PictureOCR.Width, PictureOCR.Height
    PictureOCR.Picture = PictureOCR.Image
    PicSaveLoad.SavePicture PictureOCR.Picture, AddDirSep(strPath) & "Page" & CStr(Page) & ".png", fmtPNG
End Sub
Public Function MakeOCROnFile(FileName As String) As String
Attribute MakeOCROnFile.VB_Description = "Perform character recognition on a PDF document"
    ' if the OCR is not installed we display an error and exit
    Dim OCRNotPossible As Boolean
    Dim OCRTmpPath As String
    If OCRPath = "" Then OCRNotPossible = True
    If OCRNotPossible = False Then
        If Exist(AddDirSep(OCRPath) & "tesseract.exe") = False Then OCRNotPossible = True
    End If
    If OCRNotPossible = True Then
        MsgBox "The character recognition function is not correctly installed!", vbCritical Or vbSystemModal, "OCR not installed"
        Exit Function
    End If
    
    Dim F As Integer, B() As Byte, CmdLine As String
    'if the file exists we load it
    If Exist(FileName) Then
        If PDFLib = 0 Then
            PDFLib = LoadLibraryW(StrPtr(App.Path & "\PDFium.dll"))
            FPDF_InitLibrary
            Set FormHook = New clsTrickSubclass2
            FormHook.Hook Me.hWnd
        End If
        Dim file_length As Long
        Dim fnum As Integer
        Dim bytes() As Byte
        Dim txt As String
        Dim i As Integer
        
        NbPage = 0
        TextPanel.Visible = False
        
        file_length = FileLen(FileName)
    
        fnum = FreeFile
        ReDim B(1 To file_length + 1)
    
        Open FileName For Binary As #fnum
        Get #fnum, 1, B
        Close fnum


        'we assign the PDF to PDFium
        SetPDFByteContent B
        'we recover the number of pages
        NbPage = PageCount
        If NbPage > 0 Then 'there are pages to read
            'we remove all traces of previous OCR
            OCRTmpPath = AddDirSep(KnownFolder(kfUserAppDataLocal)) & "Temp\OCR\"
            MD OCRTmpPath
            KillFile OCRTmpPath & "*.*"
            'we initialize the display controls
            PDFText.Text = ""
            PDFText.Visible = False
            lblText.Caption = "Convert PDF file ..."
            TextPanel.Visible = True
            'we save the pages as bitmaps and we prepare the batch file for Tesseract
            fnum = FreeFile
            Open OCRTmpPath & "Batch.txt" For Output As fnum
                For F = 1 To NbPage
                    SavePageForOCR F, OCRTmpPath
                    Print #fnum, OCRTmpPath & "Page" & CStr(F) & ".png"
                    lblText.Caption = "Convert PDF file ... " & CStr(Int((F / NbPage) * 100)) & " %"
                    DoEvents
                Next F
            Close #fnum
            'we launch the ocr
            lblText.Caption = "Extracting text ... Please wait"
            DoEvents
            CmdLine = AddDirSep(OCRPath) & "tesseract.exe """ & OCRTmpPath & "Batch.txt"" """ & OCRTmpPath & "OCR"" " & "--oem 1 -l " & Language
            RunShell CmdLine 'We launch the executable and we wait for the end of its execution
            If Exist(OCRTmpPath & "OCR.txt") Then 'An OCR text file exists ! we load it and we display it
                MakeOCROnFile = LoadFileInString(OCRTmpPath & "OCR.txt")
                MakeOCROnFile = UTF8_Decode(MakeOCR)
                MakeOCROnFile = Replace(MakeOCR, Chr$(13) + Chr$(10), Chr$(10))
                MakeOCROnFile = Replace(MakeOCR, Chr$(10), Chr$(13) + Chr$(10))
                PDFText.Text = MakeOCROnFile
                lblText.Caption = "OCR result"
                PDFText.Visible = True
            Else
                PDFText.Visible = True
                TextPanel.Visible = False
            End If
            CloseDocument
        End If
    End If
End Function
Private Function MakeOCR() As String
        Dim file_length As Long
        Dim fnum As Integer
        Dim bytes() As Byte
        Dim txt As String
        Dim OCRTmpPath As String
        Dim i As Integer, F As Integer, B() As Byte, CmdLine As String
        
        Dim OCRNotPossible As Boolean
        If OCRPath = "" Then OCRNotPossible = True
        If OCRNotPossible = False Then
            If Exist(AddDirSep(OCRPath) & "tesseract.exe") = False Then OCRNotPossible = True
        End If
        If OCRNotPossible = True Then
            MsgBox "The character recognition function is not correctly installed!", vbCritical Or vbSystemModal, "OCR not installed"
            Exit Function
        End If

        'we remove all traces of previous OCR
        OCRTmpPath = AddDirSep(KnownFolder(kfUserAppDataLocal)) & "Temp\OCR\"
        MD OCRTmpPath
        KillFile OCRTmpPath & "*.*"
        'we initialize the display controls
        PDFText.Text = ""
        PDFText.Visible = False
        lblText.Caption = "Convert PDF file ..."
        TextPanel.Visible = True
        'we save the pages as bitmaps and we prepare the batch file for Tesseract
        fnum = FreeFile
        Open OCRTmpPath & "Batch.txt" For Output As fnum
            For F = 1 To NbPage
                SavePageForOCR F, OCRTmpPath
                Print #fnum, OCRTmpPath & "Page" & CStr(F) & ".png"
                lblText.Caption = "Convert PDF file ... " & CStr(Int((F / NbPage) * 100)) & " %"
                DoEvents
            Next F
        Close #fnum
        'we launch the ocr
        lblText.Caption = "Extracting text ... Please wait"
        DoEvents
        CmdLine = AddDirSep(OCRPath) & "tesseract.exe """ & OCRTmpPath & "Batch.txt"" """ & OCRTmpPath & "OCR"" " & "--oem 1 -l " & Language
        RunShell CmdLine
        If Exist(OCRTmpPath & "OCR.txt") Then
            MakeOCR = LoadFileInString(OCRTmpPath & "OCR.txt")
            MakeOCR = UTF8_Decode(MakeOCR)
            MakeOCR = Replace(MakeOCR, Chr$(13) + Chr$(10), Chr$(10))
            MakeOCR = Replace(MakeOCR, Chr$(10), Chr$(13) + Chr$(10))
            PDFText.Text = MakeOCR
            lblText.Caption = "OCR result"
            PDFText.Visible = True
        Else
            PDFText.Visible = True
            TextPanel.Visible = False
        End If
End Function
Public Property Get IsToolbarVisible() As Boolean
Attribute IsToolbarVisible.VB_Description = "If True, displays the icon bar at the top of the document (File opening, Zoom ...)"
    IsToolbarVisible = m_IsToolbarVisible
End Property

Public Property Let IsToolbarVisible(ByVal New_IsToolbarVisible As Boolean)
    m_IsToolbarVisible = New_IsToolbarVisible
    ToolBar.Visible = m_IsToolbarVisible
    PropertyChanged "IsToolbarVisible"
    UserControl_Resize
End Property
Public Property Get DisplayedPage() As Integer
Attribute DisplayedPage.VB_Description = "Returns or sets the number of the page being displayed. Can only be edited when a PDF file is loaded in the control."
    DisplayedPage = PageInView
End Property

Public Property Let DisplayedPage(ByVal New_DisplayedPage As Integer)
    If New_DisplayedPage = 0 Then New_DisplayedPage = 1
    If NbPage > 0 Then
        If New_DisplayedPage > NbPage Then New_DisplayedPage = NbPage
        If PageInView <> New_DisplayedPage Then
            PageInView = New_DisplayedPage
            PDFView.Cls
            If ToShow Then PDFView.top = m_IsToolbarVisible * ToolBar.Height * -1: DoEvents
            If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
            If ToShow Then RaiseEvent PageChanged(PageInView)
            PageScroll.Value = PageInView
        End If
    Else
        PageInView = 0
    End If
    TxPageInView.Text = CStr(PageInView)
End Property
Public Property Get IsStatusBarVisible() As Boolean
Attribute IsStatusBarVisible.VB_Description = "If True, displays the status bar at the bottom of the control, allowing you to navigate within the pages of the document"
    IsStatusBarVisible = m_IsStatusBarVisible
    StatusBar.Visible = m_IsStatusBarVisible
End Property

Public Property Let IsStatusBarVisible(ByVal New_IsStatusBarVisible As Boolean)
    m_IsStatusBarVisible = New_IsStatusBarVisible
    StatusBar.Visible = m_IsStatusBarVisible
    PropertyChanged "IsStatusBarVisible"
    UserControl_Resize
End Property
Public Property Get IsPDFButtonVisible() As Boolean
Attribute IsPDFButtonVisible.VB_Description = "If True, displays the button to load a PDF file"
    IsPDFButtonVisible = m_IsPDFButtonVisible
End Property

Public Property Let IsPDFButtonVisible(ByVal New_IsPDFButtonVisible As Boolean)
    m_IsPDFButtonVisible = New_IsPDFButtonVisible
    PDFLoadButton.Visible = m_IsPDFButtonVisible
    PropertyChanged "IsPDFButtonVisible"
End Property
Private Function Exist(FilePath As String) As Boolean
    On Error GoTo ErrorHandler
    Call FileLen(FilePath)
    Exist = True
    Exit Function
ErrorHandler:
    Debug.Print Error$
    Exist = False
End Function
Private Function GetPDFByteContent() As Byte()
  GetPDFByteContent = Content
End Function

Private Sub SetPDFByteContent(PDFBytes() As Byte)
    Dim i As Long, DataLen As Long, Page_Count As Long
    CloseDocument
    Page_Count = 0
    hDoc = 0
    Content = PDFBytes
    DataLen = UBound(Content) - LBound(Content) + 1
    hDoc = FPDF_LoadMemDocument(VarPtr(Content(LBound(Content))), DataLen, "")
    If hDoc = 0 Then Exit Sub      'Err.Raise vbObjectError, , "Unable to open PDF file"
    Page_Count = FPDF_GetPageCount(hDoc)
    If Page_Count = 0 Then CloseDocument: Exit Sub      ':Err.Raise vbObjectError, , "Unable to open any page!"
    ReDim Pages(Page_Count)
    For i = 0 To Page_Count - 1
      Pages(i) = FPDF_LoadPage(hDoc, i)
      If Pages(i) = 0 Then CloseDocument: Exit Sub      ': Err.Raise vbObjectError, , "Unable to open page: " & i
    Next i
End Sub

Private Function PageCount() As Long
    On Error GoTo Fin:
    PageCount = UBound(Pages())
    Exit Function
Fin:
    PageCount = 0
End Function

Private Function PageWidthPoints(ByVal PageIdxZeroBased&) As Double
  PageWidthPoints = FPDF_GetPageWidth(Pages(PageIdxZeroBased))
End Function
Private Function PageHeightPoints(ByVal PageIdxZeroBased&) As Double
  PageHeightPoints = FPDF_GetPageHeight(Pages(PageIdxZeroBased))
End Function

Private Sub RenderPageToDC(ByVal hDC&, ByVal PageIdxZeroBased&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, Optional ByVal Rotation As ePDFPageRotation, Optional ByVal ClearType As Boolean = True, Optional ByVal ShowAnnotations As Boolean)
    Dim Flags As Long, hBM As Long, pData As Long
    If NbPage = 0 Then Exit Sub
    If dx <= 0 Then dx = 1
    If dy <= 0 Then dy = 1
    Flags = IIf(ClearType, 2, 0) Or IIf(ShowAnnotations, 1, 0)  '... Or FPDF_NO_GDIPLUS = 4 ... (could be used to speed things up on windows at the cost of quality)
    FPDF_RenderPage hDC, Pages(PageIdxZeroBased), X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, dx / Screen.TwipsPerPixelX, dy / Screen.TwipsPerPixelY, Rotation, Flags
End Sub
Private Sub CloseDocument()
  If hDoc = 0 Then Exit Sub
  Dim i As Long
  For i = 0 To UBound(Pages) - 1: FPDF_ClosePage Pages(i): Next
  FPDF_CloseDocument hDoc: hDoc = 0
  PDFLoadedFileName = ""
  NbPage = 0
  ZoomPanel.Enabled = False
  PagePanel.Enabled = False
  ToShow = False
End Sub

Private Sub ZoomPanel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ZoomScroll_Change
End Sub

Private Sub ZoomScroll_Change()
    Dim UsercontrolHeight As Long
    Dim UsercontrolWidth As Long
    Dim UsercontrolTop As Long
    Dim UsercontrolLeft As Long
    UsercontrolHeight = UserControl.ScaleHeight + (m_IsToolbarVisible * ToolBar.Height) + (m_IsStatusBarVisible * StatusBar.Height)
    UsercontrolWidth = UserControl.ScaleWidth - PageScroll.Width
    UsercontrolTop = m_IsToolbarVisible * ToolBar.Height * -1
    If NoZoomChange = False Then
            ZoomFactor = ZoomScroll.Value
            PDFView.Width = PageWith * ZoomFactor / 100
            PDFView.Height = PageHeight * ZoomFactor / 100
            PDFView.left = (UsercontrolWidth - PDFView.Width) / 2
            PDFView.Cls
            If ToShow Then RenderPageToDC PDFView.hDC, PageInView - 1, 0, 0, PDFView.Width, PDFView.Height
            Me.Refresh
            lblZoomFactor.Caption = Format(ZoomFactor, "##0.00") & "%"
    Else
        NoZoomChange = False
    End If
End Sub
Public Sub FitControl()
Attribute FitControl.VB_Description = "Adjusts the display of the PDF file to the size of the control."
    UserControl_Resize
    UserControl_Paint
End Sub
Public Sub ShowPDF()
Attribute ShowPDF.VB_Description = "Displays the previously loaded PDF file (See Load)"
    If NbPage > 0 Then
        ToShow = True
        UserControl_Paint
    End If
End Sub
Public Function GetPagesCount() As Integer
    GetPagesCount = NbPage
End Function

'File name manipulation functions
Private Function GetFilePath(PathAndName As String) As String
    If PathAndName = "" Then Exit Function
    If InStr(PathAndName, "\") = 0 Then
        GetFilePath = ""
    Else
        GetFilePath = left$(PathAndName, InStrRev(PathAndName, "\"))
    End If
End Function

Private Function GetFileLen(FilePath As String) As Long
    On Error GoTo ErrorHandlerFileLen
    GetFileLen = FileLen(FilePath)
    Exit Function
ErrorHandlerFileLen:
    GetFileLen = 0
End Function
Private Function GetFileName(PathAndName As String) As String
    If PathAndName = "" Then Exit Function
    If InStr(PathAndName, "\") = 0 Then
        GetFileName = PathAndName
    Else
        GetFileName = Mid$(PathAndName, InStrRev(PathAndName, "\") + 1)
    End If
End Function

Private Function GetFileExtension(PathAndName As String) As String
    If PathAndName = "" Then Exit Function
    Dim FileName As String
    FileName = GetFileName(PathAndName)
    If InStr(FileName, ".") = 0 Then
        GetFileExtension = ""
    Else
        GetFileExtension = Mid$(FileName, InStrRev(FileName, ".") + 1)
    End If
End Function
Private Function AddDirSep(strPathName As String) As String
    AddDirSep = Trim$(strPathName)
    If Right$(AddDirSep, 1) <> "\" Then AddDirSep = AddDirSep & "\"
End Function

Public Function GetIni(Section$, Item$, defaut$, IniName$) As String
    Dim retour$, longueur%
    retour$ = Space$(255)
    longueur% = GetPrivateProfileString(Section$, Item$, defaut$, retour$, 255, IniName$)
    GetIni$ = left$(retour$, longueur)
End Function
Sub SetIni(Section$, Item$, valeur$, IniName$)
    Dim ok
    ok = WritePrivateProfileString(Section$, Item$, valeur$, IniName$)
End Sub

Private Sub RunShell(CmdLine$, Optional WindowStyle As VbAppWinStyle = vbMinimizedFocus)
    Dim hProcess As Long
    Dim ProcessId As Long
    Dim ExitCode As Long, F As Integer
    ProcessId& = Shell(CmdLine$, WindowStyle)
    Screen.MousePointer = vbHourglass
    For F% = 1 To 10
        DoEvents
    Next F%
    hProcess& = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId&)
    Do
        Call GetExitCodeProcess(hProcess&, ExitCode&)
        DoEvents
    Loop While ExitCode& > 0
    Screen.MousePointer = vbArrow
    For F% = 1 To 10
        DoEvents
    Next F%
End Sub
Private Sub MD(DirPath As String) 'recursive directory creation
    If InStr(DirPath, "\") > 0 Then
        On Error Resume Next
        Dim strTMP$(), F%, strPath$
        strTMP$ = Split(DirPath, "\")
        For F% = 0 To UBound(strTMP()) - 1
            strPath$ = strPath$ & strTMP(F) & "\"
            MkDir strPath
        Next F%
    Else
        MkDir DirPath
    End If
    Err = 0
End Sub
Private Function LoadFileInString(ByVal FileName As String) As String
    Dim Handle As Integer
    If Exist(FileName) = False Then
        Err.Raise 53   ' File not found
    End If
    Handle = FreeFile
    Open FileName$ For Binary As #Handle
        LoadFileInString = Space$(LOF(Handle))
        Get #Handle, , LoadFileInString
    Close #Handle
End Function

Private Function DecodeURI(ByVal EncodedURI As String) As String
    Dim bANSI() As Byte
    Dim bUTF8() As Byte
    Dim lIndex As Long
    Dim lUTFIndex As Long

    If Len(EncodedURI) = 0 Then
        Exit Function
    End If
    bANSI = StrConv(EncodedURI, vbFromUnicode)          ' Convert from unicode text to ANSI values
    ReDim bUTF8(UBound(bANSI))                          ' Declare dynamic array, get length
    For lIndex = 0 To UBound(bANSI)                     ' from 0 to length of ANSI
        bUTF8(lUTFIndex) = bANSI(lIndex)                ' otherwise don't need to do anything special
        lUTFIndex = lUTFIndex + 1                       ' advance utf index
    Next
    DecodeURI = From_UTF8(bUTF8, lUTFIndex)             ' convert to string
End Function
Private Function From_UTF8(ByRef UTF8() As Byte, ByVal Length As Long) As String
    Dim lDataLength As Long
    lDataLength = MultiByteToWideChar(CP_UTF8, 0, VarPtr(UTF8(0)), Length, 0, 0)  ' Get the length of the data.
    From_UTF8 = String$(lDataLength, 0)                                         ' Create array big enough
    MultiByteToWideChar CP_UTF8, 0, VarPtr(UTF8(0)), _
                        Length, StrPtr(From_UTF8), lDataLength                  '
End Function
Public Function UTF8_Decode(ByVal Text As String) As String
    UTF8_Decode = DecodeURI(Text)
End Function
Private Function KillFile(strFileToKill) As Integer
    On Error GoTo Fin
    Kill strFileToKill
Fin:
    KillFile = Err
End Function

Public Property Get PDFiumLibraryHandle() As Long
Attribute PDFiumLibraryHandle.VB_Description = "Handle of the PDFium Library"
    PDFiumLibraryHandle = PDFLib
End Property

Public Property Let PDFiumLibraryHandle(ByVal vNewValue As Long)

End Property
Public Sub PrintPDF(Optional Copies As Integer = 1, Optional Orientation As oPrinterOrientation = oAuto, Optional FromPage As Integer = 0, Optional ToPage As Integer = 0, Optional PrinterName As String = "")
    Dim X As Printer, Imp_defaut$, F As Integer, i As Integer
    Dim imp_actuelle$
    Dim Page_Height As Double, Page_Width As Double
    Dim row As Integer
    Dim column As Integer, Width As Double, Height As Double, top As Long, left As Long, Zoom As Single, columns As Integer, rows As Integer
    Dim vpos As Long, hpos As Long
    Dim margin As Integer, colour As Long, backcolour As String
    
    If NbPage = 0 Then Exit Sub
    'We keep the current printer
    imp_actuelle$ = Printer.DeviceName
    'Check if the right printer exists
    If PrinterName <> imp_actuelle$ And PrinterName > "" Then
        'If yes, change the default printer
        For Each X In Printers
            If X.DeviceName = PrinterName Then
                Set Printer = X
                Exit For
            End If
        Next
    End If
    If Copies < 1 Then Copies = 1
    If Copies > 999 Then Copies = 999
    If FromPage < 1 Then FromPage = 1
    If FromPage > NbPage Then FromPage = 1
    If ToPage = 0 Then ToPage = NbPage
    If ToPage > NbPage Then ToPage = NbPage
    If ToPage < FromPage Then ToPage = FromPage

    For i = 1 To Copies
        'Printer initialization
        Page_Height = PageHeightPoints(0) * ConvertPointToTwips
        Page_Width = PageWidthPoints(0) * ConvertPointToTwips
        If Orientation = oAuto Then

            If Page_Width <= Page_Height Then Orientation = oPortrait Else Orientation = oLandscape
        End If
        Printer.Orientation = Orientation
        margin = 0 ' 150
        top = Printer.ScaleTop '+ margin
        left = Printer.ScaleLeft '+ margin
        columns = 1 ' to modify if you want to print several pages per sheet
        rows = 1
        Width = (Printer.ScaleWidth + margin) / columns - margin '(margin * 2)
        Height = (Printer.ScaleHeight + margin) / rows - margin '(margin * 2)
        hpos = left + (columns - 1) * (Width + margin)
        vpos = top + (rows - 1) * (Height + margin)
        'check whether width or height is limiting factor
        Zoom = Width / Page_Width
        If Zoom < Height / Page_Height Then
            ' width limited
            vpos = vpos + (Height - Page_Height * Zoom) / 2
        Else
            ' height limited
            Zoom = Height / Page_Height
            hpos = hpos + (Width - Page_Width * Zoom) / 2
        End If
        For F = FromPage To ToPage
            'we draw, store and print the image of the page
            'Printer.Print "  "
            PictureOCR.Cls
            PictureOCR.Picture = Nothing
            Page_Height = PageHeightPoints(F - 1) * ConvertPointToTwips
            Page_Width = PageWidthPoints(F - 1) * ConvertPointToTwips
            PictureOCR.Height = Page_Height
            PictureOCR.Width = Page_Width
            DoEvents
            RenderPageToDC PictureOCR.hDC, F - 1, 0, 0, PictureOCR.Width, PictureOCR.Height
            PictureOCR.Picture = PictureOCR.Image
            DoEvents
            Printer.PaintPicture PictureOCR.Picture, hpos, vpos, PictureOCR.Width * Zoom, PictureOCR.Height * Zoom
            Printer.NewPage
        Next F
        Printer.EndDoc
    Next i
    'Possible replacement of the default printer
    If imp_actuelle <> Printer.DeviceName Then
        For Each X In Printers
            If X.DeviceName = imp_actuelle Then
                Set Printer = X
                Exit For
            End If
        Next
    End If
End Sub
Public Function CopyPageToClipboard() As Boolean
    Dim Page_Height As Double, Page_Width As Double
    CopyPageToClipboard = False
    If NbPage = 0 Then Exit Function
    On Error GoTo Fin
    PictureOCR.Cls
    PictureOCR.Picture = Nothing
    Page_Height = PageHeightPoints(PageInView - 1) * ConvertPointToTwips
    Page_Width = PageWidthPoints(PageInView - 1) * ConvertPointToTwips
    PictureOCR.Height = Page_Height
    PictureOCR.Width = Page_Width
    DoEvents
    RenderPageToDC PictureOCR.hDC, PageInView - 1, 0, 0, PictureOCR.Width, PictureOCR.Height
    PictureOCR.Picture = PictureOCR.Image
    Clipboard.Clear
    Clipboard.SetData PictureOCR.Picture, vbCFBitmap
    CopyPageToClipboard = True
    Exit Function
Fin:

End Function
