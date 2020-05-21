VERSION 5.00
Object = "*\APDFReader.vbp"
Begin VB.Form frmTest 
   Caption         =   "PDFReader"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15735
   Icon            =   "TEST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin PDF_Reader.PDFReader PDFReader1 
      Height          =   8985
      Left            =   2280
      TabIndex        =   22
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   15849
      BorderStyle     =   1
      TesseractPath   =   "C:\Program Files\Tesseract-OCR"
      OCRLanguage     =   "eng"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copy the page image to the clipboard"
      Height          =   560
      Left            =   135
      TabIndex        =   21
      Top             =   7605
      Width           =   2040
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print first page"
      Height          =   560
      Left            =   135
      TabIndex        =   20
      Top             =   6885
      Width           =   2040
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OCR"
      Height          =   560
      Left            =   120
      TabIndex        =   19
      Top             =   6180
      Width           =   2040
   End
   Begin VB.Frame Frame4 
      Caption         =   "Back Color"
      Height          =   465
      Left            =   135
      TabIndex        =   14
      Top             =   5640
      Width           =   2040
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1500
         ScaleHeight     =   165
         ScaleWidth      =   390
         TabIndex        =   18
         Top             =   195
         Width           =   420
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1020
         ScaleHeight     =   165
         ScaleWidth      =   390
         TabIndex        =   17
         Top             =   195
         Width           =   420
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFFFF&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   540
         ScaleHeight     =   165
         ScaleWidth      =   390
         TabIndex        =   16
         Top             =   195
         Width           =   420
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         ScaleHeight     =   165
         ScaleWidth      =   390
         TabIndex        =   15
         Top             =   195
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Show PDF Load button"
      Height          =   465
      Left            =   135
      TabIndex        =   11
      Top             =   5070
      Width           =   2040
      Begin VB.OptionButton Option1 
         Caption         =   "No"
         Height          =   195
         Index           =   5
         Left            =   1185
         TabIndex        =   13
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   12
         Top             =   225
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Show status bar"
      Height          =   465
      Left            =   150
      TabIndex        =   8
      Top             =   4545
      Width           =   2040
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No"
         Height          =   195
         Index           =   2
         Left            =   1185
         TabIndex        =   9
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show toolbar"
      Height          =   465
      Left            =   150
      TabIndex        =   5
      Top             =   4020
      Width           =   2040
      Begin VB.OptionButton Option1 
         Caption         =   "No"
         Height          =   195
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Top             =   225
         Width           =   690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Goto page ..."
      Height          =   560
      Left            =   140
      TabIndex        =   4
      Top             =   3225
      Width           =   2040
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Zoom 100%"
      Height          =   560
      Left            =   140
      TabIndex        =   3
      Top             =   1776
      Width           =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fit to Control"
      Height          =   560
      Left            =   140
      TabIndex        =   2
      Top             =   2499
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load a PDF file"
      Height          =   560
      Left            =   140
      TabIndex        =   1
      Top             =   1053
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Test.pdf"
      Height          =   560
      Left            =   140
      TabIndex        =   0
      Top             =   330
      Width           =   2040
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    PDFReader1.Load App.Path & "\test.pdf"
End Sub

Private Sub Command2_Click()
    PDFReader1.SelectPDFFile App.Path
End Sub

Private Sub Command3_Click()
    PDFReader1.FitControl
End Sub

Private Sub Command4_Click()
    PDFReader1.Zoom = 100
End Sub

Private Sub Command5_Click()
    Dim Numero As Integer
    Numero = Val(InputBox("Go to the page (1 à " & PDFReader1.GetPagesCount & ")", "Enter page number"))
    
    If Numero > 0 Then PDFReader1.DisplayedPage = Numero
End Sub

Private Sub Command6_Click()
    PDFReader1.SelectPDFFileForOCR
End Sub

Private Sub Command7_Click()
    PDFReader1.PrintPDF , , 1, 1
End Sub

Private Sub Command8_Click()
    PDFReader1.CopyPageToClipboard
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PDFReader1.Height = Me.ScaleHeight
    PDFReader1.Width = Me.ScaleWidth - PDFReader1.Left
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value Then PDFReader1.IsToolbarVisible = True
    If Option1(1).Value Then PDFReader1.IsToolbarVisible = False
    If Option1(2).Value Then PDFReader1.IsStatusBarVisible = False
    If Option1(3).Value Then PDFReader1.IsStatusBarVisible = True
    If Option1(4).Value Then PDFReader1.IsPDFButtonVisible = True
    If Option1(5).Value Then PDFReader1.IsPDFButtonVisible = False
End Sub

Private Sub PDFReader1_PageChanged(PageViewed As Integer)
    Debug.Print PageViewed
End Sub

Private Sub PDFReader1_PDFLoaded(FileName As String, FilePath As String)
    Debug.Print FileName, FilePath
    Me.Caption = "PDFReader" & " : " & FileName & " (" & FilePath & ")"
End Sub

Private Sub Picture1_Click(Index As Integer)
    PDFReader1.BackColor = Picture1(Index).BackColor
End Sub
