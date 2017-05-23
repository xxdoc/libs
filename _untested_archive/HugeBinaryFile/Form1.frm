VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HugeBinaryFile Demo"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2280
      Top             =   660
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label lblRead 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   1110
      Width           =   3375
   End
   Begin VB.Label lblWritten 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   330
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Timer-driven demo of HugeBinaryFile class.
'

Private hbfFile As HugeBinaryFile
Private blnWriting As Boolean
Private bytBuf(1 To 1000000) As Byte
Private lngBlocks As Long
Private Const MAX_BLOCKS As Long = 5000

Private Sub cmdRead_Click()
    cmdWrite.Enabled = False
    cmdRead.Enabled = False
    lngBlocks = 0
    lblRead.Caption = ""
    blnWriting = False
    Set hbfFile = New HugeBinaryFile
    hbfFile.OpenFile "test.dat"
    lblStatus = " Reading " _
              & Format$(hbfFile.FileLen, "##,###,###,###,##0") _
              & " bytes"
    Timer1.Enabled = True
End Sub

Private Sub cmdWrite_Click()
    cmdWrite.Enabled = False
    cmdRead.Enabled = False
    On Error Resume Next
    Kill "test.dat"
    On Error GoTo 0
    lngBlocks = 0
    lblWritten.Caption = ""
    lblStatus = " Writing " _
              & Format$(CCur(MAX_BLOCKS) * CCur(UBound(bytBuf)), "##,###,###,###,##0") _
              & " bytes"
    blnWriting = True
    Set hbfFile = New HugeBinaryFile
    hbfFile.OpenFile "test.dat"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (hbfFile Is Nothing) Then
        If hbfFile.IsOpen Then hbfFile.CloseFile
        Set hbfFile = Nothing
    End If
End Sub

Private Sub Timer1_Timer()
    If blnWriting Then
        hbfFile.WriteBytes bytBuf
        lngBlocks = lngBlocks + 1
        lblWritten.Caption = _
                Format$(CCur(lngBlocks) * CCur(UBound(bytBuf)), "##,###,###,###,##0") _
              & " bytes written"
        If lngBlocks >= MAX_BLOCKS Then
            Timer1.Enabled = False
            hbfFile.CloseFile
            Set hbfFile = Nothing
            lblStatus = ""
            cmdWrite.Enabled = True
            cmdRead.Enabled = True
        End If
    Else
        hbfFile.ReadBytes bytBuf
        If hbfFile.EOF Then
            Timer1.Enabled = False
            hbfFile.CloseFile
            Set hbfFile = Nothing
            lblStatus = ""
            cmdWrite.Enabled = True
            cmdRead.Enabled = True
        Else
            lngBlocks = lngBlocks + 1
            lblRead.Caption = _
                    Format$(CCur(lngBlocks) * CCur(UBound(bytBuf)), "##,###,###,###,##0") _
                  & " bytes read"
        End If
    End If
End Sub
