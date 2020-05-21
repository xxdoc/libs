VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   13110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Download HTML"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtHTML 
      Height          =   5175
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   12015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download Pic"
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   5160
      ScaleHeight     =   2235
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   240
      Width           =   6375
   End
   Begin Project1.ctlDownload ctlDownloadPicture 
      Height          =   960
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
   End
   Begin Project1.ctlDownload ctlDownload1 
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
ctlDownloadPicture.Download "http://www.vbforums.com/attachment.php?attachmentid=162305&d=1538942156"
End Sub


Private Sub Command2_Click()
ctlDownload1.Download "http://www.vbforums.com/showthread.php?867057-ActiveX-Download-File"
End Sub

Private Sub ctlDownload1_Finished(X As AsyncProperty)
    If X.StatusCode = vbAsyncStatusCodeEndDownloadData Then
        'to get website text
        Dim MainText As String
        MainText = Replace(StrConv(X.Value, vbUnicode), vbLf, vbNewLine)
        txtHTML.Text = MainText
        
    End If
End Sub



Private Sub ctlDownloadPicture_Finished(X As AsyncProperty)

    Dim Bytes() As Byte, fnum As Integer
    If X.StatusCode = vbAsyncStatusCodeEndDownloadData Then

        Bytes = X.Value
        ' Save the file.
        fnum = FreeFile
        Open App.Path & "\Largeimage.jpg" For Binary As #fnum
        Put #fnum, 1, Bytes()
        Close fnum

        Erase Bytes
        
        Picture1.Picture = LoadPicture(App.Path & "\Largeimage.jpg")

    End If

End Sub
