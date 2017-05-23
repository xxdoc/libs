VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zip Demo"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
   End
   Begin ZipDemo.Zipper Zipper 
      Left            =   900
      Top             =   1680
      _extentx        =   741
      _extenty        =   741
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Zipper"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1673
      TabIndex        =   1
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ZipWriter"
      Height          =   495
      Left            =   1673
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   4515
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4500
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim Bytes() As Byte
    
    Command1.Enabled = False
    Refresh
    
    On Error Resume Next
    Kill "test1.zip"
    On Error GoTo 0
    
    Bytes = StrConv("Hello world!" & vbNewLine & vbNewLine _
                  & "A brave new day for all concerned as we " _
                  & "venture forth striving for success." & vbNewLine, _
                    vbFromUnicode)
    
    With New ZipWriter
        If .OpenZip("test.zip") Then
            MsgBox "OpenZip Error " & CStr(.Result)
        Else
            If .OpenFileInZip("FolderA\A.txt", , , vbReadOnly) Then
                MsgBox "OpenFileInZip A Error " & CStr(.Result)
            Else
                If .WriteBytes(Bytes) Then
                    MsgBox "WriteBytes A Error " & CStr(.Result)
                Else
                    If .CloseFileInZip() Then
                        MsgBox "CloseFileInZip A Error " & CStr(.Result)
                    Else
                        MsgBox "Success A"
                    End If
                End If
            End If
                        
            If .OpenFileInZip("B.txt") Then
                MsgBox "OpenFileInZip B Error " & CStr(.Result)
            Else
                If .WriteBytes(Bytes) Then
                    MsgBox "WriteBytes B Error " & CStr(.Result)
                Else
                    If .CloseFileInZip() Then
                        MsgBox "CloseFileInZip B Error " & CStr(.Result)
                    Else
                        MsgBox "Success B"
                    End If
                End If
            End If
            
            If .CloseZip() Then
                MsgBox "CloseZip Error " & CStr(.Result)
            Else
                MsgBox "Success, Complete"
            End If
        End If
    End With
    
    Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    Dim FileName As String
    Dim Failed As Boolean
    
    Command2.Enabled = False
    Command3.Enabled = True
    Label1.Caption = ""
    Refresh
    
    FileName = Dir$("samples2\*.*")
    With Zipper
        Do While Len(FileName) > 0
            If (GetAttr("samples2\" & FileName) And vbDirectory) = 0 Then
                If .AddFile("samples2\" & FileName, "ZipperTest\" & FileName) Then
                    Label1.Caption = "AddFile failed on " & FileName
                    Failed = True
                    Exit Do
                End If
            End If
            FileName = Dir$()
        Loop
        If Not Failed Then
            If .Zip("test.zip", APPEND_STATUS_ADDINZIP) Then
                Label1.Caption = "Err " & CStr(.Result) & " in " & .Failed
            End If
        End If
    End With
End Sub

Private Sub Command3_Click()
    Zipper.Cancel
End Sub

Private Sub Zipper_Complete(ByVal Canceled As Boolean)
    If Canceled Then
        Label1.Caption = "Canceled"
    Else
        Label1.Caption = "Success"
    End If
    Command3.Enabled = False
End Sub

Private Sub Zipper_EndFile()
    Label1.Caption = "Zipped " & Zipper.CurrentFile.SourceFile
    Label1.Refresh
End Sub

Private Sub Zipper_Error()
    With Zipper
        Label1.Caption = "Err " & CStr(.Result) & " in " & .Failed
    End With
End Sub

Private Sub Zipper_Progress()
    With Zipper
        ProgressBar1.Value = .BytesZipped / .BytesToZip * 100
    End With
End Sub

Private Sub Zipper_StartFile()
    Label1.Caption = "Zipping " & Zipper.CurrentFile.SourceFile
    Label1.Refresh
End Sub
