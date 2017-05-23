VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFileTest 
      Caption         =   "File Test"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Test "
      Height          =   1035
      Left            =   60
      TabIndex        =   5
      Top             =   3780
      Width           =   8475
      Begin VB.TextBox txtDecompFile 
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "c:\file.decompressed"
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox txtCompFile 
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "c:\file.compressed"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1620
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   240
         Width           =   5355
      End
      Begin VB.Label Label4 
         Caption         =   "Decompressed"
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Compressed"
         Height          =   255
         Left            =   420
         TabIndex        =   8
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "File In (Drag && Drop)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   8595
   End
   Begin VB.CommandButton cmdDecompress 
      Caption         =   "Decompress"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Compress"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8595
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   3180
      Width           =   5475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim buf() As Byte
Dim ucl As New CUclCompression


Private Sub Form_Load()
    Text1 = "aaaa bbbb this is my text this is my text its mine! aaaa, bbbb"
End Sub

Private Sub cmdCompress_Click()
    
    Dim bout() As Byte
    
    buf() = StrConv(Text1, vbFromUnicode)
    
    If Not ucl.Compress(buf, bout) Then
        MsgBox ucl.LastError
        Exit Sub
    End If
    
    Text2 = hexdump(bout)
    
    Label1 = "Buffer has been compressed from " & UBound(buf) & " to " & UBound(bout)
    
    buf() = bout()
    
    cmdDecompress.Enabled = True
    
End Sub

Private Sub cmdDecompress_click()
    
    Dim b() As Byte
    
    If Not ucl.Decompress(buf, 2000, b) Then
        MsgBox "Error: " & ucl.LastError
        Exit Sub
    End If
    
    Text1 = StrConv(b, vbUnicode)
    Text2 = hexdump(b)
    
    Label1 = "Buffer has been expanded from " & UBound(buf) & " to " & UBound(b)

End Sub

Private Sub cmdFileTest_Click()

    Dim orgSize As Long
    
    If Not ucl.FileExists(txtFile) Then
        MsgBox "File to compress not found"
        Exit Sub
    End If
    
    orgSize = FileLen(txtFile)
    
    If Not ucl.CompressFile(txtFile, txtCompFile, True) Then
        MsgBox "Compression Failed " & ucl.LastError
        Exit Sub
    End If
    
    If Not ucl.DeCompressFile(txtCompFile, txtDecompFile, True, orgSize * 2) Then
        MsgBox "Decompression failed " & ucl.LastError
        Exit Sub
    End If
    
   MsgBox "Original Size: " & FileLen(txtFile) & vbCrLf & _
                 "Compressed: " & FileLen(txtCompFile) & vbCrLf & _
                 "Decompressed:" & FileLen(txtDecompFile) & vbCrLf & vbCrLf & _
                 "Now you should MD5 before and after files to make sure"
                 
End Sub




'----------------------------------------------------------------------
'Library functions below
'----------------------------------------------------------------------

Function hexdump(it() As Byte)
    Dim my, i, c, s, a, b
    Dim lines() As String
    
    my = ""
    For i = 1 To UBound(it) + 1
        a = it(i - 1)
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        b = b & IIf(a > 65 And a < 145, Chr(a), ".")
        my = my & c & " "
        If i Mod 16 = 0 Then
            push lines(), my & "  [" & b & "]"
            my = Empty
            b = Empty
        End If
    Next
    
    If Len(b) > 0 Then
        If Len(my) < 48 Then
            my = my & String(48 - Len(my), " ")
        End If
        If Len(b) < 16 Then
             b = b & String(16 - Len(b), " ")
        End If
        push lines(), my & "  [" & b & "]"
    End If
        
    If UBound(it) < 16 Then
        hexdump = my & "  [" & b & "]" & vbCrLf
    Else
        hexdump = Join(lines, vbCrLf)
    End If
    
    
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtFile = Data.Files(1)
End Sub
