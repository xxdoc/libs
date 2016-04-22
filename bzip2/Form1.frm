VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   2805
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   2520
      Width           =   9240
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   765
      TabIndex        =   3
      Top             =   585
      Width           =   6180
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Byte Array Test"
      Height          =   375
      Left            =   7065
      TabIndex        =   2
      Top             =   630
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "File Compress Test"
      Height          =   375
      Left            =   7065
      TabIndex        =   1
      Top             =   45
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   765
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "Drag and drop file "
      Top             =   45
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bzip As New clsBzip2
Dim hash As Object 'CWinHash 'optional..my vb devkit installer: http://sandsprite.com/CodeStuff/vbdevkit.exe

Private Sub Command1_Click()

    Dim comp As String, dec As String
    Dim a As String, b As String
    
    List1.Clear
    Text2 = Empty
    
    If Not bzip.FileExists(Text1) Then
        List1.AddItem "Failed to open file: " & Text1
        Exit Sub
    End If
    
    If Not hash Is Nothing Then
        List1.AddItem "Input file hash: " & hash.HashFile(Text1) & " Sz: " & FileLen(Text1)
    Else
        List1.AddItem "Input file Sz: " & FileLen(Text1)
    End If
    
    comp = "c:\compressed.txt"
    dec = "c:\decompressed.txt"
    
    If Not bzip.CompressFile(Text1, comp) Then
        List1.AddItem "Compress failed"
        Exit Sub
    End If
    
    List1.AddItem "comp ok"
    
    If Not bzip.DecompressFile(comp, dec) Then
        List1.AddItem "deCompress failed"
        Exit Sub
    End If
    
    If Not hash Is Nothing Then
        List1.AddItem "Decompressed file hash: " & hash.HashFile(dec) & " Sz: " & FileLen(dec)
    Else
        List1.AddItem "Decompressed file Sz: " & FileLen(dec)
    End If

    If hash Is Nothing Then
        a = ReadFile(Text1)
        b = ReadFile(dec)
        If a = b Then
            List1.AddItem "Data matched!"
        Else
            List1.AddItem "Data failed!"
            Text2 = hexdump(b)
        End If
    End If
    
End Sub

Private Sub Command2_Click()
    
    Dim b() As Byte, b2() As Byte, b3() As Byte
    Dim org As String, comp As String, dec As String
    
    Text2 = Empty
    List1.Clear
    
    org = String(50, "A")
    b() = StrConv(org, vbFromUnicode)
    
    If Not bzip.CompressData(b, b2) Then
        List1.AddItem "Compress failed"
        Exit Sub
    End If

    List1.AddItem "Compressed size: " & UBound(b2)
    
    If Not bzip.DecompressData(b2, b3) Then
        List1.AddItem "Decompress failed"
        Exit Sub
    End If
    
    List1.AddItem "Decompressed size: " & UBound(b3)

    dec = StrConv(b3, vbUnicode)
    
    If dec = org Then
        List1.AddItem "Data matched!"
    Else
        List1.AddItem "Data failed!"
        Text2 = hexdump(dec)
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set hash = CreateObject("vbDevKit.CWinHash")
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text1 = Data.Files(1)
End Sub

Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Function hexdump(it)
    Dim my, i, c, s, a, b
    Dim lines() As String
    
    my = ""
    For i = 1 To Len(it)
        a = Asc(Mid(it, i, 1))
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        b = b & IIf(a >= 20 And a < 123, Chr(a), ".")
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
        
    If Len(it) < 16 Then
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


