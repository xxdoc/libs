VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   6060
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   480
      Width           =   2115
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   6060
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   0
      Width           =   2115
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   60
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   900
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "memEntropy last 1000 bytes"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   8
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "fileEntropy last 1000 bytes"
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   6
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "file"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "memEntropy"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "fileEntropy"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    
    Dim b() As Byte
    Dim f As Long
    Dim pth As String
    
    pth = "c:\windows\notepad.exe"
    f = FreeFile
    Open pth For Binary As f
    ReDim b(LOF(f) - 1)
    Get f, , b()
    Close f
    

    Text1 = memEntropy(b)
    Text5 = memEntropy(b, UBound(b) - 1000)
    Text2 = pth
    
    Text3 = fileEntropy(pth)
    Text4 = fileEntropy(pth, UBound(b) - 1000) 'just last 1000 bytes
    
        
End Sub

'ported from the entropy calc in Detect It Easy...
Function fileEntropy(pth As String, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Dim nSize As Long
    Dim nTemp As Long
    Const BUFFER_SIZE = &H1000
    Dim buf() As Byte
    Dim f As Long
    
    On Error Resume Next
    
    f = FreeFile
    Open pth For Binary Access Read As f
    If Err.Number <> 0 Then GoTo ret0
    
    sz = LOF(f) - 1
    
    If leng = 0 Then GoTo ret0
    
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    Seek f, offset
    nSize = leng
    fEntropy = 1.44269504088896
    ReDim buf(BUFFER_SIZE)
    
    'read the file in chunks and count how many times each byte value occurs
    While (nSize > 0)
        nTemp = IIf(nSize < BUFFER_SIZE, nSize, BUFFER_SIZE)
        If nTemp <> BUFFER_SIZE Then ReDim buf(nTemp) 'last chunk, partial buffer
        Get f, , buf()
        For i = 0 To UBound(buf)
            bytes(buf(i)) = bytes(buf(i)) + 1
        Next
        nSize = nSize - nTemp
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    fileEntropy = fEntropy / CSng(leng)
    
Exit Function
ret0:
    Close f
End Function


Function memEntropy(buf() As Byte, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Const BUFFER_SIZE = &H1000
    
    sz = UBound(buf)
    
    If leng = 0 Then GoTo ret0
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    fEntropy = 1.44269504088896
    
    While (offset < sz)
        'count each byte value occurance
        bytes(buf(offset)) = bytes(buf(offset)) + 1
        offset = offset + 1
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    memEntropy = fEntropy / CSng(leng)
    
Exit Function
ret0:
End Function
