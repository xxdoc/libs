VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   135
      TabIndex        =   4
      Top             =   6165
      Width           =   12525
   End
   Begin VB.TextBox txtDecomp 
      Height          =   5685
      Left            =   6480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   6225
   End
   Begin VB.TextBox txtCompressed 
      Height          =   5685
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   6225
   End
   Begin VB.Label Label2 
      Caption         =   "Decompressed"
      Height          =   240
      Left            =   6525
      TabIndex        =   3
      Top             =   45
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Compressed"
      Height          =   240
      Left            =   135
      TabIndex        =   2
      Top             =   45
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'note if you hit stop in the ide without closing the form, you probably
'wont be able to recompile the dll without closing out the ide.
'thats what the freelibrary call in form_unload is for..

Dim hLib As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Enum eMsg
    em_version = 0
    em_lastErr = 1
End Enum
    
Private Declare Function LZOGetMsg Lib "minilzo.dll" ( _
                    ByVal buf As String, _
                    ByVal sz As Long, _
                    Optional ByVal msgid As eMsg = em_version _
                ) As Long

'int __stdcall Compress(char* buf, int bufsz, char* bOut, int bOutSz)
Private Declare Function Compress Lib "minilzo.dll" ( _
            ByVal bufIn As String, _
            ByVal inSz As Long, _
            ByVal bufOut As String, _
            ByVal outSz As Long _
        ) As Long

Private Declare Function DeCompress Lib "minilzo.dll" ( _
            ByVal bufIn As String, _
            ByVal inSz As Long, _
            ByVal bufOut As String, _
            ByVal outSz As Long _
        ) As Long



Private Sub Form_Load()

    hLib = LoadLibrary("minilzo.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\minilzo.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\minilzo.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\..\minilzo.dll")
    
    If hLib = 0 Then
        MsgBox "We could not find the dll? or its corrupt?"
        End
    End If
    
    List1.AddItem "minilzo.dll found.."
    List1.AddItem LZOMsg(em_version)
    
    Dim a As String
    Dim compressed As String
    Dim decompressed As String
    
    a = String(100000, "A")
    If Not LZO(a, compressed) Then Exit Sub
    txtCompressed = hexdump(compressed)
    List1.AddItem Len(a) & " bytes compressed down to " & Len(compressed)
    
    If Not LZO(compressed, decompressed, Len(a)) Then Exit Sub
    txtDecomp = hexdump(decompressed)
    List1.AddItem "Decompressed size is now " & Len(decompressed)
    
    If decompressed = a Then
        List1.AddItem "Success original and decompressed strings match!"
    Else
        List1.AddItem "FAIL! - len(org) = " & Len(a) & " len(decomp) = " & Len(decompressed)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'so the ide doesnt hang onto it and we can recompile..
    If hLib <> 0 Then FreeLibrary hLib
End Sub

'for decompression, it would probably be better to pass in the original size to get an idea
'of the buffer size to allocate. in practice I would include a header in comporessed data
'that included original size and original md5
'
'note: passing in orgCompressedSize tells it you want to decompress the data..
Function LZO(buf As String, ByRef retVal As String, Optional orgCompressedSize As Long = 0) As Boolean
    
    Dim bOut As String
    Dim inSz As Long
    Dim outlen As Long
    
    '/* We want to compress the data block at 'in' with length 'IN_LEN' to
    '* the block at 'out'. Because the input block may be incompressible,
    '* we must provide a little more output space in case that compression
    '* is not possible.
    '*/
    
    inSz = Len(buf)
    If orgCompressedSize = 0 Then
        outlen = inSz * 2
    Else
        outlen = orgCompressedSize * 2
    End If
    
    bOut = String(outlen, Chr(0))
    
    If orgCompressedSize = 0 Then
        sz = Compress(buf, inSz, bOut, outlen)
    Else
        sz = DeCompress(buf, inSz, bOut, outlen)
    End If
    
    If sz < 1 Then
        List1.AddItem IIf(orgCompressedSize = 0, "De", "") & "Compression failed: " & LZOMsg()
        Exit Function
    End If
    
    retVal = Mid(bOut, 1, sz)
    LZO = True
        
End Function

Function LZOMsg(Optional m As eMsg = em_lastErr)
    Dim ver As String
    Dim sz As Long
    
    ver = String(500, Chr(0))
    sz = LZOGetMsg(ver, Len(ver), m)
    If sz > 0 Then
        ver = Mid(ver, 1, sz)
        LZOMsg = ver
    End If
    
End Function





Function hexdump(it)
    Dim my, i, c, s, a, b
    Dim lines() As String
    
    my = ""
    For i = 1 To Len(it)
        a = Asc(Mid(it, i, 1))
        c = Hex(a)
        c = IIf(Len(c) = 1, "0" & c, c)
        b = b & IIf(a > 65 And a < 120, Chr(a), ".")
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


