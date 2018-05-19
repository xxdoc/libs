VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SPrintF"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5760
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Prt(ParamArray Text() As Variant)
    Dim I As Long

    With Text1
        .SelStart = &H7FFF
        For I = 0 To UBound(Text)
            .SelText = Text(I)
        Next
        .SelText = vbNewLine
    End With
End Sub

Private Sub Form_Load()
    Dim LongLong As Variant

    With New SPrintF
        'Unicode char (c) format which ignores the upper word::
        Prt "[", .SPrintF(10, "%c", 34), "]"
        Prt "[", .SPrintF(10, "%c", &HFFFF0141), "]" 'The "L with stroke," shows ANSI L in the TextBox.
        'ANSI char (C) format which ignores the upper 3 bytes:
        Prt "[", .SPrintF(10, "%C", &HABCD141), "]" 'Produces "A" (&H41).
        Prt
        'Strings:
        Prt "[", .SPrintF(10, "%s", "Test"), "]"
        Prt "[", .SPrintF(10, "%5s", "Test"), "]"
        Prt "[", .SPrintF(10, "%-5s", "Test"), "]"
        Prt "[", .SPrintF(10, "%.3s", "Test"), "]"
        Prt "[", .SPrintF(30, "%s, %s, and %s", "Test", "test", "testing"), "]"
        'Note that a vbNullChar will terminate strings early:
        Prt "[", .SPrintF(30, "%s, %s, and %s", "Test", "test", "test" & vbNullChar & "ing"), "]"
        'Percent (%) insertion after format spec:
        Prt "[", .SPrintF(10, "%s%%", "39.2"), "]"
        Prt
        '"Decimal" (as in not hex) integers:
        Prt "[", .SPrintF(10, "%+d %+d", -678, 678), "]"
        Prt "[", .SPrintF(10, "%d", 888), "]"
        Prt "[", .SPrintF(10, "%5d", 666), "]"
        Prt "[", .SPrintF(10, "%*d", 5, 111), "]"
        Prt "[", .SPrintF(10, "%-*d", 5, 222), "]"
        'These give leading 0's two different ways:
        Prt "[", .SPrintF(10, "%.5d", 777), "]"
        Prt "[", .SPrintF(10, "%0*d", 5, 333), "]"
        'These use "invisible +" i.e. a space for + and - for -:
        Prt "[", .SPrintF(10, "% d", -123), "]"
        Prt "[", .SPrintF(10, "% d", 123), "]"
        Prt "[", .SPrintF(10, "% *d", 7, -123), "]"
        Prt "[", .SPrintF(10, "% *d", 7, 123), "]"
        Prt "[", .SPrintF(10, "% -*d", 7, -123), "]"
        Prt "[", .SPrintF(10, "% -*d", 7, 123), "]"
        'Use the (i) instead of the (d), same result:
        Prt "[", .SPrintF(10, "% -*i", 7, 123), "]"
        Prt
        'Use a "sHort" i.e. Integer, 16-bit (h) format which ignores the upper word:
        Prt "[", .SPrintF(10, "%hd", &H7FFF0020), "]"
        Prt
        Prt "[", .SPrintF(10, "%X", &HEEDDCC), "]"
        Prt "[", .SPrintF(10, "%x", &HEEDDCC), "]"
        Prt "[", .SPrintF(10, "%8x", &HEEDDCC), "]"
        Prt "[", .SPrintF(10, "%-8x", &HEEDDCC), "]"
        Prt "[", .SPrintF(10, "%08x", &HEEDDCC), "]"
        'Pass two Long values but use a LongLong (ll - two lowercase L) format:
        Prt "[", .SPrintF(20, "%016llx", &HFFEEDDCC, 0), "]"
        Prt "[", .SPrintF(18, "&H%016llX", &H12345678, &HFFEEDDCC), "]"
        Prt
        'Pass a 64-bit LongLong, use the alternate 64-bit size (I64) format:
        LongLong = (.CLongLong(&H7FFFFFFF) + 1) * &H10000 * &H10000 _
                Or (.CLongLong(&HACBDCED) * &H10 Or &HF&)
        Prt "[", .SPrintF(22, "%I64x", LongLong), "]"
        Prt "[", .SPrintF(22, "%I64X", LongLong), "]"
        Prt
        'Pass some VB-native 64-bit types:
        Prt "[", .SPrintF(22, "%016I64x", 1.0001@), "]"
        Prt "[", .SPrintF(22, "%016I64x", 1#), "]"
        Prt "[", .SPrintF(22, "%016I64X", #1/31/2001 11:59:59 PM#), "]"
        Prt
        'Octal:
        Prt "[", .SPrintF(22, "%llo", &H12345678, &HFFEEDDCC), "]"
        Prt "[", .SPrintF(22, "%I64o", &H12345678, &HFFEEDDCC), "]"
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Text1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub
