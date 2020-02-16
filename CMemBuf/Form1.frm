VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   15225
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'as a 7 byte structure there will be an extra null pad after b for alignment while in memory
'this is different than when doing a put to disk where it gets packed i believe just be aware
Private Type udt1
    b As Byte
   ' b2 As Byte
    i As Integer
    l As Long
    s As String * 4 'fixed length strings are ok
    bb(3) As Byte   'fixed length arrays are ok
    'cc() As Byte    'will not work, only pointer saved in udt (can fool you if first udt is still in mem and pointer valid)
End Type

'note that udts with strings will fail, you will only get the pointer to the string.
'if you need this use my CMemStruct class

'output
    'udt dump: 000000    44 00 22 11 66 77 88 99                           .D.".fw..
    'encrypted udt: 000000    EA 8F 05 00 86 5D F5 BC                           ......]..
    'saving to: C:\Users\home\AppData\Local\Temp\xx.bin
    'loading from: C:\Users\home\AppData\Local\Temp\xx.bin
    'Reloaded/Decrypted/Rebuilt udt: 44 1122 99887766 abcd 22 dd
    '000000    22 11 33 44 55 66 22 11 44 33 66 55 88 77 AA    .".3DUf".D3fU.w.
    '000010   99 74 65 73 74 69 6E 67 21 00 04 00 74 68 69 73    .testing!...this
    '000020   02 00 69 73 02 00 6D 79 06 00 73 74 72 69 6E 67    ..is..my..string
    '000030   34 34                                              44
    '55443311
    'testing!
    'this
    'is
    'my
    'string
    '44

Sub d(X)
    Dim tmp() As String
    
    Debug.Print X
    
    tmp = Split(X, vbCrLf)
    For Each X In tmp
        If Len(X) > 0 Then List1.AddItem X
    Next
    
End Sub

Private Sub Form_Load()
    
    Dim mb As New CMemBuf
    Dim mb2 As New CMemBuf
    Dim i As Integer
    Dim l As Long
    Dim ii(4) As Integer
    Dim ss() As String
    Dim strStart As Long
    Dim u As udt1
    Dim u2 As udt1
    Dim emptyUDT As udt1
    Dim tmp As String
    Dim b() As Byte

    'mb.optBase0 = True
    'mb.optRaiseErr = True
    tmp = Environ("temp") & "\xx.bin"
    
'testing changing array bounds on load up
'    ReDim b(4)
'    For i = 0 To UBound(b)
'        b(i) = &H21 + i
'    Next
'    mb.Buffer = b
'    d mb.HexDump
'    mb.truncate 2
'    mb.write_ CInt(&H33)    'note pointer at start, int is two bytes
'    mb.write_ CByte(&H55), 4 'auto extends size
'    d mb.HexDump
'    mb.clear
'    Exit Sub
    
    
    u.b = &H44
    'u.b2 = &H33
    u.i = &H1122
    u.l = &H99887766
    u.s = "abcd"
    u.bb(0) = &H22
    u.bb(3) = &HDD
    
    'MsgBox LenB(u)
    'ReDim u.cc(100)
    'MsgBox LenB(u)
    'u.cc(0) = &HEE
    'u.cc(100) = &HCC
    
    mb.writeFromMem VarPtr(u), LenB(u)
    d "udt dump: " & mb.HexDump
    u = emptyUDT 'this makes sure that our redim u.cc is gone from mem
    
    mb.rc4 "test", True
    d "encrypted udt: " & mb.HexDump
    
    d "saving to: " & tmp
    mb.toFile tmp
    d "loading from: " & tmp
    mb2.fromFile tmp
    
    mb2.rc4 "test", True
    If Len(mb2.lastErr) Then d mb2.lastErr
    mb2.saveToMem VarPtr(u2), LenB(u2) ' Hex(u2.cc(0)), Hex(u2.cc(100) would give subscript out of range error...
    d "Reloaded/Decrypted/Rebuilt udt: " & Join(Array(Hex(u2.b), Hex(u2.i), Hex(u2.l), u2.s, Hex(u2.bb(0)), Hex(u2.bb(3))), " ")
    'd "aryisEmpty(u.cc) = " & mb2.AryIsEmpty(u.cc)
    
    Kill tmp
    mb2.clear
    mb.clear

    i = &H1122
    l = &H66554433
    
    ii(0) = &H1122
    ii(1) = &H3344
    ii(2) = &H5566
    ii(3) = &H7788
    ii(4) = &H99AA
    
    ss = Split("this is my string", " ")
    
    mb.write_ i
    mb.write_ l
    mb.write_ ii
    
    strStart = mb.curOffset
    mb.writeStr "testing!"
    mb.writeStr ss, , mb_bstr
    mb.writeStr 44, , mb_fixedSize
    
    d mb.HexDump
    
    mb.read l, 2
    d Hex(l)
    
    d mb.readStr(strStart)
    d mb.readStr(, mb_bstr)
    d mb.readStr(, mb_bstr)
    d mb.readStr(, mb_bstr)
    d mb.readStr(, mb_bstr)
    d mb.readStr(, mb_fixedSize, 2)
    d ""
    
    If Not mb.write_(Me) Then
        d "mb.write(object) failed - expected"
    End If
    
    
End Sub
