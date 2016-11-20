VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCrc32Test 
      Caption         =   "Crc32"
      Height          =   675
      Left            =   8520
      TabIndex        =   5
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdUByte 
      Caption         =   "UByte Tests"
      Height          =   585
      Left            =   8520
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdUIntTest 
      Caption         =   "UInt Tests"
      Height          =   555
      Left            =   8520
      TabIndex        =   3
      Top             =   1710
      Width           =   2625
   End
   Begin VB.CommandButton cmdUlong 
      Caption         =   "ULong Tests"
      Height          =   555
      Left            =   8475
      TabIndex        =   2
      Top             =   855
      Width           =   2715
   End
   Begin VB.CommandButton cmdx64Test 
      Caption         =   "64 bit tests"
      Height          =   510
      Left            =   8475
      TabIndex        =   1
      Top             =   90
      Width           =   2715
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   8025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'just a useful bonus method to add into the C dll might as well...
Private Declare Function crc32 Lib "utypes.dll" (ByRef b As Byte, ByVal sz As Long) As Long
Private Declare Function crc32w Lib "utypes.dll" (ByVal b As Long, ByVal sz As Long) As Long

Function crc(buf As String) As String

    Dim v As Long
    Dim l As Long
    
    Dim tmp As New UByte 'this is only to make sure the dll is loaded..
    tmp.use0x = True     'class_init runs now..
    
    'Dim b() As Byte
    'b() = StrConv(buf, vbFromUnicode, &H409)
    'v = crc32(b(0), UBound(b) + 1)
    
    l = Len(buf)
    If l = 0 Then Exit Function
    v = crc32w(StrPtr(buf), l)
    crc = Hex(v)
    
End Function

Private Sub cmdCrc32Test_Click()
    Dim a As String
    a = String(2000, "A")
    List1.Clear
    List1.AddItem "crc32(String(2000, 'A')) = " & crc(a) & " =? FFCC0057"
    'Clipboard.Clear
    'Clipboard.SetText a
End Sub

Private Sub cmdUByte_Click()
    Dim a As New UByte
    
    List1.Clear
    a = 255
    List1.AddItem "255+1 = " & a.add(1)
    
    a = 0
    List1.AddItem "0-1 = " & a.subtract(1)
    
    
End Sub

Private Sub cmdUIntTest_Click()
    
    Dim a As New UInt
    Dim b As New UInt
    
    List1.Clear
    
    a.use0x = True
    a = a.MAX_SIGNED
    List1.AddItem "start value: (MAX_SIGNED) " & a
    a = a.add(1)
    List1.AddItem "Max+1 signed = " & a
    List1.AddItem "Max+1 unsigned = " & a.toString(False)
    
    If b.fromString(a.toString(False)) Then
        List1.AddItem "Max+1 signed string transfer = " & b
    Else
        List1.AddItem "string transfer failed..."
    End If
        
    
End Sub

Private Sub cmdUlong_Click()
    
    Dim a As New ULong
    Dim b As New ULong
    
    List1.Clear
    a = 60
    Set b = a.rshift(1)
    List1.AddItem "60 >> 1 = " & b & " = 30?"
    List1.AddItem String(10, "-")
    
    a = 60
    a = a.rshift(2).Value
    List1.AddItem "60 >> 2 = " & a & " = 15?"
    List1.AddItem String(10, "-")
    
    a = 7
    a = a.lshift(3)
    List1.AddItem "7 << 3 = " & a & " = 56?"
    List1.AddItem String(10, "-")
    
    a.fromString "0x11223344"
    List1.AddItem "0x11223344 =? " & a.toString
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_SIGNED
    b = a.add(2)
    List1.AddItem "MAX_SIGNED = " & a
    List1.AddItem "MAX_SIGNED + 2 = " & b.toString(False)
    List1.AddItem "MAX_SIGNED + 2 as unsigned = " & b
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_SIGNED
    Set b = a.add(b)
    List1.AddItem "MaxSigned + maxSigned+2 = " & b & " = 0?"
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_UNSIGNED
    Set b = a.add(1)
    List1.AddItem "Max UNSigned + 1 = " & b.Value & " = 0?"
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_UNSIGNED
    List1.AddItem "Max UNSigned + 1 inline = " & a.add(1).toString
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_SIGNED
    a.Value = a.add(1).Value
    b.Value = 0
    List1.AddItem "MAX_SIGNED+1 = " & a.Value & " (native signed value)"
    List1.AddItem "MAX_SIGNED+1 > 0 signed ? " & (a.Value > b.Value) & " (native cmp)"

    List1.AddItem "MAX_SIGNED+1 unsigned = " & a.fromString(False)
    List1.AddItem "MAX_SIGNED+1 > 0 unsigned ? " & a.greaterThan(b)
    
    
    
    
End Sub

Private Sub cmdx64Test_Click()

    Dim a As New ULong64
    Dim b As ULong64
    
    List1.Clear
    a.fromString -1
    List1.AddItem "-1 = " & a.toString(msigned)
    
    a = 0
    a = a.subtract(1)
    List1.AddItem "0-1 signed = " & a.toString(msigned) & " =? -1"
    List1.AddItem "0-1 unsigned = " & a.toString(munSigned) & " =? 18446744073709551615"
    List1.AddItem "0-1 hex = " & a.toString(mhex) & " =?  FFFFFFFFFFFFFFFF"
    List1.AddItem String(10, "-")
    
    a.SetLongs &H11223344, &H55667788
    List1.AddItem "setlong(&H11223344, &H55667788) hi=" & Hex(a.hi) & " lo=" & Hex(a.lo)
    a.hi = &H88776655
    a.lo = &H44332211
    List1.AddItem "hi=" & Hex(a.hi) & " lo=" & Hex(a.lo) & " -> " & a.toString
    List1.AddItem String(10, "-")
    
    a.SetLongs 0, 1
    Set b = a.lshift(60)
    List1.AddItem "1 << 60 (unsigned) = " & b.toString(munSigned) & " =? 1152921504606846976"
    List1.AddItem "hex: " & b.toString & " =? 1000000000000000"
    List1.AddItem String(10, "-")
    
    a.fromString a.MAX_UNSIGNED64
    Set b = a.add(1)
    List1.AddItem "MAX_UNSIGNED64 + 1 = " & b.toString(munSigned)
    List1.AddItem String(10, "-")

    a.fromString a.MAX_SIGNED64
    Set b = a.add(1)
    List1.AddItem "MAX_SIGNED64 + 1 = " & b.toString(msigned)
    List1.AddItem "isNegBitSet = " & b.isNegBitSet
    List1.AddItem "a > b " & (a.Value > b.Value)
    List1.AddItem "a < b " & (a.Value < b.Value)
    List1.AddItem "unsigned a > b ? " & a.greaterThan(b)
    List1.AddItem String(10, "-")
    
    a.fromString a.MIN_SIGNED64
    Set b = a.subtract(1)
    List1.AddItem "MIN_SIGNED64 - 1 = " & b.toString(msigned)
    List1.AddItem "native a > b ? " & (a.Value > b.Value)
    List1.AddItem "native a < b ? " & (a.Value < b.Value)
    List1.AddItem "unsigned  a < b ? " & a.lessThan(b)
    List1.AddItem String(10, "-")

    a.fromString "8877665544332211"
    List1.AddItem "a.fromString '8877665544332211' = " & a.toString
    List1.AddItem String(10, "-")

    a.fromString a.MAX_UNSIGNED64
    List1.AddItem "a.fromString a.MAX_UNSIGNED64 = MAX_UNSIGNED64 ? " & (a.toString = a.MAX_UNSIGNED64)
    List1.AddItem String(10, "-")
    
    a.SetLongs 0, 32
    List1.AddItem a.toString(msigned) & " is 32bit safe? " & a.is32BitSafe
    List1.AddItem String(10, "-")

    a.SetLongs 1, 32
    a.useTick = True
    a.padLeft = True
    List1.AddItem a.toString & " is 32bit safe? " & a.is32BitSafe
    List1.AddItem String(10, "-")

    a.SetLongs 0, 32
    a.useTick = False
    a.padLeft = False
    a.use0x = True
    List1.AddItem a.toString
    List1.AddItem String(10, "-")

    'this doesnt work
    'a.value = &H12345678
    'List1.AddItem "&H12345678 = " & a.sValue()
    
    'it might be safer to use an explicit C function for compares..this works..
    'a.SetLongs &H81223344, &H55667788 + 1
    'b.SetLongs &H81223344, &H55667788 - 1
    '
    'Debug.Print a.value < b.value
    
    
    'this doesnt
    'a.SetLongs 0, 1
    'b.SetLongs 0, -1
    '
    'Debug.Print a.value > b.value
    
    'works...
    'a.svalue =a.MAX_SIGNED64
    'b.svalue =b.MIN_SIGNED64
    '
    'Debug.Print a.value > b.value
    
    'equals should be safe..
        

End Sub

 
Private Sub Command1_Click()

End Sub
