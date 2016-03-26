VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUlong 
      Caption         =   "ULong Tests"
      Height          =   555
      Left            =   4905
      TabIndex        =   2
      Top             =   855
      Width           =   2715
   End
   Begin VB.CommandButton cmdx64Test 
      Caption         =   "64 bit tests"
      Height          =   510
      Left            =   4905
      TabIndex        =   1
      Top             =   90
      Width           =   2715
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUlong_Click()
    
    Dim a As New Ulong
    Dim b As New Ulong
    
    List1.Clear
    a.value = 60
    Set b = a.DoOp(1, op_rsh)
    List1.AddItem "60 >> 1 = " & b.value
    
    a.value = 60
    a.value = a.DoOp(2, op_rsh).value
    List1.AddItem "60 >> 2 = " & a.value
        
    a.value = 7
    a.value = a.DoOp(3, op_lsh).value
    List1.AddItem "7 << 3 = " & a.value

    a.sValue = "0x11223344"
    List1.AddItem a.sValue
    
    a.sValue = a.MAX_SIGNED
    Set b = a.DoOp(2, op_add)
    List1.AddItem a.value
    List1.AddItem b.sValue(False)
    List1.AddItem b.value
    
    a.sValue = a.MAX_SIGNED
    Set b = a.DoOp(b, op_add)
    List1.AddItem b.value
    
    a.sValue = a.MAX_UNSIGNED
    Set b = a.DoOp(1, op_add)
    List1.AddItem b.value
    
    a.sValue = a.MAX_UNSIGNED
    List1.AddItem a.DoOp(1, op_add).sValue
    
    a.sValue = a.MAX_SIGNED
    a.value = a.DoOp(1, op_add).value
    b.value = 0
    List1.AddItem "MAX_SIGNED+1 = " & a.value & " (native signed value)"
    List1.AddItem "MAX_SIGNED+1 > 0 signed ? " & (a.value > b.value) & " (native cmp)"
    
    List1.AddItem "MAX_SIGNED+1 unsigned = " & a.sValue(False)
    List1.AddItem "MAX_SIGNED+1 > 0 unsigned ? " & CBool(a.DoOp(b, op_gt).value)
    
    
    
    
End Sub

Private Sub cmdx64Test_Click()

    Dim a As New ULong64
    Dim b As ULong64
    
    List1.Clear
    
    a.SetLongs &H11223344, &H55667788
    List1.AddItem Hex(a.hi) & " " & Hex(a.lo)
    a.hi = &H88776655
    a.lo = &H44332211
    List1.AddItem Hex(a.hi) & " " & Hex(a.lo)
    
    a.SetLongs 0, 1
    Set b = a.DoOp(60, op_lsh)
    List1.AddItem "1 << 60 = " & b.sValue
    
    a.sValue = a.MAX_UNSIGNED64
    Set b = a.DoOp(1, op_add)
    List1.AddItem "MAX_UNSIGNED64 + 1 = " & b.sValue(mUnsigned)

    a.sValue = a.MAX_SIGNED64
    Set b = a.DoOp(1, op_add)
    List1.AddItem "MAX_SIGNED64 + 1 = " & b.sValue(mSigned)
    List1.AddItem "isNegAsSigned = " & b.isNegAsSigned
    List1.AddItem "a > b " & (a.value > b.value)
    List1.AddItem "a < b " & (a.value < b.value)

    a.sValue = a.MIN_SIGNED64
    Set b = a.DoOp(1, op_sub)
    List1.AddItem "MIN_SIGNED64 - 1 = " & b.sValue(mSigned)
    List1.AddItem (a.value > b.value)
    List1.AddItem (a.value < b.value)

    a.SetLongs &H11223344, &H55667788
    List1.AddItem a.sValue

    a.sValue = "8877665544332211"
    List1.AddItem a.sValue

    a.sValue = a.MAX_UNSIGNED64
    List1.AddItem (a.sValue = a.MAX_UNSIGNED64)

    a.SetLongs 0, 32
    List1.AddItem a.is32BitSafe
    List1.AddItem a.sValue

    a.SetLongs 1, 32
    a.useTick = True
    a.padLeft = True
    List1.AddItem a.sValue

    a.SetLongs 0, 32
    a.useTick = False
    a.padLeft = False
    a.use0x = True
    List1.AddItem a.sValue

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

 
