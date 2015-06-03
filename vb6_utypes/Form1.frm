VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Native Tests"
      Height          =   330
      Left            =   3195
      TabIndex        =   1
      Top             =   3375
      Width           =   1140
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum op
    op_add = 0
    op_sub = 1
    op_div = 2
    op_mul = 3
    op_mod = 4
    op_xor = 5
    op_and = 6
    op_or = 7
End Enum

Enum modes
    mUnsigned = 0
    mSigned = 1
    mHex = 2
End Enum

'unsigned math operations
Private Declare Function ULong Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByVal operation As Long) As Long
Private Declare Function UInt Lib "utypes.dll" (ByVal v1 As Integer, ByVal v2 As Integer, ByVal operation As Long) As Integer
Private Declare Function U64 Lib "utypes.dll" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal operation As op) As Currency

'signed math for 64 bit numbers (necessary?)
Private Declare Function S64 Lib "utypes.dll" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal operation As op) As Currency

'create 64 bit number from hi and lo longs
Private Declare Function toU64 Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long) As Currency

'create a 64 bit number from a string in specified base (16 for a hex string)
Private Declare Function Str264 Lib "utypes.dll" (ByVal s As String, Optional ByVal base As Long = 10) As Currency

'convert a 64 bit number to string in specified format
Private Declare Function U642Str Lib "utypes.dll" (ByVal v1 As Currency, ByVal buf As String, ByVal cBufferSize As Long, ByVal mode As modes) As Long

'get hi and lo longs from 64 numbers
Private Declare Sub fromU64 Lib "utypes.dll" (ByVal v0 As Currency, ByRef v1 As Long, ByRef v2 As Long)

'convert an unsigned long (or int) to unsigned string (vb6 hex and signed displays are fine so ommited..)
Private Declare Function U2Str Lib "utypes.dll" (ByVal v1 As Long, ByVal buf As String, ByVal cBufferSize As Long) As Long



Private Sub Command1_Click()
    On Error Resume Next
    Dim a As Long
    Dim b As Integer
    
    a = 2147483647 + 1
    MsgBox Hex(2147483647) & " + 1 = " & Hex(a) & " Error: " & Err.Description
    Err.Clear
    
    a = -2147483648# - 1
    MsgBox Hex(-2147483648#) & " - 1 = " & Hex(a) & " Error: " & Err.Description
    Err.Clear
     
    b = 32767 + 1
    MsgBox Hex(32767) & " + 1 = " & Hex(b) & " Error: " & Err.Description
    Err.Clear
    
    b = -32768 - 1
    MsgBox Hex(-32768) & " - 1 = " & Hex(b) & " Error: " & Err.Description
    Err.Clear
     
End Sub

Private Sub Form_Load()

    testLong 2147483647, 1, op_add
    testLong -2147483648#, 1, op_sub
    
    testInt 32767, 1, op_add
    testInt -32768, 1, op_sub
    
    Dim d As Currency
    d = toU64(&HAAAAAAAA, &HBBBBBBBB)
    List1.AddItem Get64(d, mHex)
    List1.AddItem Get64(d, mUnsigned)
    
    d = U64(toU64(&HCCCCCCCC, 0), toU64(0, &HDDDDDDDD), op_add)
    List1.AddItem Get64(d, mHex)
    
    Dim l As Long, hi As Long, lo As Long
    l = ULong(2147483647, 1, op_add)
    List1.AddItem "Unsigned: " & GetUnsigned(l) & " hex:" & Hex(l) & " signed:" & l
    
    d = Str264("1122334455667788", 16)
    List1.AddItem Get64(d, mHex)
    
    fromU64 d, hi, lo
    List1.AddItem Hex(hi) & " " & Hex(lo)
    
    d = Str264("2147483648") 'max signed long + 1
    fromU64 d, hi, lo
    List1.AddItem "hi: " & Hex(hi) & " lo: " & Hex(lo)
    
    d = S64(Str264("-1"), Str264("1"), op_sub)
    List1.AddItem Get64(d, mSigned) & " unsigned: " & Get64(d, mUnsigned)
    
    d = U64(Str264("-1"), Str264("1"), op_sub)
    List1.AddItem Get64(d, mUnsigned)
    
End Sub

Function Get64(v As Currency, Optional m As modes = mHex) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U642Str(v, tmp, 64, m)
    If i > 0 Then Get64 = Mid(tmp, 1, i)
End Function

Function GetUnsigned(v As Long) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U2Str(v, tmp, 64)
    If i > 0 Then GetUnsigned = Mid(tmp, 1, i)
End Function


Sub testLong(a As Long, b As Long, opp As op)
    
    Dim ret As Long, o As Variant, msg As String
    o = Array("+", "-", "/", "*", "mod", "xor", "and", "or")
    
    ret = ULong(a, b, opp)
    msg = Hex(a) & " " & o(opp) & " " & Hex(b) & " = " & Hex(ret)
    
    List1.AddItem msg
    Debug.Print msg
End Sub

Sub testInt(a As Integer, b As Integer, opp As op)
    
    Dim ret As Integer, o As Variant, msg As String
    o = Array("+", "-", "/", "*", "mod", "xor", "and", "or")
    
    ret = UInt(a, b, opp)
    msg = Hex(a) & " " & o(opp) & " " & Hex(b) & " = " & Hex(ret)
    
    List1.AddItem msg
    Debug.Print msg
    
End Sub

