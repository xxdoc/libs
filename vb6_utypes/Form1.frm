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
      Left            =   5760
      TabIndex        =   1
      Top             =   3420
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
      Left            =   405
      TabIndex        =   0
      Top             =   180
      Width           =   7125
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
    mhex = 2
End Enum


Private Declare Function ULong Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long, ByVal operation As Long) As Long
Private Declare Function UInt Lib "utypes.dll" (ByVal v1 As Integer, ByVal v2 As Integer, ByVal operation As Long) As Integer

Private Declare Function toU64 Lib "utypes.dll" (ByVal v1 As Long, ByVal v2 As Long) As Currency

'int __stdcall U642Str(unsigned __int64 v1, LPSTR pszString, LONG cSize, int mode){
Private Declare Function U642Str Lib "utypes.dll" (ByVal v1 As Currency, ByVal buf As String, ByVal cBufferSize As Long, ByVal mode As modes) As Long
Private Declare Function U2Str Lib "utypes.dll" (ByVal v1 As Long, ByVal buf As String, ByVal cBufferSize As Long, ByVal mode As modes) As Long


'unsigned __int64 __stdcall U64(unsigned __int64 v1, unsigned __int64 v2, int operation){
Private Declare Function U64 Lib "utypes.dll" (ByVal v1 As Currency, ByVal v2 As Currency, ByVal operation As op) As Currency


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
    List1.AddItem Getx64(d, mhex)
    List1.AddItem Getx64(d, mUnsigned)
    
    d = U64(toU64(&HCCCCCCCC, 0), toU64(0, &HDDDDDDDD), op_add)
    List1.AddItem Getx64(d, mhex)
    
    Dim l As Long
    l = ULong(2147483647, 1, op_add)
    List1.AddItem GetUnsigned(l)
        
End Sub

Function Getx64(v As Currency, Optional m As modes = mhex) As String
    Dim tmp As String
    tmp = Space(64)
    U642Str v, tmp, 64, m
    Getx64 = tmp
End Function

Function GetUnsigned(v As Long, Optional m As modes = mUnsigned) As String
    Dim tmp As String
    tmp = Space(64)
    U2Str v, tmp, 64, m
    GetUnsigned = tmp
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

