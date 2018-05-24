VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   10290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bitFlags(31) As Long
'https://www.jameshbyrd.com/using-bitwise-operators-in-vb/

Private Sub Form_Load()
    Dim i As Long
    
    For i = 0 To UBound(bitFlags) - 1
        bitFlags(i) = 2 ^ i
        'Debug.Print i & "=" & Hex(bitFlags(i))
    Next

    bitFlags(31) = &H80000000 'vb would consider this an overflow if calc as above
    
    List1.AddItem "getBit(&H80000000, 31) = " & getBit(&H80000000, 31)
    List1.AddItem "setBit(&H80000000, 30) = " & toBinary(setBit(&H80000000, 30))
    List1.AddItem "clearBit(&H80000000, 30) = " & toBinary(clearBit(&H80000000, 30))

    List1.AddItem ""
    List1.AddItem "Dump flag/binary values: "
    List1.AddItem String(30, "-")

    For i = 0 To UBound(bitFlags)
        List1.AddItem Left(Hex(bitFlags(i)) & Space(10), 10) & " = " & toBinary(bitFlags(i))
    Next

    List1.AddItem Left(Hex(CLng(-1)) & Space(10), 10) & " = " & toBinary(CLng(-1))

    Dim v As Variant
    v = CInt(0)
    v = setBit(v, 9)
    Debug.Print "type: " & TypeName(v)
    Debug.Print toBinary(v) & " " & getBit(v, 9)
    
    
    
End Sub


Function getBit(value, bit) As Boolean

     If TypeName(value) = "Long" Then
        If bit > 31 Then Exit Function
        getBit = (CLng(value) And bitFlags(bit)) <> 0
     End If
     
     If TypeName(value) = "Integer" Then
        If bit > 15 Then Exit Function
        getBit = (CInt(value) And bitFlags(bit)) <> 0
     End If
        
     If TypeName(value) = "Byte" Then
        If bit > 7 Then Exit Function
        getBit = (CByte(value) And bitFlags(bit)) <> 0
     End If
     
End Function

'Function setBit(value, bit)
'    setBit = value Or bitFlags(bit)
'End Function
'
'Function clearBit(value, bit)
'    clearBit = value And (Not bitFlags(bit))
'End Function

Function setBit(value, bit)
    'validateBitSizeForVar bit, "setBit"
    'value = value Or bitFlags(bit) 'will change value to long
    If TypeName(value) = "Byte" Then setBit = CByte(value Or bitFlags(bit))
    If TypeName(value) = "Integer" Then setBit = CInt(value Or bitFlags(bit))
    If TypeName(value) = "Long" Then setBit = CLng(value Or bitFlags(bit))
End Function

Function clearBit(value, bit)
    'validateBitSizeForVar bit, "clearBit"
    'value = value And (Not bitFlags(bit)) 'will change value to long
    If TypeName(value) = "Byte" Then clearBit = CByte(value And (Not bitFlags(bit)))
    If TypeName(value) = "Integer" Then clearBit = CInt(value And (Not bitFlags(bit)))
    If TypeName(value) = "Long" Then clearBit = CLng(value And (Not bitFlags(bit)))
End Function


Function toBinary(value) As String
    
    Dim x
    Dim bits As Long
    Dim bytes As Long
    Dim l As Long
    Dim b As Byte
    Dim i As Integer
    
    If TypeName(value) = "Long" Then bytes = 4
    If TypeName(value) = "Integer" Then bytes = 2
    If TypeName(value) = "Byte" Then bytes = 1
    
    If bytes = 0 Then
        toBinary = TypeName(value) & " unsupported variable type"
        Exit Function
    End If
    
    v = CLng(value)
    bits = (bytes * 8) - 1
    
    For i = bits To 0 Step -1
        If i > 0 And i Mod 8 = 0 Then x = x & " "
        x = x & IIf((v And bitFlags(i)) <> 0, "1", "0")
    Next
    
    toBinary = x
    
End Function
