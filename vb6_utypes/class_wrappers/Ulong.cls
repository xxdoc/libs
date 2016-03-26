VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ulong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Enum op
'    op_add = 0
'    op_sub = 1
'    op_div = 2
'    op_mul = 3
'    op_mod = 4
'    op_xor = 5
'    op_and = 6
'    op_or = 7
'    op_rsh = 8
'    op_lsh = 9
'    op_gt = 10
'    op_lt = 11
'End Enum
'
'Enum modes
'    mUnsigned = 0
'    mSigned = 1
'    mHex = 2
'End Enum

Public use0x As Boolean
Public value As Long
Public padLeft As Boolean

'unsigned math operations
Private Declare Function Ulong Lib "utypes.dll" Alias "ULong" (ByVal v1 As Long, ByVal v2 As Long, ByVal operation As Long) As Long

'convert an unsigned long (or int) to unsigned string (vb6 hex and signed displays are fine so ommited..)
Private Declare Function U2Str Lib "utypes.dll" (ByVal v1 As Long, ByVal buf As String, ByVal cBufferSize As Long) As Long

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long


Property Get MAX_SIGNED() As String
     MAX_SIGNED = "7FFFFFFF"
     If use0x Then MAX_SIGNED = "0x" & MAX_SIGNED
End Property

Property Get MIN_SIGNED() As String
    MIN_SIGNED = "80000000"
    If use0x Then MIN_SIGNED = "0x" & MIN_SIGNED
End Property

Property Get MAX_UNSIGNED() As String
    MAX_UNSIGNED = "FFFFFFFF"
    If use0x Then MAX_UNSIGNED = "0x" & MAX_UNSIGNED
End Property

Property Get isNegAsSigned() As Boolean
    isNegAsSigned = (Left(CStr(value), 1) = "-")
End Property

'as a property, the IDE can show it results as a tooltip when debugging
Property Get sValue(Optional isHex As Boolean = True) As String
    Dim tmp As String
    
    If Not isHex Then
        sValue = GetUnsigned(value)
    Else
        tmp = Hex(value)
         
        If padLeft Then
            tmp = Right("00000000" & tmp, 8)
        End If
             
        If use0x Then tmp = "0x" & tmp
        
        sValue = tmp
    End If
    
End Property

'ok this is weird but works!
Property Let sValue(Optional isHex As Boolean = True, ByVal s As String)
    s = Trim(s)
    If LCase(Left(s, 2)) = "0x" Or LCase(Left(s, 2)) = "&h" Then
        isHex = True
        s = Mid(s, 3)
    End If
    If isHex Then
        value = CLng("&h" & s)
    Else
        value = CLng(s)
    End If
End Property

Private Function GetUnsigned(v As Long) As String
    Dim tmp As String, i As Long
    tmp = Space(64)
    i = U2Str(v, tmp, 64)
    If i > 0 Then GetUnsigned = Mid(tmp, 1, i)
End Function



'for signed math..you can use native vb6 operations..so mode is omitted..
Function DoOp(operand2, operation As op) As Ulong
    Dim ret As New Ulong
    Dim v2 As Long
    
    If TypeName(operand2) = "Ulong" Then
        v2 = operand2.value
    Else
        v2 = CLng(operand2)
    End If
    
    ret.value = Ulong(value, v2, operation)
    
    Set DoOp = ret
    
End Function
 


Private Sub Class_Initialize()
    hLib = LoadLibrary("UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\..\UTypes.dll")
    padLeft = True
End Sub




