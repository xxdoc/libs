VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ULong64"
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
'End Enum
'
'Enum modes
'    mUnsigned = 0
'    mSigned = 1
'    mHex = 2
'End Enum

Public mode As modes
Public value As Currency
Public padLeft As Boolean
Public useTick As Boolean
Public use0x As Boolean

'Currency (scaled integer) 8 bytes
'-922,337,203,685,477.5808 to 922,337,203,685,477.5807

Dim hLib As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'unsigned math operations
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

Property Get MAX_SIGNED64() As String
     If useTick Then
        MAX_SIGNED64 = "7FFFFFFF`FFFFFFFF"
     Else
        MAX_SIGNED64 = "7FFFFFFFFFFFFFFF"
     End If
     If use0x Then MAX_SIGNED64 = "0x" & MAX_SIGNED64
End Property

Property Get MIN_SIGNED64() As String
    If useTick Then
        MIN_SIGNED64 = "80000000`00000000"
    Else
        MIN_SIGNED64 = "8000000000000000"
    End If
    If use0x Then MIN_SIGNED64 = "0x" & MIN_SIGNED64
End Property

Property Get MAX_UNSIGNED64() As String
    If useTick Then
        MAX_UNSIGNED64 = "FFFFFFFF`FFFFFFFF"
    Else
        MAX_UNSIGNED64 = "FFFFFFFFFFFFFFFF"
    End If
    If use0x Then MAX_UNSIGNED64 = "0x" & MAX_UNSIGNED64
End Property

Property Get is32BitSafe() As Boolean
    Dim hi As Long, lo As Long
    GetLongs hi, lo
    If hi = 0 Then is32BitSafe = True
End Property

Property Get isNegAsSigned() As Boolean
    isNegAsSigned = (Left(sValue(mSigned), 1) = "-")
End Property

'as a property, the IDE can show it results as a tooltip when debugging
Property Get sValue(Optional m As modes = mHex) As String
    Dim tmp As String, i As Long, lo As String, hi As String, a As Long
    tmp = Space(64)
    i = U642Str(value, tmp, 64, m)
    If i > 0 Then
        tmp = UCase(Mid(tmp, 1, i))
        If m = mHex Then
            If padLeft Then
                tmp = Right("0000000000000000" & tmp, 16)
            End If
            If useTick Then
                a = Len(tmp) - 8
                If a > 0 Then
                    lo = Right(tmp, 8)
                    hi = Mid(tmp, 1, Len(tmp) - 8)
                    tmp = hi & "`" & lo
                End If
            End If
            If use0x Then tmp = "0x" & tmp
        End If
        sValue = tmp
    End If
End Property

Property Let sValue(Optional m As modes = mHex, ByVal s As String)
    s = Trim(s)
    s = Replace(s, "`", Empty)
    If LCase(Left(s, 2)) = "0x" Then
        m = mHex
        s = Mid(s, 3)
    End If
    value = Str264(s, IIf(m = mHex, 16, 10))
End Property

Sub SetLongs(hi As Long, lo As Long)
    value = toU64(hi, lo)
End Sub

Sub GetLongs(ByRef hi As Long, ByRef lo As Long)
    fromU64 value, hi, lo
End Sub

Function DoOp(operand2, operation As op) As ULong64
    Dim ret As New ULong64
    Dim tmp As New ULong64
    
    If TypeName(operand2) <> "ULong64" Then
        tmp.SetLongs 0, CLng(operand2)
    Else
        Set tmp = operand2
    End If
    
    If mode = mSigned Then
        ret.value = S64(value, tmp.value, operation)
    Else
        ret.value = U64(value, tmp.value, operation)
    End If
    
    Set DoOp = ret
    
End Function


Private Sub Class_Initialize()
    hLib = LoadLibrary("UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\UTypes.dll")
    If hLib = 0 Then hLib = LoadLibrary(App.Path & "\..\..\UTypes.dll")
    padLeft = True
End Sub

