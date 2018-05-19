VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPrintF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetMem8 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any)

Private Const VT_I8 As Long = 20

Private Declare Function VariantChangeType Lib "oleaut32" ( _
    ByRef vargDest As Variant, _
    ByRef varSrc As Variant, _
    ByVal wFlags As Integer, _
    ByVal vt As VbVarType) As Long

Private Declare Function wvnsprintf Lib "shlwapi" Alias "wvnsprintfW" ( _
    ByVal lpOut As Long, _
    ByVal cchLimitIn As Long, _
    ByVal pszFmt As Long, _
    ByRef arglist As Long) As Long

Private mAbortOnErrors As Boolean

Public Property Get AbortOnErrors() As Boolean
    AbortOnErrors = mAbortOnErrors
End Property

Public Property Let AbortOnErrors(ByVal RHS As Boolean)
    mAbortOnErrors = RHS
End Property

Public Function CLong(ByVal Expr As Variant) As Long
    VariantChangeType CLong, Expr, 0, vbLong
End Function

Public Function CLongLong(ByVal Expr As Long) As Variant
    VariantChangeType CLongLong, Expr, 0, VT_I8
End Function

Public Function SPrintF( _
    ByVal MaxLen As Long, _
    ByRef Fmt As String, _
    ParamArray Args() As Variant) As String

    Dim ArgLongCount As Long
    Dim I As Long
    Dim arglist() As Long
    Dim J As Long
    Dim Result As Long

    For I = 0 To UBound(Args)
        Select Case VarType(Args(I))
            Case vbInteger, vbLong, vbSingle, vbString, vbBoolean
                ArgLongCount = ArgLongCount + 1
            Case vbDouble, vbCurrency, vbDate, VT_I8
                ArgLongCount = ArgLongCount + 2
            Case Else
                If AbortOnErrors Then
                    Err.Raise 13 'Type mismatch.
                Else
                    Exit Function
                End If
        End Select
    Next
    ReDim arglist(ArgLongCount - 1)
    For I = 0 To UBound(Args)
        Select Case VarType(Args(I))
            Case vbInteger, vbLong, vbSingle, vbBoolean
                arglist(J) = Args(I)
                J = J + 1
            Case vbString
                arglist(J) = StrPtr(Args(I))
                J = J + 1
            Case vbDouble, vbCurrency, vbDate, VT_I8
                GetMem8 ByVal &H80000000 Xor ((VarPtr(Args(I)) Xor &H80000000) + 8), arglist(J)
                J = J + 2
        End Select
    Next
    SPrintF = Space$(MaxLen)
    Result = wvnsprintf(StrPtr(SPrintF), MaxLen + 1, StrPtr(Fmt), arglist(0))
    If Result >= 0 Then
        SPrintF = Left$(SPrintF, Result)
    Else
        'MaxLen too small, bad Fmt string, etc.
        If AbortOnErrors Then
            Err.Raise &H80047700, , "The wvnsprintf() function failed"
        Else
            SPrintF = vbNullString
        End If
    End If
End Function

