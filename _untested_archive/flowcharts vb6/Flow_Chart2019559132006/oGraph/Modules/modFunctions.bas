Attribute VB_Name = "modFunctions"
Option Explicit
Private Const Pi = 3.14159265358979

'Taken from: C:\Program Files\Microsoft Visual Studio\MSDN98\98VSa\1033\office95.chm::/html/S11624.HTM
Public Function Arcsin(X#) As Double
    If Abs(X) = 1 Then
        Arcsin = X * 1.5707963267949
    Else
        Arcsin = Atn(X / Sqr(-X * X + 1))
    End If
End Function

Public Function Arccos(X#) As Double
    If X = -1 Then
        Arccos = 3.14159265359
    ElseIf X = 1 Then
        Arccos = 0
    Else
        Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
    End If
End Function

Public Function SuitableSide(ByVal cx As Double, ByVal cy As Double, ByVal W As Double, ByVal H As Double, ByVal X As Double, ByVal Y As Double) As Byte
Dim A As Currency
  If cx - X <> 0 Then
     A = Atn((cy - Y) / (cx - X))
  Else
     If cy > Y Then
        A = Pi / 2
     Else
        A = -Pi / 2
     End If
  End If

If cx >= X Then
   If cy < Y Then
      A = 2 * Pi + A
   End If
Else
      A = Pi + A
End If

If A < Atn(H / W) Then SuitableSide = 1
If A > Atn(H / W) Then SuitableSide = 2
If A > Pi - Atn(H / W) Then SuitableSide = 3
If A > Atn(H / W) + Pi Then SuitableSide = 4
If A > 2 * Pi - Atn(H / W) Then SuitableSide = 1

End Function

