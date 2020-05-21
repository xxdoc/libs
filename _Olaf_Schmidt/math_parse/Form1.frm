VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'just a group of Test-calls for the simple Evaluator (in comparison to VB-outputs)
'the VB6-resolved expression is always located directly below the Eval-String to
'be able to compare the test-expressions more easily ...
'(results are printed side-by-side and should come out the same in all test-cases)

Private Sub Form_Load()

   Debug.Print tEval("0x20+ 5*4")
   Exit Sub

  'simple operator-precedence without parentheses
  Debug.Print Eval("3 + 5 * 9 + 7 + 2 * 5"), _
                    3 + 5 * 9 + 7 + 2 * 5
                    
  Debug.Print Eval("1 + 6 / 3 - 7"), _
                    1 + 6 / 3 - 7

  'unary-operator test
  Debug.Print Eval("-1 + -6 / 3 - -7 "), _
                    -1 + -6 / 3 - -7
                    
  'simple parentheses test
  Debug.Print Eval("-14 / 7 * -(1 + 2)"), _
                    -14 / 7 * -(1 + 2)

  'a complex case, including exponent-handling
  Debug.Print Eval("((1 + -2) * -3 + 4) * 2 / 7 * 216 ^ (-1 / -3) "), _
                    ((1 + -2) * -3 + 4) * 2 / 7 * 216 ^ (-1 / -3)

  'operator-precedence (mainly to test Mod and Div operators)
  Debug.Print Eval("27 / 3 Mod (5 \ 2) + 23"), _
                    27 / 3 Mod (5 \ 2) + 23

  'function-calls
  Debug.Print Eval("43 + -(-2 - 3) * Abs(Cos(4 * Atn(1)))"), _
                    43 + -(-2 - 3) * Abs(Cos(4 * Atn(1)))
  
  'simple case, but mixed with a string-concat (math-ops have precedence)
  Debug.Print Eval("5 + 3 & 2"), _
                    5 + 3 & 2
                    
  'simple case of a string-concat (notation for string-literals as in SQL)
  Debug.Print Eval("'abc ' & '123 ' & 'xyz'"), _
                    "abc " & "123 " & "xyz"
  
  'simple comparison-ops follow... (starting with a string-comparison)
  Debug.Print Eval("'abc' > '123'"), _
                    "abc" > "123"
                    
  'comparison of numbers... math-ops have precedence
  Debug.Print Eval("3 = 1 + 2"), _
                    3 = 1 + 2
                    
  'comparison of strings per Like-Operator...
  Debug.Print Eval("'abc' Like '*b*'"), _
                    "abc" Like "*b*"
             
  'comparison of strings per Like-Operator (using the "in-range" notation)
  Debug.Print Eval("'3xB..foo' Like '[1-5]?[A-C]*o'"), _
                    "3xB..foo" Like "[1-5]?[A-C]*o"
  
  'and here logical comparisons, involving the And-Operator (a mix of String- and Value-Compares)
  Debug.Print Eval("'foobar' Like 'foo*' And 'foobar' Like '*bar' And (1 + 2) * 4 - 1 = 11"), _
                    "foobar" Like "foo*" And "foobar" Like "*bar" And (1 + 2) * 4 - 1 = 11
End Sub
 
'tiny version ------------------------------------------
 Public Function tEval(ByVal Expr As String) As Double
  Dim L As String, R As String
  Expr = Replace(Expr, "0x", "&h")
  If tSpl(Expr, "+", L, R) Then tEval = tEval(L) + tEval(R): Exit Function
  If tSpl(Expr, "-", L, R) Then tEval = tEval(L) - tEval(R): Exit Function
  If tSpl(Expr, "*", L, R) Then tEval = tEval(L) * tEval(R): Exit Function
  If Len(Expr) Then tEval = Val(Expr)
End Function

Private Function tSpl(Expr As String, Op$, L$, R$) As Long
  tSpl = InStrRev(Expr, Op)
  If tSpl Then R = Mid$(Expr, tSpl + Len(Op)): L = Left$(Expr, tSpl - 1)
End Function
'-----------------------------------------------------------

Public Function Eval(ByVal Expr As String)
Dim L As String, R As String
  Do While HandleParentheses(Expr): Loop

  If 0 Then
    ElseIf Spl(Expr, "Or", L, R) Then:   Eval = Eval(L) Or Eval(R)
    ElseIf Spl(Expr, "And", L, R) Then:  Eval = Eval(L) And Eval(R)
    ElseIf Spl(Expr, ">=", L, R) Then:   Eval = Eval(L) >= Eval(R)
    ElseIf Spl(Expr, "<=", L, R) Then:   Eval = Eval(L) <= Eval(R)
    ElseIf Spl(Expr, "=", L, R) Then:    Eval = Eval(L) = Eval(R)
    ElseIf Spl(Expr, ">", L, R) Then:    Eval = Eval(L) > Eval(R)
    ElseIf Spl(Expr, "<", L, R) Then:    Eval = Eval(L) < Eval(R)
    ElseIf Spl(Expr, "Like", L, R) Then: Eval = Eval(L) Like Eval(R)
    ElseIf Spl(Expr, "&", L, R) Then:    Eval = Eval(L) & Eval(R)
    ElseIf Spl(Expr, "-", L, R) Then:    Eval = Eval(L) - Eval(R)
    ElseIf Spl(Expr, "+", L, R) Then:    Eval = Eval(L) + Eval(R)
    ElseIf Spl(Expr, "Mod", L, R) Then:  Eval = Eval(L) Mod Eval(R)
    ElseIf Spl(Expr, "\", L, R) Then:    Eval = Eval(L) \ Eval(R)
    ElseIf Spl(Expr, "*", L, R) Then:    Eval = Eval(L) * Eval(R)
    ElseIf Spl(Expr, "/", L, R) Then:    Eval = Eval(L) / Eval(R)
    ElseIf Spl(Expr, "^", L, R) Then:    Eval = Eval(L) ^ Eval(R)
    ElseIf Trim(Expr) >= "A" Then:       Eval = Fnc(Expr)
    ElseIf Len(Expr) Then:               Eval = IIf(InStr(Expr, "'"), _
                            Replace(Trim(Expr), "'", ""), Val(Expr))
  End If
End Function

Private Function HandleParentheses(Expr As String) As Boolean
Dim P As Long, i As Long, C As Long
  P = InStr(Expr, "(")
  If P Then HandleParentheses = True Else Exit Function

  For i = P To Len(Expr)
    If Mid(Expr, i, 1) = "(" Then C = C + 1
    If Mid(Expr, i, 1) = ")" Then C = C - 1
    If C = 0 Then Exit For
  Next i

  Expr = Left(Expr, P - 1) & Str(Eval(Mid(Expr, P + 1, i - P - 1))) & Mid(Expr, i + 1)
End Function

Private Function Spl(Expr As String, Op$, L$, R$) As Boolean
Dim P As Long
  P = InStrRev(Expr, Op, , 1)
  If P Then Spl = True Else Exit Function
  If P < InStrRev(Expr, "'") And InStr("*-", Op) Then P = InStrRev(Expr, "'", P) - 1

  R = Mid(Expr, P + Len(Op))
  L = Trim(Left$(Expr, IIf(P > 0, P - 1, 0)))

  Select Case Right(L, 1)
    Case "", "+", "*", "/", "A" To "z": Spl = False
    Case "-": R = "-" & R
  End Select
End Function

Private Function Fnc(Expr As String)
  Expr = LCase(Trim(Expr))

  Select Case Left(Expr, 3)
    Case "abs": Fnc = Abs(Val(Mid$(Expr, 4)))
    Case "sin": Fnc = Sin(Val(Mid$(Expr, 4)))
    Case "cos": Fnc = Cos(Val(Mid$(Expr, 4)))
    Case "atn": Fnc = Atn(Val(Mid$(Expr, 4)))
    Case "log": Fnc = Log(Val(Mid$(Expr, 4)))
    Case "exp": Fnc = Exp(Val(Mid$(Expr, 4)))
    'etc...
  End Select
End Function

