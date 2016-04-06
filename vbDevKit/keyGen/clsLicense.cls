VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private backUp As String

Sub Register(user, pass)
    Dim mypass As String, i As Integer, l As String, u As Single
    mypass = CalcPass(user)
    If Len(mypass) <> Len(pass) Then GoTo failed
    If Len(user) <> Len(mypass) Then GoTo failed
    u = Len(mypass) + 1
    For i = 1 To u - 1
        l = Mid(mypass, u - i, 1)
        'Debug.Print u - i & " " & l & " =? " & Mid(pass, i, 1)
        If l <> Mid(pass, i, 1) Then GoTo failed
        If i = Len(mypass) Then
            MsgBox "Regsitration Successful"
            Exit Sub
        End If
    Next
    
failed:
    MsgBox "Failed"
End Sub

Function CalcPass(user) As String
    Dim b() As Byte, ret As String, i As Integer
        
    If Len(user) < 8 Then Err.Raise 1
    If Len(user) > 8 Then Err.Raise 2
    
    b() = StrConv(user, vbFromUnicode)
    For i = 0 To UBound(b)
        b(i) = b(i) Xor i + 20
        If b(i) < 32 Then b(i) = b(i) + 32
        If b(i) > 97 Then b(i) = 97 - i
        ret = ret & hexit(b(i), ((i + 1) ^ (i + 1)))
    Next

    ret = Mid(ret, 4, 4) & Mid(ret, 9, 4)
    CalcPass = Trim(ret)
    backUp = Trim(ret)
    
    'InputBox "", , ret
End Function

Private Function hexit(n, i) As String
    hexit = Hex(n + i)
    If Len(hexit) = 0 Then hexit = "0" & hexit
End Function

