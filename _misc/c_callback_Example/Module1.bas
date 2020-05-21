Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function myFunc Lib "dll2.dll" ( _
    ByVal callback As Long, _
    ByVal arg1 As Long _
) As Long

Sub Main()
    Dim i As Long
    i = myFunc(AddressOf myCallBack, 1)
    MsgBox i
End Sub

Function myCallBack(ByVal arg1 As Long) As Long
    'MsgBox "in callback " & arg1
    myCallBack = arg1 + 1
End Function

