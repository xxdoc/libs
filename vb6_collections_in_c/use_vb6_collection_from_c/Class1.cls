VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private myCol As New Collection

Function getCol() As Collection
    Set getCol = myCol
End Function

Function test()
    On Error Resume Next
    
    Dim c
    MsgBox "Count: " & myCol.Count
    
    For Each c In myCol
        MsgBox TypeName(c) & " = " & CStr(c)
        If Err.Number <> 0 Then
            MsgBox Err.Description
            Err.Clear
        End If
    Next
    
End Function
