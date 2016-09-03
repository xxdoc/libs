VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim c As New CollectionEx

    c.Add "test"
    c.Add "test2", "mykey"
    c.ext.changeKeyByIndex 1, "haha"
    MsgBox c.toString(, True)
    End
    
'    Dim cc As New Class1, c2 As Class1
'    cc.name = "test"
'
'    c.Add "my string", "mykey"
'    Debug.Print "mykey=" & c.KeyExists("mykey")
'
'
'    c.Add cc, "test"
'
'    If c.KeyExists("test") Then
'        Debug.Print "found object item by key"
'        Set c2 = c("test")
'        Debug.Print c2.name
'    Else
'        Debug.Print "failed to find object item by key"
'    End If
'
'    Dim v
'    For Each v In c
'        Debug.Print TypeName(v)
'    Next

    For i = 0 To 5
        c.Add "test " & i, i
    Next
    
    If c.toFile("C:\test.bin") Then
        Dim cc As New CollectionEx
        MsgBox "Added :" & cc.fromFile("C:\test.bin")
        MsgBox "Appended :" & cc.fromFile("C:\test.bin", True)
        MsgBox "Keys: " & vbCrLf & Join(cc.ext.Keys(), vbCrLf)
        MsgBox "Values to string: " & vbCrLf & cc.toString()
        MsgBox "Item key=4 value=" & cc(4, 1)
    End If
    
    
    
       
      
'
'    c.ext.ChangeIndex(1) = 5
'

'    Dim tmp(5) As Class1
'    For i = 0 To 5
'        Set tmp(i) = New Class1
'        tmp(i).name = "test " & i
'    Next
    
    'c.ext.FromArray Array(1, 2, 3, 4, 5, 6)
    'c.ext.FromArray Array(7, 8, 9, 10, 11, 12), False
 
'    Dim j()
'    c.ext.FromArray tmp
'    j = c.ext.ToArray()
    
 End
    

End Sub
