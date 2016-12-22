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

    Dim s As New StringEx
    Const x As Boolean = False
    
    d = esc(" \r\n\t This is my string  \t\r\n ")
    
    s = d
    
    Debug.Print unesc("'" & s & "'")
    Debug.Print unesc("'" & s.rTrim(x) & "'")
    Debug.Print unesc("'" & s.lTrim(x) & "'")
    
        

    
End Sub

Function unesc(x)
    unesc = Replace(x, vbTab, "\t")
    unesc = Replace(unesc, vbCr, "\r")
    unesc = Replace(unesc, vbLf, "\n")
End Function

Function esc(x)
    esc = Replace(x, "\t", vbTab)
    esc = Replace(esc, "\r", vbCr)
    esc = Replace(esc, "\n", vbLf)
End Function
