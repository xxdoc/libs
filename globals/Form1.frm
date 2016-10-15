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
    
     
    
    Dim tmp() As String
    push tmp, 1
    
    globals.hash.HashString ("test")
    
    
End Sub




'Sub push(ary, value) 'this modifies parent ary object
'    On Error GoTo init
'    X = UBound(ary) '<-throws Error If Not initalized
'    ReDim Preserve ary(UBound(ary) + 1)
'    ary(UBound(ary)) = value
'    Exit Sub
'init:     ReDim ary(0): ary(0) = value
'End Sub
