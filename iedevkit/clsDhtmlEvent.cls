VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDhtmlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private myParent As Object
Private myFunction As String
Private myCallingArgument

Sub SetReference(parent As Object, sFunction As String, Optional CallingArgument As String)
    Set myParent = parent
    myFunction = sFunction
    myCallingArgument = CallingArgument
End Sub

Function DefaultFunction()
Attribute DefaultFunction.VB_UserMemId = 0
    CallByName myParent, myFunction, VbMethod, myCallingArgument
End Function

 
