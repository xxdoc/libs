VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' weCollection Class
'
' Test class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 04/18/2000
'
'*********************************************************************************************
Option Explicit

Event TestEvent(ByVal Param1 As Long, ByRef FormatedParam As String)

Sub RaiseTheEvent()
Dim ByRefParam As String
   
   RaiseEvent TestEvent(Int(Rnd * 10000), ByRefParam)
   
   Debug.Print "Returned string is = " & ByRefParam
   
End Sub


