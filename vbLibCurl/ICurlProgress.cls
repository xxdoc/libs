VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICurlProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'interface for Implements

Sub Init(obj As CCurlResponse)

End Sub
 
Sub Progress(obj As CCurlResponse)
    
End Sub

Sub Complete(obj As CCurlResponse)

End Sub

Sub Header(obj As CCurlResponse, msg As String)

End Sub

Sub InfoMsg(obj As CCurlResponse, info As curl_infotype, msg As String)

End Sub
