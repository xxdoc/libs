VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hList As Long

Private Sub Class_Initialize()
    hList = vbcurl_slist_create()
End Sub

Private Sub Class_Terminate()
    vbcurl_slist_free hList
End Sub

Function Append(aryStrOrCol) As Boolean
    On Error GoTo hell
    Dim s As String, x
    If IsArray(aryStrOrCol) Or TypeName(aryStrOrCol) = "Collection" Then
        For Each x In aryStrOrCol
            s = CStr(x)
            If Len(s) > 0 Then vbcurl_slist_append hList, strPtr(s)
        Next
    Else
        s = CStr(aryStrOrCol)
        If Len(s) > 0 Then vbcurl_slist_append hList, strPtr(s)
    End If
    Append = True
Exit Function
hell:
End Function

Function Attach(hCurl As Long) As CURLcode
    Attach = vbcurl_easy_setopt(hCurl, CURLOPT_HTTPHEADER, hList)
End Function
