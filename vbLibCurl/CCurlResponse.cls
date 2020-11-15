VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CurlReturnCode As CURLcode
Public ContentType
Public ResponseCode 'As Long
'Public DownloadSize 'As Long
Public DownloadLength 'As Long
Public TotalTime 'As Long
Public BytesReceived As Double 'huge downloads could overflow
Public abort As Boolean
Public hCurl As Long
Public memFile As CMemBuffer
Public headers As New Collection
Public DebugMsgs As New Collection
Public localPath As String
Public Url As String
Public INotify As ICurlProgress
Private gotSizes As Boolean

Property Get isMemFile() As Boolean
    isMemFile = Not memFile Is Nothing
End Property

Friend Function addHeader(msg As String)
    headers.Add msg
    If Not INotify Is Nothing Then INotify.Header Me, msg
End Function

Friend Function addInfo(info As curl_infotype, msg As String)
    DebugMsgs.Add msg
    If Not INotify Is Nothing Then INotify.InfoMsg Me, info, msg
End Function

Friend Function notifyParent(size As Long) As Long
    
    On Error Resume Next
    Dim ret As CURLcode, v As Variant
    
    BytesReceived = BytesReceived + size
    
    If Not gotSizes Then
        ret = vbcurl_easy_getinfo(hCurl, CURLINFO_CONTENT_LENGTH_DOWNLOAD, v)
        If ret = CURLE_OK Then DownloadLength = v
        
        'ret = vbcurl_easy_getinfo(hCurl, CURLINFO_SIZE_DOWNLOAD, v) 'this must be bytes downloaded so far..
        'If ret = CURLE_OK Then DownloadSize = v
        
        If Not INotify Is Nothing Then INotify.Init Me
        gotSizes = True
    End If
    
    If Not INotify Is Nothing Then INotify.Progress Me 'consumer can set abort if desired...
    
    If abort Then
        notifyParent = 0
    Else
        notifyParent = size
    End If

End Function

Function dump() As String
    Dim tmp() As String
    
    push tmp, "CurlCode = " & curlCode2Text(Me.CurlReturnCode)
    push tmp, "ContentType = " & ContentType
    push tmp, "ResponseCode = " & ResponseCode
    'push tmp, "DownloadSize  = " & DownloadSize
    push tmp, "DownloadLength  = " & DownloadLength
    push tmp, "TotalTime = " & TotalTime
    
    If isMemFile Then
        push tmp, "isMemFile = True"
    Else
        push tmp, "localPath = " & localPath
    End If
    
    push tmp, "url = " & Url
    
    push tmp, "headers:"
    push tmp, Join(c2a(headers), "")

    dump = Join(tmp, vbCrLf)
    
End Function

Friend Sub queryHeaders()
    Dim v As Variant, ret As CURLcode
    
    ret = vbcurl_easy_getinfo(hCurl, CURLINFO_CONTENT_TYPE, v)
    If ret = CURLE_OK Then ContentType = v
    
    ret = vbcurl_easy_getinfo(hCurl, CURLINFO_RESPONSE_CODE, v)
    If ret = CURLE_OK Then ResponseCode = v
    
    ret = vbcurl_easy_getinfo(hCurl, CURLINFO_TOTAL_TIME, v)
    If ret = CURLE_OK Then TotalTime = v

End Sub
