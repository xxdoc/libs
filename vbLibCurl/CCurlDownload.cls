VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'todo: support POST and form fields

'so we broke this out into its own class because libcurl supports soooo many options
'once you get to wanting to configure it more you are beyond the number of args you want to
'pass to a single download function. You could just use the setopt api and enum vals but the names
'arent always friendly and are all longer by default

Public hCurl As Long
Public errList As ListBox  'for when we move this to a dll, could be a collection but thats not immediate feedback, should be a function on iNotify i suppose

Event Init(obj As CCurlResponse)
Event Progress(obj As CCurlResponse)
Event Complete(obj As CCurlResponse)
Event Header(obj As CCurlResponse, ByVal msg As String)
Event InfoMsg(obj As CCurlResponse, ByVal info As curl_infotype, ByVal msg As String)

Enum curlEvents
    ceInit
    ceProgress
    ceComplete
    ceHeader
    ceInfo
End Enum

Friend Sub TriggerEvent(e As curlEvents, obj As CCurlResponse, Optional arg1, Optional arg2)
    If e = ceInit Then RaiseEvent Init(obj)
    If e = ceProgress Then RaiseEvent Progress(obj)
    If e = ceComplete Then RaiseEvent Complete(obj)
    If e = ceHeader Then RaiseEvent Header(obj, arg1)
    If e = ceInfo Then RaiseEvent InfoMsg(obj, arg1, arg2)
End Sub


Private Sub Class_Initialize()
    hCurl = vbcurl_easy_init()
End Sub

Private Sub Class_Terminate()
    vbcurl_easy_cleanup hCurl
End Sub

'these ones are typically set once per class use so give one liner shortcut
Sub Configure(Useragent As String, _
             Optional followRedirects As Boolean = True, _
             Optional maxTimeout As Long, _
             Optional connectTimeout As Long, _
             Optional erList As ListBox = Nothing _
)

    Me.Useragent = Useragent
    Me.followRedirects = followRedirects
    Me.maxTime = maxTimeout
    Me.connectTimeout = connectTimeout
    'If Not notifyUI Is Nothing Then Set INotify = notifyUI
    If Not erList Is Nothing Then Set errList = erList
    
End Sub

Property Let Useragent(ua As String)
    vbcurl_easy_setopt hCurl, CURLOPT_USERAGENT, ua
End Property

Property Let connectTimeout(secs As Long)
    vbcurl_easy_setopt hCurl, CURLOPT_CONNECTTIMEOUT, secs
End Property

Property Let followRedirects(v As Boolean)
     vbcurl_easy_setopt hCurl, CURLOPT_FOLLOWLOCATION, IIf(v, 1, 0)
End Property

Property Let maxTime(secs As Long)
    vbcurl_easy_setopt hCurl, CURLOPT_TIMEOUT, secs
End Property

Function escape(ByVal buf As String) As String
    Dim tmp As Long, s As String
    tmp = vbcurl_string_escape(strPtr(buf), Len(buf))
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    escape = s
End Function

Function unescape(ByVal buf As String) As String
    Dim tmp  As Long, s As String
    tmp = vbcurl_string_unescape(strPtr(buf), Len(buf))
    CopyMemory ByVal VarPtr(s), tmp, 4 'steal a ref to an existing BSTR so we now own it
    unescape = s
End Function

'these are usually set per request
Property Let Referrer(v As String)
     vbcurl_easy_setopt hCurl, CURLOPT_REFERER, v
End Property

Property Let Cookie(c As String)
     vbcurl_easy_setopt hCurl, CURLOPT_COOKIE, c
End Property


Function Download(url As String, Optional toFile As String) As CCurlResponse
 
    Set Download = New CCurlResponse 'always a valid class to return
    'If Not INotify Is Nothing Then Set Me.INotify = INotify
     
    If Not initLib(errList) Then Exit Function
    
    If Not modDzTest.ActiveResponse Is Nothing Then
        If Not errList Is Nothing Then errList.AddItem "Limit Only one download active at a time right now you can fix this.."
        Exit Function
    End If
    
    Set modDzTest.ActiveResponse = Download 'one active instance
    
    With ActiveResponse
    
        Set .owner = Me
        If Not .Initilize(url, toFile) Then Exit Function
        
        vbcurl_easy_setopt hCurl, CURLOPT_URL, url
        vbcurl_easy_setopt hCurl, CURLOPT_WRITEDATA, ObjPtr(ActiveResponse)
        vbcurl_easy_setopt hCurl, CURLOPT_WRITEFUNCTION, AddressOf WriteFunction
        vbcurl_easy_setopt hCurl, CURLOPT_DEBUGFUNCTION, AddressOf DebugFunction
        vbcurl_easy_setopt hCurl, CURLOPT_VERBOSE, True
        If FileExists(caBundle) Then vbcurl_easy_setopt hCurl, CURLOPT_CAINFO, caBundle
    
        .CurlReturnCode = vbcurl_easy_perform(hCurl)
        .Finalize
        
        'vbcurl_easy_reset hCurl  'wipe out entire config for next time
        
    End With
    
    Set ActiveResponse = Nothing
    
End Function


