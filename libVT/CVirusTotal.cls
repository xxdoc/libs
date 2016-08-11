VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVirusTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'8.5.16 - includes first seen date in default report if using private api key

Dim http As XMLHTTP

Const URL_API_BASIS = "http://www.virustotal.com/vtapi/v2/"
Const URL_SCAN_FILE = "file/scan"
Const URL_FILE_REPORT = "file/report"
Const URL_SCAN_URL = "url/scan"
Const URL_URL_REPORT = "url/report"
Const URL_MAKE_COMMENT = "comments/put"
Public API_KEY As String

Const PUBLIC_API_KEY = "a949ea9c64e7145a065b0e562673a66216a132712e958168c8c00ee5f451485b"
Public usingPrivateKey As Boolean

Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long

Public DebugLog As New Collection 'for logging...
Public owner As Object 'form with public abort as boolean variable that we query

Private WithEvents Timer1 As Timer 'we cant make this public it will switch to object type in typelib
Attribute Timer1.VB_VarHelpID = -1

Private NeedsDelay As Boolean
Private LastQueryTime As Long
Private m_delayInterval As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public report_cache_dir As String 'if this is set it will cache the results here in case looked up again.
Dim repeatOnce As Boolean

Const LANG_US = &H409

Property Let TimerObj(tmr As Object)
    On Error Resume Next
    Set Timer1 = tmr
End Property

Public Function RawDownloadFile(URL, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(URL), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then RawDownloadFile = True
End Function

Property Get delayInterval() As Long
    delayInterval = m_delayInterval
End Property

Property Let delayInterval(x As Long)
    If Not usingPrivateKey Then
        m_delayInterval = x
    End If
End Property


Public Function DownloadFile(hash As String, toDir As String) As Boolean

    If Len(API_KEY) = 0 Then Class_Initialize
    If Not usingPrivateKey Then Exit Function
    If Not FolderExists(toDir) Then Exit Function
    
'private api: GET /vtapi/v2/file/download
'  apikey
'  hash    The md5/sha1/sha256 hash of the file you want to download.

    Dim URL As String
    URL = "https://www.virustotal.com/vtapi/v2/file/download?apikey=" & API_KEY & "&hash=" & hash
    DownloadFile = RawDownloadFile(URL, toDir & "\" & hash)

End Function

'this class supports a couple nice things..
'1) if you set the cache dir, it will save raw json there as optimization in case you lookup the same file
'    file again. this will sve you the 4sec delay per request.
'2) this class tracks when the last request was made, and implements its own optimized delay so that the
'   4 second threshold is not exceeded.
'3) since vb doesnt support json, a CScan object is always passed back with details of the scan. parsing
'    was just done in VB and is VT json specific and maybe a bit shady as it was done quickly..

Private Function testAbort() As Boolean
    On Error Resume Next
    If Not (owner Is Nothing) Then testAbort = owner.abort
    If Err.Number <> 0 Then testAbort = False
End Function

Function GetReport(hash As String) As CScan

    On Error Resume Next
    
    Dim scan As New CScan
   
    Dim my_json As String
    Dim sStatus As String
    Dim status As Long
    Dim curtime As Long
    Dim success As Boolean
    Dim errorMsg As String
    
    If Len(API_KEY) = 0 Then Class_Initialize
    
    Set DebugLog = New Collection
    
    If Timer1 Is Nothing Then
        MsgBox "Must set Timer1 reference before using GetReport!"
        Exit Function
    End If
    
    'Set Timer1 = tmr
    Timer1.Enabled = False
    Timer1.Interval = m_delayInterval

    Set http = New XMLHTTP
    
    If http Is Nothing Then
        errorMsg = "Could not create XMLHTTP Object"
        GoTo errorExit
    End If
        
    If FolderExists(report_cache_dir) Then
        If FileExists(report_cache_dir & "\" & hash & ".txt") Then
            DebugLog.Add "Cached report found for " & hash
            my_json = ReadFile(report_cache_dir & "\" & hash & ".txt")
            GoTo parse_json
        End If
    End If
            
    DelayIfNeeded
    
    DebugLog.Add "Connecting to VirusTotal to query report for " & hash
    
    DoEvents
    
    NeedsDelay = True 'need 4 second delay for next if using public api key
    
    LastQueryTime = GetTickCount
    success = Get_Web_Report(hash, my_json, sStatus, status)
    scan.HTTP_ResponseCode = status
    scan.HTTP_StatusMessage = sStatus
    
    If Not success Then
        errorMsg = "Web request failure - status code: " & status & " " & sStatus
        GoTo errorExit
    End If
   
    DebugLog.Add "Query successful for hash: " & hash
         
    If FolderExists(report_cache_dir) Then
        writeFile report_cache_dir & "\" & hash & ".txt", my_json
    End If
    
parse_json:
    scan.ParseResponse my_json
    Set GetReport = scan
    
    DebugLog.Add "Detections: " & scan.positives & " Response Code: " & scan.response_code
    If scan.HadError Then DebugLog.Add "Had JSON Parsing error..."
    
Exit Function

errorExit:

    DebugLog.Add errorMsg
    scan.Errors.Add errorMsg
    scan.HadError = True
    Set GetReport = scan
    
End Function



Private Function Get_Web_Report(hash, out_response As String, out_status As String, out_statusCode As Long) As Boolean
    
    Err.Clear
        
    Dim x As Variant
    out_status = Empty
    out_response = Empty
    repeatOnce = True
    
tryAgain:

On Error GoTo hell

     Set http = Nothing
     Set http = New XMLHTTP
     
1    http.Open "POST", URL_API_BASIS & URL_FILE_REPORT, False
2    http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
     If Not doSend("key=" & API_KEY & "&resource=" & hash & IIf(usingPrivateKey, "&allinfo=1", "")) Then
        DebugLog.Add "Error in Get_Web_Report http.send failed desc: " & Err.Description & " ReadyState: " & http.readyState & " Status: " & http.status
        If repeatOnce Then
            repeatOnce = False
            DebugLog.Add "Failure Must delay for " & m_delayInterval & " milliseconds"
            Timer1.Interval = m_delayInterval
            Timer1.Enabled = True
            While Timer1.Enabled
                DoEvents
                Sleep 20
                If testAbort Then Exit Function
            Wend
            GoTo tryAgain
        End If
     End If
        
     DoEvents
     
5    out_status = http.statusText
6    out_statusCode = http.status
7    out_response = http.responseText
     If out_status = "OK" Then Get_Web_Report = True
     repeatOnce = True
     
hell:
    DoEvents
    If Err.Number <> 0 Then
        DebugLog.Add "Error in Get_Web_Report Line: " & Erl & " desc: " & Err.Description & " ReadyState: " & http.readyState & " Status: " & http.status
    End If
    
End Function

'status code of 1223 and drop all response headers if the server returns a HTTP/204 No Content
Private Function doSend(s As String) As Boolean
    On Error Resume Next
    http.send s
    doSend = IIf(Err.Number = 0, True, False)
End Function


Sub Closeit()
    On Error Resume Next
    http.abort
End Sub

Function AddComment(hash, comment, out_response As String, out_status As String, out_statusCode As Long) As Boolean
    
    If Len(API_KEY) = 0 Then Class_Initialize
    
    Err.Clear
    On Error GoTo hell
    
    Dim x As Variant
    out_status = Empty
    out_response = Empty
    
1    http.Open "POST", URL_API_BASIS & URL_MAKE_COMMENT, False
2    http.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
3    http.send "key=" & API_KEY & "&resource=" & hash & "&comment=" & comment
    
     DoEvents
     
5    out_status = http.statusText
6    out_statusCode = http.status
7    out_response = http.responseText
     If out_status = "OK" Then AddComment = True
    
hell:
    DoEvents
    If Err.Number <> 0 Then
        DebugLog.Add "Error in AddComment Line: " & Erl & " desc: " & Err.Description
    End If
    
End Function

Private Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Private Function FileExists(p) As Boolean
    If Len(p) = 0 Then Exit Function
    If Dir(p, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  f = FreeFile
  Dim b() As Byte
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Private Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = True
    Exit Function
hell: writeFile = False
End Function

Function SetPrivateApiKey(ByVal key As String)

    key = Trim(key)
    
    If key <> API_KEY Then
        SaveSetting "VirusTotal", "config", "private_api_key", key
    End If
    
    If Len(key) > 0 Then
        API_KEY = key
        m_delayInterval = 0
        usingPrivateKey = True
    Else
        API_KEY = PUBLIC_API_KEY
        m_delayInterval = 17300 '4 requests per minute max
        usingPrivateKey = False
    End If
    
End Function

Function ReadPrivateApiKey() As String
    ReadPrivateApiKey = GetSetting("VirusTotal", "config", "private_api_key", "")
End Function

Private Sub Class_Initialize()
    NeedsDelay = False
    SetPrivateApiKey ReadPrivateApiKey()
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub

Private Function DelayIfNeeded()

    If usingPrivateKey Then Exit Function
    
    curtime = GetTickCount
    If LastQueryTime = 0 Then
        LastQueryTime = curtime
        NeedsDelay = False
    Else
        If NeedsDelay = True Then
            Interval = curtime - LastQueryTime
            If Interval < m_delayInterval Then
                Interval = m_delayInterval - Interval
                DebugLog.Add "Must delay for " & Interval & " milliseconds"
                Timer1.Interval = Interval
                Timer1.Enabled = True
                While Timer1.Enabled
                    DoEvents
                    Sleep 20
                    If testAbort Then Timer1.Enabled = False
                Wend
                NeedsDelay = False
                LastQueryTime = curtime
            End If
        Else
            NeedsDelay = False
            LastQueryTime = curtime
        End If
    End If
    
End Function

Function SubmitFile(fpath As String) As CScan

    'base example from: http://wqweto.wordpress.com/2011/07/12/vb6-using-wininet-to-post-binary-file/
    Const STR_BOUNDARY  As String = "3fbd04f5-b1ed-4060-99b9-fca7ff59c113"
    Dim sPostData       As String
    Dim fName As String
    Dim fData As String
    
    On Error Resume Next
    
    If Len(API_KEY) = 0 Then Class_Initialize
    
    Dim scan As New CScan
    scan.LocalFilePath = fpath
    
    Dim my_json As String
    Dim sStatus As String
    Dim status As Long
    Dim curtime As Long
    Dim success As Boolean
    Dim errorMsg As String
    
    Set DebugLog = New Collection
    
    If Timer1 Is Nothing Then
        MsgBox "Must set Timer1 reference before using SubmitFile!"
        Exit Function
    End If

    Timer1.Enabled = False
    Timer1.Interval = m_delayInterval
    
    If Not FileExists(fpath) Then
        errorMsg = "Upload file does not exist: " & fpath
        GoTo errorExit
    End If
    
    Set http = New XMLHTTP
    
    If http Is Nothing Then
        errorMsg = "Could not create XMLHTTP Object"
        GoTo errorExit
    End If
    
    fName = fpath
    If InStr(fName, "\") > 0 Then fName = Mid$(fpath, InStrRev(fpath, "\") + 1)
    
    fData = ReadFile(fpath)
    
    sPostData = "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & fName & """" & vbCrLf & _
        "Content-Type: application/octet-stream" & vbCrLf & vbCrLf & _
        fData & vbCrLf & _
        "--" & STR_BOUNDARY & vbCrLf & _
        "Content-Disposition: form-data; name=""key"" " & vbCrLf & vbCrLf & _
        API_KEY & vbCrLf & _
        "--" & STR_BOUNDARY & "--"
    
    DelayIfNeeded
    
    DebugLog.Add "Connecting to VirusTotal to submit file: " & fName
    NeedsDelay = True 'need 4 second delay for next
    DoEvents
    LastQueryTime = GetTickCount
    
    http.Open "POST", URL_API_BASIS & URL_SCAN_FILE, False
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & STR_BOUNDARY
    http.send pvToByteArray(sPostData) 'blocking call until it returns..
    
    DoEvents
     
    scan.HTTP_ResponseCode = http.status
    scan.HTTP_StatusMessage = http.statusText
    scan.ParseResponse http.responseText
    Set SubmitFile = scan
    
    'Response Codes:
    ' -4: file not found.
    ' -3: public API request rate exceeded.
    ' -2: resource is currently being analyzed.
    ' -1: you do not have the required privileges (wrong API key?).
     
    DebugLog.Add "Result: " & scan.verbose_msg & " Response Code: " & scan.response_code
    
    Exit Function

errorExit:

    DebugLog.Add errorMsg
    scan.Errors.Add errorMsg
    scan.HadError = True
    Set SubmitFile = scan
    
    
End Function


Private Function pvToByteArray(sText As String) As Byte()
    pvToByteArray = StrConv(sText, vbFromUnicode, LANG_US)
End Function


