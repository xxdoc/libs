Attribute VB_Name = "modDzTest"
Option Explicit

Public hLib As Long
Public hLib2 As Long
Private fHand As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'one instance for simplicity, if using multi api I would pass in
'objptr and collection lookup over stolen references like in origial demos
Private resp As CCurlResponse
Public dllPath As String
Public caBundleFound As Boolean

Function initLib() As Boolean
    
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
        Exit Function
    End If
    
    Dim base() As String, b
    Const dll = "libcurl.dll"
    
    push base, App.path
    push base, App.path & "\bin"
    push base, GetParentFolder(App.path)
    push base, GetParentFolder(App.path) & "\bin"
    push base, GetParentFolder(App.path, 2)
    push base, GetParentFolder(App.path, 2) & "\bin"
    
    For Each b In base
        hLib = LoadLibrary(b & "\" & dll)
        If hLib <> 0 Then
            dllPath = b
            Form1.List1.AddItem "Loaded " & b & "\" & dll
            hLib2 = LoadLibrary(b & "\" & "vb" & dll)
            If hLib2 = 0 Then
                Form1.List1.AddItem "Failed to load vbLibcurl.dll from same directory?!"
            Else
                Form1.List1.AddItem "Loaded vbLibCurl.dll from same directory."
                If FileExists(dllPath & "\curl-ca-bundle.crt") Then
                    caBundleFound = True
                    Form1.List1.AddItem "Found curl-ca-bundle.crt"
                Else
                    Form1.List1.AddItem "curl-ca-bundle.crt not found ssl will not work..."
                End If
            End If
            Exit For
        End If
    Next
        
    If hLib <> 0 And hLib2 <> 0 Then
        initLib = True
    Else
        Form1.List1.AddItem "Could not load libcurl.dll"
    End If

End Function

'if toFile is empty then we download to CMemBuffer only and return that
'else we will return the written file size as long
'if we fail to open the output file this will raise an error
Function Download(Url As String, Optional toFile As String, Optional INotify As ICurlProgress, Optional connectTimeout As Long = 15, Optional totalDLTimeout As Long = 0) As CCurlResponse
    
    Dim easy As Long, v As Variant
    Dim ret As CURLcode 'enum
 
    Set resp = New CCurlResponse
    Set resp.INotify = INotify
    Set Download = resp
    
    resp.Url = Url
    
    If Len(toFile) > 0 Then
        If FileExists(toFile) Then DeleteFile toFile
        resp.localPath = toFile
        fHand = FreeFile
        Open toFile For Binary As fHand 'this can throw an error
    Else
        Set resp.memFile = New CMemBuffer
    End If
    
    easy = vbcurl_easy_init()
    resp.hCurl = easy
    vbcurl_easy_setopt easy, CURLOPT_URL, Url
    'vbcurl_easy_setopt easy, CURLOPT_WRITEDATA, ObjPtr(buf) 'either 0 or live objptr
    vbcurl_easy_setopt easy, CURLOPT_WRITEFUNCTION, AddressOf WriteFunction
    vbcurl_easy_setopt easy, CURLOPT_DEBUGFUNCTION, AddressOf DebugFunction
    vbcurl_easy_setopt easy, CURLOPT_VERBOSE, True
    If totalDLTimeout > 0 Then vbcurl_easy_setopt easy, CURLOPT_TIMEOUT, totalDLTimeout
    If connectTimeout > 0 Then vbcurl_easy_setopt easy, CURLOPT_CONNECTTIMEOUT, connectTimeout
    If caBundleFound Then vbcurl_easy_setopt easy, CURLOPT_CAINFO, dllPath & "\curl-ca-bundle.crt"

    resp.CurlReturnCode = vbcurl_easy_perform(easy)
    resp.queryHeaders
   
    vbcurl_easy_cleanup easy
    resp.hCurl = 0
    
    If Not resp.isMemFile Then
        'Download = LOF(fHand)
        Close fHand
        fHand = 0
    End If
    
    If Not INotify Is Nothing Then INotify.Complete resp
    Set resp = Nothing
    
End Function


Private Function DebugFunction(ByVal info As curl_infotype, ByVal rawBytes As Long, ByVal numBytes As Long, ByVal extra As Long) As Long

    Dim tmp As String, i As Long
    Dim b() As Byte
    
    If info >= 3 Then Exit Function
    
'    If info = CURLINFO_DATA_IN Then Exit Function '3 every data packet we dont care...maybe disable debug func now?
'    If info = CURLINFO_DATA_OUT Then Exit Function
'    If info = CURLINFO_SSL_DATA_IN Then Exit Function
'    If info = CURLINFO_SSL_DATA_OUT Then Exit Function
    
    ReDim b(numBytes - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, numBytes
    tmp = StrConv(b, vbUnicode, &H409)
    
    If info = CURLINFO_HEADER_IN Then
        resp.addHeader tmp
    Else
        resp.addInfo info, tmp
    End If
    
    DebugFunction = 0
    
End Function


Private Function WriteFunction(ByVal rawBytes As Long, ByVal sz As Long, ByVal nmemb As Long, ByVal extra As Long) As Long
    
    On Error Resume Next
    Dim totalBytes As Long, i As Long, b() As Byte, ret As CURLcode, v As Variant
    
    totalBytes = sz * nmemb
    
    If resp.isMemFile Then
        resp.memFile.memAppendBuf rawBytes, totalBytes
    Else
        ReDim b(totalBytes - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, totalBytes
        Put fHand, , b()
    End If
    
    DoEvents
    WriteFunction = resp.notifyParent(totalBytes)  'if this returns 0 dl is aborted...
    
End Function

Function info2Text(i As curl_infotype) As String
    
    Dim s As String
    
    If i = CURLINFO_TEXT Then s = "TEXT"                '0
    If i = CURLINFO_HEADER_IN Then s = "HEADER_IN"      '1
    If i = CURLINFO_HEADER_OUT Then s = "HEADER_OUT"    '2
    If i = CURLINFO_DATA_IN Then s = "DATA_IN"          '3
    If i = CURLINFO_DATA_OUT Then s = "DATA_OUT"        '4
    If i = CURLINFO_SSL_DATA_IN Then s = "SSL_IN"       '5
    If i = CURLINFO_SSL_DATA_OUT Then s = "SSL_OUT"     '6
    If i = CURLINFO_END Then s = "END"                  '7
    If Len(s) = 0 Then s = "Unknown " & i
    
    info2Text = s
    
End Function

Function curlCode2Text(x As CURLcode) As String
    
    Dim s As String
    s = "Unknown " & x
    
    If x = 0 Then s = "CURLE_OK"
    If x = 42 Then s = "CURLE_ABORTED_BY_CALLBACK"
    If x = 44 Then s = "CURLE_BAD_CALLING_ORDER"
    If x = 61 Then s = "CURLE_BAD_CONTENT_ENCODING"
    If x = 36 Then s = "CURLE_BAD_DOWNLOAD_RESUME"
    If x = 43 Then s = "CURLE_BAD_FUNCTION_ARGUMENT"
    If x = 46 Then s = "CURLE_BAD_PASSWORD_ENTERED"
    If x = 7 Then s = "CURLE_COULDNT_CONNECT"
    If x = 6 Then s = "CURLE_COULDNT_RESOLVE_HOST"
    If x = 5 Then s = "CURLE_COULDNT_RESOLVE_PROXY"
    If x = 2 Then s = "CURLE_FAILED_INIT"
    If x = 63 Then s = "CURLE_FILESIZE_EXCEEDED"
    If x = 37 Then s = "CURLE_FILE_COULDNT_READ_FILE"
    If x = 9 Then s = "CURLE_FTP_ACCESS_DENIED"
    If x = 15 Then s = "CURLE_FTP_CANT_GET_HOST"
    If x = 16 Then s = "CURLE_FTP_CANT_RECONNECT"
    If x = 32 Then s = "CURLE_FTP_COULDNT_GET_SIZE"
    If x = 19 Then s = "CURLE_FTP_COULDNT_RETR_FILE"
    If x = 29 Then s = "CURLE_FTP_COULDNT_SET_ASCII"
    If x = 17 Then s = "CURLE_FTP_COULDNT_SET_BINARY"
    If x = 25 Then s = "CURLE_FTP_COULDNT_STOR_FILE"
    If x = 31 Then s = "CURLE_FTP_COULDNT_USE_REST"
    If x = 30 Then s = "CURLE_FTP_PORT_FAILED"
    If x = 21 Then s = "CURLE_FTP_QUOTE_ERROR"
    If x = 64 Then s = "CURLE_FTP_SSL_FAILED"
    If x = 10 Then s = "CURLE_FTP_USER_PASSWORD_INCORRECT"
    If x = 14 Then s = "CURLE_FTP_WEIRD_227_FORMAT"
    If x = 11 Then s = "CURLE_FTP_WEIRD_PASS_REPLY"
    If x = 13 Then s = "CURLE_FTP_WEIRD_PASV_REPLY"
    If x = 8 Then s = "CURLE_FTP_WEIRD_SERVER_REPLY"
    If x = 12 Then s = "CURLE_FTP_WEIRD_USER_REPLY"
    If x = 20 Then s = "CURLE_FTP_WRITE_ERROR"
    If x = 41 Then s = "CURLE_FUNCTION_NOT_FOUND"
    If x = 52 Then s = "CURLE_GOT_NOTHING"
    If x = 34 Then s = "CURLE_HTTP_POST_ERROR"
    If x = 33 Then s = "CURLE_HTTP_RANGE_ERROR"
    If x = 22 Then s = "CURLE_HTTP_RETURNED_ERROR"
    If x = 45 Then s = "CURLE_INTERFACE_FAILED"
    If x = 67 Then s = "CURLE_LAST"
    If x = 38 Then s = "CURLE_LDAP_CANNOT_BIND"
    If x = 62 Then s = "CURLE_LDAP_INVALID_URL"
    If x = 39 Then s = "CURLE_LDAP_SEARCH_FAILED"
    If x = 40 Then s = "CURLE_LIBRARY_NOT_FOUND"
    If x = 24 Then s = "CURLE_MALFORMAT_USER"
    If x = 50 Then s = "CURLE_OBSOLETE"
    If x = 28 Then s = "CURLE_OPERATION_TIMEOUTED"
    If x = 27 Then s = "CURLE_OUT_OF_MEMORY"
    If x = 18 Then s = "CURLE_PARTIAL_FILE"
    If x = 26 Then s = "CURLE_READ_ERROR"
    If x = 56 Then s = "CURLE_RECV_ERROR"
    If x = 55 Then s = "CURLE_SEND_ERROR"
    If x = 65 Then s = "CURL_SEND_FAIL_REWIND"
    If x = 57 Then s = "CURLE_SHARE_IN_USE"
    If x = 60 Then s = "CURLE_SSL_CACERT"
    If x = 58 Then s = "CURLE_SSL_CERTPROBLEM"
    If x = 59 Then s = "CURLE_SSL_CIPHER"
    If x = 35 Then s = "CURLE_SSL_CONNECT_ERROR"
    If x = 66 Then s = "CURLE_SSL_ENGINE_INITFAILED"
    If x = 53 Then s = "CURLE_SSL_ENGINE_NOTFOUND"
    If x = 54 Then s = "CURLE_SSL_ENGINE_SETFAILED"
    If x = 51 Then s = "CURLE_SSL_PEER_CERTIFICATE"
    If x = 49 Then s = "CURLE_TELNET_OPTION_SYNTAX"
    If x = 47 Then s = "CURLE_TOO_MANY_REDIRECTS"
    If x = 48 Then s = "CURLE_UNKNOWN_TELNET_OPTION"
    If x = 1 Then s = "CURLE_UNSUPPORTED_PROTOCOL"
    If x = 3 Then s = "CURLE_URL_MALFORMAT"
    If x = 4 Then s = "CURLE_URL_MALFORMAT_USER"
    If x = 23 Then s = "CURLE_WRITE_ERROR"

    curlCode2Text = s
    
End Function

'multi way...

' This function illustrates a couple of key concepts in libcurl.vb.
' First, the data passed in rawBytes is an actual memory address
' from libcurl. Hence, the data is read using the MemByte() function
' found in the VBVM6Lib.tlb type library. Second, the extra parameter
' is passed as a raw long (via ObjPtr(buf)) in Sub EasyGet()), and
' we use the AsObject() function in VBVM6Lib.tlb to get back at it.

 'If extra <> 0 Then
        'Set obj = AsObject(extra)
        'Set buf = obj
        'buf.
        'ObjectPtr(obj) = 0&
