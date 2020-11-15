Attribute VB_Name = "modDzTest"
Option Explicit

Public hLib As Long
Public hLib2 As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'one instance for simplicity, if using multi api I would pass in
'objptr and collection lookup over stolen references like in origial demos
Private resp As CCurlResponse
Private caBundle As String

'we probably need a CCurlRequest object to configure and add headers/cookies etc..but not till i need it...
Public Referrer As String
Public UserAgent As String

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
            Form1.List1.AddItem "Loaded " & b & "\" & dll
            hLib2 = LoadLibrary(b & "\vb" & dll)
            If hLib2 = 0 Then
                Form1.List1.AddItem "Failed to load vbLibcurl.dll from same directory?!"
            Else
                Form1.List1.AddItem "Loaded vbLibCurl.dll from same directory."
                caBundle = b & "\curl-ca-bundle.crt"
                If FileExists(caBundle) Then
                    Form1.List1.AddItem "Found curl-ca-bundle.crt"
                Else
                    caBundle = Empty
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


Function Download( _
    url As String, _
    Optional toFile As String, _
    Optional INotify As ICurlProgress, _
    Optional connectTimeout As Long = 15, _
    Optional totalDLTimeout As Long = 0, _
    Optional followRedirects As Boolean = True, _
    Optional cookie As String = Empty _
) As CCurlResponse
 
    Set resp = New CCurlResponse
    Set Download = resp
    
    With resp
    
        If Not .Initilize(url, toFile, INotify) Then Exit Function
        
        vbcurl_easy_setopt .hCurl, CURLOPT_URL, url
        vbcurl_easy_setopt .hCurl, CURLOPT_WRITEDATA, ObjPtr(resp)
        vbcurl_easy_setopt .hCurl, CURLOPT_WRITEFUNCTION, AddressOf WriteFunction
        vbcurl_easy_setopt .hCurl, CURLOPT_DEBUGFUNCTION, AddressOf DebugFunction
        vbcurl_easy_setopt .hCurl, CURLOPT_VERBOSE, True
        
        If followRedirects Then vbcurl_easy_setopt .hCurl, CURLOPT_FOLLOWLOCATION, 1
        If totalDLTimeout > 0 Then vbcurl_easy_setopt .hCurl, CURLOPT_TIMEOUT, totalDLTimeout
        If connectTimeout > 0 Then vbcurl_easy_setopt .hCurl, CURLOPT_CONNECTTIMEOUT, connectTimeout
        If Len(Referrer) > 0 Then vbcurl_easy_setopt .hCurl, CURLOPT_REFERER, Referrer
        If Len(UserAgent) > 0 Then vbcurl_easy_setopt .hCurl, CURLOPT_USERAGENT, UserAgent
        If Len(cookie) > 0 Then vbcurl_easy_setopt .hCurl, CURLOPT_COOKIE, cookie
        If FileExists(caBundle) Then vbcurl_easy_setopt .hCurl, CURLOPT_CAINFO, caBundle
    
        .CurlReturnCode = vbcurl_easy_perform(.hCurl)
        .Finalize
        
    End With
    
    Set resp = Nothing
    
End Function


Private Function DebugFunction(ByVal info As curl_infotype, ByVal rawBytes As Long, ByVal numBytes As Long, ByVal extra As Long) As Long

    Dim tmp As String, i As Long
    Dim b() As Byte
    
    If info >= 3 Then 'DATA_IN/OUT msgs
        vbcurl_easy_setopt resp.hCurl, CURLOPT_DEBUGFUNCTION, 0 'should we turn off dbg msgs now? less callbacks to us slow down
        Exit Function
    End If
    
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

'extra is the objPtr(activeClassObject)
Private Function WriteFunction(ByVal rawBytes As Long, ByVal sz As Long, ByVal nmemb As Long, ByVal extra As Long) As Long
    
    On Error Resume Next
    Dim totalBytes As Long, i As Long, b() As Byte, ret As CURLcode, v As Variant
    
    totalBytes = sz * nmemb
    
    If resp.isMemFile Then
        resp.memFile.memAppendBuf rawBytes, totalBytes
    Else
        ReDim b(totalBytes - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, totalBytes
        Put resp.hFile, , b()
    End If
    
    DoEvents
    WriteFunction = resp.notifyParent(totalBytes)  'if this returns 0 dl is aborted...
    
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
        
        
'I would prefer the below: (and no tlb requirement)
'there isnt really a point in using the multi downloads though, unless one server was super slow
'out of several, your bandwidth is limited anyway so just do one at a time should be fine.
'rarely worth the extra complexity for most projects unless your literally making a multi downloader app.
'-----------------------------------------------
'Dim responses As New Collection 'of CCurlResponse
'
'Function getObj(ptr As Long) As CCurlResponse
'    Dim c As CCurlResponse
'    For Each c In responses
'        If ObjPtr(c) = ptr Then
'            Set getObj = c
'            Exit Function
'        End If
'    Next
'End Function
'
'Dim resp As CCurlResponse
'Set resp = getObj(extra)
'If resp Is Nothing Then Exit Function
