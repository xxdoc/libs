Attribute VB_Name = "modDzTest"
Option Explicit

Public hLib As Long
Public hLib2 As Long
Public debugMsg As New Collection
Public headers As New Collection

Private buf As CMemBuffer 'I dont want to include a tlb for portability...
Private fHand As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long


'if toFile is empty then we download to CMemBuffer only and return that
'else we will return the written file size as long
'if we fail to open the output file this will raise an error
Function Download(url As String, Optional toFile As String, Optional dbgMessages As Boolean)
    
    Dim easy As Long
    Dim ret As CURLcode 'enum
    'Dim buf As CMemBuffer
    
    Set buf = Nothing  'since we are only single threaded I am going to simplify with a cached global instanced...
    Set debugMsg = New Collection
    Set headers = New Collection
    
    If Len(toFile) > 0 Then
        If fso.FileExists(toFile) Then fso.DeleteFile toFile
        fHand = FreeFile
        Open toFile For Binary As fHand 'this can throw an error
    Else
        Set buf = New CMemBuffer
    End If
    
    easy = vbcurl_easy_init()
    vbcurl_easy_setopt easy, CURLOPT_URL, url
    'vbcurl_easy_setopt easy, CURLOPT_WRITEDATA, ObjPtr(buf) 'either 0 or live objptr
    vbcurl_easy_setopt easy, CURLOPT_WRITEFUNCTION, AddressOf WriteFunction
    
    If dbgMessages Then
        vbcurl_easy_setopt easy, CURLOPT_DEBUGFUNCTION, AddressOf DebugFunction
        vbcurl_easy_setopt easy, CURLOPT_VERBOSE, True
    End If
    
    ret = vbcurl_easy_perform(easy)
    vbcurl_easy_cleanup easy
    
    If buf Is Nothing Then
        Download = LOF(fHand)
        Close fHand
        fHand = 0
    Else
        Set Download = buf
        Set buf = Nothing
    End If
    
End Function


Private Function DebugFunction(ByVal info As curl_infotype, ByVal rawBytes As Long, ByVal numBytes As Long, ByVal extra As Long) As Long

    Dim tmp As String, i As Long
    Dim b() As Byte
    
    If info = CURLINFO_DATA_IN Then Exit Function '3
    
    ReDim b(numBytes - 1)
    CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, numBytes
    tmp = StrConv(b, vbUnicode, &H409)
    
    If info = CURLINFO_HEADER_IN Then
        headers.Add tmp
    End If
    
    debugMsg.Add "info: " & info & ", debugMsg:  " & tmp
    DebugFunction = 0
    
End Function

' This function illustrates a couple of key concepts in libcurl.vb.
' First, the data passed in rawBytes is an actual memory address
' from libcurl. Hence, the data is read using the MemByte() function
' found in the VBVM6Lib.tlb type library. Second, the extra parameter
' is passed as a raw long (via ObjPtr(buf)) in Sub EasyGet()), and
' we use the AsObject() function in VBVM6Lib.tlb to get back at it.

Private Function WriteFunction(ByVal rawBytes As Long, ByVal sz As Long, ByVal nmemb As Long, ByVal extra As Long) As Long
    
    Dim totalBytes As Long, i As Long
    Dim obj As Object ', buf As CMemBuffer
    Dim b() As Byte
    
    totalBytes = sz * nmemb
    'debugMsg.Add "Writing " & totalBytes & " bytes"

    'If extra = 0 Then
    If buf Is Nothing Then
        ReDim b(totalBytes - 1)
        CopyMemory ByVal VarPtr(b(0)), ByVal rawBytes, totalBytes
        Put fHand, , b()
    Else
        'Set obj = AsObject(extra)
        'Set buf = obj
        buf.memAppendBuf rawBytes, totalBytes
        'ObjectPtr(obj) = 0&
    End If
    
    DoEvents
    WriteFunction = totalBytes ' Return value
    
End Function

        
