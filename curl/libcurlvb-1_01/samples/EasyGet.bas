' $Id: EasyGet.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' Demonstrate Easy Get Capability

Attribute VB_Name = "EasyGet"

Private Sub Test()
    EasyGet ("http://www.google.com")
End Sub

Private Sub EasyGet(url As String)
    Dim easy As Long
    Dim ret As CURLcode
    Dim buf As New Buffer
    
    easy = vbcurl_easy_init()
    vbcurl_easy_setopt easy, CURLOPT_URL, url
    vbcurl_easy_setopt easy, CURLOPT_WRITEDATA, ObjPtr(buf)
    vbcurl_easy_setopt easy, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
    'vbcurl_easy_setopt easy, CURLOPT_DEBUGFUNCTION, _
    '    AddressOf DebugFunction
    'vbcurl_easy_setopt easy, CURLOPT_VERBOSE, True
    ret = vbcurl_easy_perform(easy)
    vbcurl_easy_cleanup easy
    Debug.Print "Here's the HTML:"
    Debug.Print buf.stringData
End Sub

' This function illustrates a couple of key concepts in libcurl.vb.
' First, the data passed in rawBytes is an actual memory address
' from libcurl. Hence, the data is read using the MemByte() function
' found in the VBVM6Lib.tlb type library. Second, the extra parameter
' is passed as a raw long (via ObjPtr(buf)) in Sub EasyGet()), and
' we use the AsObject() function in VBVM6Lib.tlb to get back at it.
Private Function WriteFunction(ByVal rawBytes As Long, _
    ByVal sz As Long, ByVal nmemb As Long, _
    ByVal extra As Long) As Long
    
    Dim totalBytes As Long, i As Long
    Dim obj As Object, buf As Buffer
    
    totalBytes = sz * nmemb
    
    Set obj = AsObject(extra)
    Set buf = obj
    ' append the binary characters to the HTML string
    For i = 0 To totalBytes - 1
        ' Append the write data
        buf.stringData = buf.stringData & Chr(MemByte(rawBytes + i))
    Next
    ' Need this line below since AsObject gets a stolen reference
    ObjectPtr(obj) = 0&
    
    ' Return value
    WriteFunction = totalBytes
End Function

' Again, rawBytes comes straight from libcurl and extra is a
' long, though we're not using it here.
Private Function DebugFunction(ByVal info As curl_infotype, _
    ByVal rawBytes As Long, ByVal numBytes As Long, _
    ByVal extra As Long) As Long
    Dim debugMsg As String
    Dim i As Long
    debugMsg = ""
    For i = 0 To numBytes - 1
        debugMsg = debugMsg & Chr(MemByte(rawBytes + i))
    Next
    Debug.Print "info=" & info & ", debugMsg=" & debugMsg
    DebugFunction = 0
End Function

