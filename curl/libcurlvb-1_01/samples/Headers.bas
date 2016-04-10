' $Id: Headers.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' Dump the headers for a web site.

Attribute VB_Name = "Headers"

Private Sub Test()
    Headers ("http://www.google.com")
End Sub

Private Sub Headers(url As String)
    Dim easy As Long
    Dim ret As CURLcode
    Dim buf As New Buffer
    
    easy = vbcurl_easy_init()
    vbcurl_easy_setopt easy, CURLOPT_URL, url
    vbcurl_easy_setopt easy, CURLOPT_HEADERFUNCTION, _
        AddressOf HeaderFunction
    vbcurl_easy_setopt easy, CURLOPT_HEADERDATA, ObjPtr(buf)
    ret = vbcurl_easy_perform(easy)
    vbcurl_easy_cleanup easy
    Debug.Print "Here are the headers:"
    Debug.Print buf.stringData
End Sub

' The logic here is similar to that in the WriteFunction in
' EasyGet.bas.
Private Function HeaderFunction(ByVal rawBytes As Long, _
    ByVal sz As Long, ByVal nmemb As Long, _
    ByVal extra As Long) As Long

    Dim totalBytes As Long, i As Long
    Dim obj As Object, buf As Buffer
    
    totalBytes = sz * nmemb
    
    Set obj = AsObject(extra)
    Set buf = obj
    ' append the binary characters to the HTML string
    For i = 0 To totalBytes - 1
        ' Append the header data
        buf.stringData = buf.stringData & Chr(MemByte(rawBytes + i))
    Next
    ' Need this line below since AsObject gets a stolen reference
    ObjectPtr(obj) = 0&
    
    ' Return value
    HeaderFunction = totalBytes
End Function

