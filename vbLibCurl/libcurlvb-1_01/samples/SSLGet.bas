' $Id: SSLGet.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' SSLGet.bas - demonstrate trivial SSL Get capability

Attribute VB_Name = "SSLGet"

Private Sub Test()
    SSLGet "https://sourceforge.net", "d:\vblibcurl\bin\ca-bundle.crt"
End Sub

Private Sub SSLGet(url As String, caFile As String)
    Dim context As Long
    Dim ret As Long
    Dim buf As New Buffer
    
    context = vbcurl_easy_init()
    
    vbcurl_easy_setopt context, CURLOPT_URL, url
    vbcurl_easy_setopt context, CURLOPT_CAINFO, caFile
    vbcurl_easy_setopt context, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
    vbcurl_easy_setopt context, CURLOPT_WRITEDATA, ObjPtr(buf)
    vbcurl_easy_setopt context, CURLOPT_NOPROGRESS, 0
    vbcurl_easy_setopt context, CURLOPT_PROGRESSFUNCTION, _
        AddressOf ProgressFunction
    
    'vbcurl_easy_setopt context, CURLOPT_SSL_CTX_FUNCTION, _
    '    AddressOf SSLContextFunction
        
    ret = vbcurl_easy_perform(context)
    vbcurl_easy_cleanup context
    Debug.Print "Here's the SSL HTML:"
    Debug.Print buf.stringData
End Sub

' See WriteFunction() in EasyGet.bas for more detailed explanation.
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

Private Function ProgressFunction(ByVal extra As Long, _
    ByVal dlTotal As Double, ByVal dlNow As Double, _
    ByVal ulTotal As Double, ByVal ulNow As Double) As Long
    ' just print the data
    Debug.Print "dlTotal=" & dlTotal & ", dlNow=" & dlNow & _
        ", ulTotal=" & ulTotal & ", ulNow=" & ulNow
    ProgressFunction = 0
End Function

' The context parameter is an OpenSSL SSL_CTX pointer, if you
' care to muck with OpenSSL from VB.
Private Function SSLContextFunction(ByVal context As Long, _
    ByVal extra As Long) As Long
    Debug.Print "In SSLContextFunction"
    SSLContextFunction = 0
End Function

