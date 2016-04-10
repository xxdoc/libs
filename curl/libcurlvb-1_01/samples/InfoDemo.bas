' $Id: InfoDemo.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' Extract info from a web site

Attribute VB_Name = "InfoDemo"

Private Sub Test()
    InfoDemo "http://www.drudgereport.com"
End Sub

Private Sub InfoDemo(url As String)
    Dim easy As Long, i As Long
    Dim ret As CURLcode

   ' This needs to be a Variant!
    Dim info As Variant
    
    easy = vbcurl_easy_init()
    vbcurl_easy_setopt easy, CURLOPT_URL, url
    vbcurl_easy_setopt easy, CURLOPT_FILETIME, True
    ret = vbcurl_easy_perform(easy)

    vbcurl_easy_getinfo easy, CURLINFO_CONNECT_TIME, info
    Debug.Print "Connect Time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_CONTENT_LENGTH_DOWNLOAD, info
    Debug.Print "Content Length (Download): " & info
    vbcurl_easy_getinfo easy, CURLINFO_CONTENT_LENGTH_UPLOAD, info
    Debug.Print "Content Length (Upload): " & info
    vbcurl_easy_getinfo easy, CURLINFO_CONTENT_TYPE, info
    Debug.Print "Content Type: " & info
    vbcurl_easy_getinfo easy, CURLINFO_EFFECTIVE_URL, info
    Debug.Print "Effective URL: " & info
    vbcurl_easy_getinfo easy, CURLINFO_FILETIME, info
    Debug.Print "File time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_HEADER_SIZE, info
    Debug.Print "Header Size: " & info
    vbcurl_easy_getinfo easy, CURLINFO_HTTPAUTH_AVAIL, info
    Debug.Print "Authentication Bitmask: " & info
    vbcurl_easy_getinfo easy, CURLINFO_HTTP_CONNECTCODE, info
    Debug.Print "HTTP Connect Code: " & info
    vbcurl_easy_getinfo easy, CURLINFO_NAMELOOKUP_TIME, info
    Debug.Print "Name Lookup Time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_OS_ERRNO, info
    Debug.Print "OS Errno: " & info
    vbcurl_easy_getinfo easy, CURLINFO_PRETRANSFER_TIME, info
    Debug.Print "Pretransfer time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_PROXYAUTH_AVAIL, info
    Debug.Print "Proxy Authentication Schemes: " & info
    vbcurl_easy_getinfo easy, CURLINFO_REDIRECT_COUNT, info
    Debug.Print "Redirect Count: " & info
    vbcurl_easy_getinfo easy, CURLINFO_REDIRECT_TIME, info
    Debug.Print "Redirect time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_REQUEST_SIZE, info
    Debug.Print "Request Size: " & info
    vbcurl_easy_getinfo easy, CURLINFO_RESPONSE_CODE, info
    Debug.Print "Response Code: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SIZE_DOWNLOAD, info
    Debug.Print "Download size: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SIZE_UPLOAD, info
    Debug.Print "Upload size: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SPEED_DOWNLOAD, info
    Debug.Print "Download speed: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SPEED_UPLOAD, info
    Debug.Print "Upload speed: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SSL_VERIFYRESULT, info
    Debug.Print "SSL verification result: " & info
    vbcurl_easy_getinfo easy, CURLINFO_STARTTRANSFER_TIME, info
    Debug.Print "Start transfer time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_TOTAL_TIME, info
    Debug.Print "Total time: " & info
    vbcurl_easy_getinfo easy, CURLINFO_SSL_ENGINES, info
    Debug.Print "SSL Engines:"
    For i = 0 To UBound(info)
        Debug.Print info(i)
    Next
End Sub

