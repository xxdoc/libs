' $Id: MultiDemo.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' MultiDemo.bas - retrieve two URLs with the multi interface

Attribute VB_Name = "MultiDemo"

Private Sub Test()
    MultiDemo "http://www.yahoo.com", "http://www.google.com"
End Sub

Private Sub MultiDemo(ByVal url1 As String, ByVal url2 As String)
    Dim e1 As Long, e2 As Long, multi As Long
    Dim ret As CURLcode
    Dim url1Copy As String, url2Copy As String
    
    url1Copy = url1
    e1 = vbcurl_easy_init()
    vbcurl_easy_setopt e1, CURLOPT_URL, url1
    ' Use StrPtr() since CURLOPT_WRITEDATA's arg is a long
    vbcurl_easy_setopt e1, CURLOPT_WRITEDATA, StrPtr(url1Copy)
    vbcurl_easy_setopt e1, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
        
    url2Copy = url2
    e2 = vbcurl_easy_init()
    vbcurl_easy_setopt e2, CURLOPT_URL, url2
    ' Use StrPtr() since CURLOPT_WRITEDATA's arg is a long
    vbcurl_easy_setopt e2, CURLOPT_WRITEDATA, StrPtr(url2Copy)
    vbcurl_easy_setopt e2, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
    
    multi = vbcurl_multi_init()
    vbcurl_multi_add_handle multi, e1
    vbcurl_multi_add_handle multi, e2
    
    Dim stillRunning As Long
    stillRunning = 1
    While vbcurl_multi_perform(multi, stillRunning) = CURLM_CALL_MULTI_PERFORM
    Wend
    
    Dim sel As Long
    While stillRunning <> 0
        vbcurl_multi_fdset multi
        sel = vbcurl_multi_select(multi, 1000)
        If (sel = 0) Then
            stillRunning = 0
        Else
            While vbcurl_multi_perform(multi, stillRunning) = CURLM_CALL_MULTI_PERFORM
            Wend
        End If
    Wend
    
    ' extract info after a multi transfer
    Dim msg As CURLMSG, easy As Long, code As CURLcode
    Dim info As Variant
    Do
        vbcurl_multi_info_read multi, msg, easy, code
        If easy = 0 Then
            Exit Do
        End If
        vbcurl_easy_getinfo easy, CURLINFO_EFFECTIVE_URL, info
        Debug.Print "vbcurl_multi_info_read: msg=" & msg & ", URL=" & _
            info & ", code=" & code
    Loop While True
    
    vbcurl_multi_cleanup multi
    vbcurl_easy_cleanup e2
    vbcurl_easy_cleanup e1
End Sub

Private Function WriteFunction(ByVal rawBytes As Long, _
    ByVal sz As Long, ByVal nmemb As Long, _
    ByVal extra As Long) As Long
    
    Dim totalBytes As Long, i As Long, strExtra As String
    
    totalBytes = sz * nmemb
    ' Get the string from the long extra parameter
    strExtra = AsString(extra)
    Debug.Print "Got " & totalBytes & " bytes from " & strExtra
    
    WriteFunction = totalBytes
End Function

