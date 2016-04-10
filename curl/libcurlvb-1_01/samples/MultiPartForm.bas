' $Id: MultiPartForm.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' MultiPartForm.bas - demonstrate RFC 1867 Multi-Part Form Capability
' This is a case where we diverge a bit from libcurl. For forms, we
' have an intermediate part object that gets created with the call
' vbcurl_form_create_part() and updated with the functions:
'
'     vbcurl_form_add_pair_to_part()
'     vbcurl_form_add_four_to_part()
'     vbcurl_form_add_six_to_part()
'
' After you've built the "part", then you call the function
' vbcurl_form_add_part(), which calls curl_formadd() internally.
' The reason I decomposed curl_formadd() this way was that I
' couldn't get a variable argument list passed from VB to
' work properly. This is a bit inelegant, I know.
 
Attribute VB_Name = "MultiPartForm"

' just use a global variable for the write function
Dim strResponse As String

Private Sub MultiPartForm()
    Dim context As Long
    Dim ret As CURLcode
    Dim form As Long, part As Long
    
    context = vbcurl_easy_init()
    form = vbcurl_form_create()
    
    part = vbcurl_form_create_part(form)
    vbcurl_form_add_four_to_part part, CURLFORM_COPYNAME, _
        "frmUsername", CURLFORM_COPYCONTENTS, "userName"
    vbcurl_form_add_part form, part
    
    part = vbcurl_form_create_part(form)
    vbcurl_form_add_four_to_part part, CURLFORM_COPYNAME, _
        "frmPassword", CURLFORM_COPYCONTENTS, "userPwd"
    vbcurl_form_add_part form, part
    
    part = vbcurl_form_create_part(form)
    vbcurl_form_add_four_to_part part, CURLFORM_COPYNAME, _
        "frmFileOrigPath", CURLFORM_COPYCONTENTS, _
        "d:\temp\atl71.dll"
    vbcurl_form_add_part form, part
    
    part = vbcurl_form_create_part(form)
    vbcurl_form_add_four_to_part part, CURLFORM_COPYNAME, _
        "frmFileDate", CURLFORM_COPYCONTENTS, "08/01/2004"
    vbcurl_form_add_part form, part
    
    part = vbcurl_form_create_part(form)
    vbcurl_form_add_six_to_part part, CURLFORM_COPYNAME, _
        "f1", CURLFORM_FILE, "d:\temp\atl71.dll", _
        CURLFORM_CONTENTTYPE, "application/binary"
    vbcurl_form_add_part form, part
    
    vbcurl_easy_setopt context, CURLOPT_URL, _
        "http://www.mysite.net/FormPage.asp"
    vbcurl_easy_setopt context, CURLOPT_HTTPPOST, form
    vbcurl_easy_setopt context, CURLOPT_WRITEFUNCTION, _
        AddressOf WriteFunction
    vbcurl_easy_setopt context, CURLOPT_NOPROGRESS, 0
    vbcurl_easy_setopt context, CURLOPT_PROGRESSFUNCTION, _
        AddressOf ProgressFunction
        
    ret = vbcurl_easy_perform(context)
    
    vbcurl_form_free (form)
    vbcurl_easy_cleanup (context)

    Debug.Print "Here's the Response:"
    Debug.Print strResponse
End Sub

Private Function WriteFunction(ByVal rawBytes As Long, _
    ByVal sz As Long, ByVal nmemb As Long, _
    ByVal extra As Long) As Long
    
    Dim totalBytes As Long, i As Long
    
    totalBytes = sz * nmemb
    ' append the binary characters to the HTML string
    For i = 0 To totalBytes - 1
        ' Append the response data global variable
        strResponse = strResponse & Chr(MemByte(rawBytes + i))
    Next
    
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

