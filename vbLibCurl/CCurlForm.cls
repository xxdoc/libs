VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private hForm As Long
'there is much more than can be done with this, this is a basic implementation with room to grow.
'we could streamline it into the CCurlDownload but this might get big and it will only grow itself..

Private Sub Class_Initialize()
    hForm = vbcurl_form_create()
End Sub

Private Sub Class_Terminate()
    vbcurl_form_free hForm
End Sub

Function AddNameValueField(name As String, value As String) As CURLFORMcode
    Dim part As Long
    
    part = vbcurl_form_create_part(hForm)
    
    AddNameValueField = vbcurl_form_add_four_to_part(part, _
                            CURLFORM_COPYNAME, name, _
                            CURLFORM_COPYCONTENTS, value _
                        )
    
    vbcurl_form_add_part hForm, part
End Function

Function AddFileUpload(name As String, path As String, Optional contentType As String = "application/binary") As CURLFORMcode
    Dim part As Long
    
    part = vbcurl_form_create_part(hForm)
    
    AddFileUpload = vbcurl_form_add_six_to_part(part, _
                        CURLFORM_COPYNAME, name, _
                        CURLFORM_FILE, path, _
                        CURLFORM_CONTENTTYPE, contentType _
                    )
                    
    vbcurl_form_add_part hForm, part
    
End Function

Function Attach(hCurl As Long) As CURLcode
    Attach = vbcurl_easy_setopt(hCurl, CURLOPT_HTTPPOST, hForm)
End Function
