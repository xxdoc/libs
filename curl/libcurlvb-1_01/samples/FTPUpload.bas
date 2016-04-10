' $Id: FTPUpload.bas,v 1.1 2005/03/01 00:06:26 jeffreyphillips Exp $
' Demonstrate FTP Upload capability

Attribute VB_Name = "FTPUpload"
Const GENERIC_READ = &H80000000
Const FILE_ATTRIBUTE_NORMAL = &H80
Const OPEN_EXISTING = 3
Const INVALID_HANDLE_VALUE = -1

Declare Function CreateFile Lib "kernel32" _
    Alias "CreateFileA" (ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, ByVal lpBuffer As Long, _
    ByVal nBytesToRead As Long, ByRef nBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long
Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long

Private Sub Test()
    FTPUpload "d:\temp\myfile.dat", "ftp://ftp.mysite.net/myfile.dat", _
        "ftpUserName", "ftpPassword"
End Sub

Private Sub FTPUpload(fileName As String, destURL As String, _
    userID As String, password As String)
    Dim easy As Long
    Dim code As CURLcode
    Dim userPwd As String
    Dim fileSize As Long, fileSizeHigh As Long
    
    Dim fHandle As Long
    fHandle = CreateFile(fileName, GENERIC_READ, 0, 0, _
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If (fHandle = INVALID_HANDLE_VALUE) Then
        Exit Sub
    End If

    easy = vbcurl_easy_init()

    vbcurl_easy_setopt easy, CURLOPT_READFUNCTION, _
        AddressOf ReadFunction
    vbcurl_easy_setopt easy, CURLOPT_READDATA, fHandle
    vbcurl_easy_setopt easy, CURLOPT_URL, destURL
    userPwd = userID & ":" & password
    vbcurl_easy_setopt easy, CURLOPT_USERPWD, userPwd
    vbcurl_easy_setopt easy, CURLOPT_UPLOAD, 1
    fileSize = GetFileSize(fHandle, fileSizeHigh)
    vbcurl_easy_setopt easy, CURLOPT_INFILESIZE, fileSize
    vbcurl_easy_setopt easy, CURLOPT_NOPROGRESS, 0
    vbcurl_easy_setopt easy, CURLOPT_PROGRESSFUNCTION, _
        AddressOf ProgressFunction
    
    code = vbcurl_easy_perform(easy)
    vbcurl_easy_cleanup easy
    
    CloseHandle (fHandle)
End Sub

' This is where the thin libcurl.vb architecture shines! In this case,
' we've passed the handle of the opened file to upload in the extra
' parameter and the address to which to write the file data in the
' bytePtr parameter. Note that these are both used directly in the
' call to the ReadFile API function, without the need for any kind
' of intermediate processing.
Private Function ReadFunction(ByVal bytePtr As Long, ByVal sz As Long, _
    ByVal nmemb As Long, ByVal extra As Long) As Long
    Dim bytesToRead As Long, bytesRead As Long, readResult As Long
    bytesToRead = sz * nmemb
    readResult = ReadFile(extra, bytePtr, bytesToRead, bytesRead, 0)
    ReadFunction = bytesRead
End Function

' Prototype for a standard progress function
Private Function ProgressFunction(ByVal extra As Long, _
    ByVal dlTotal As Double, ByVal dlNow As Double, _
    ByVal ulTotal As Double, ByVal ulNow As Double) As Long
    ' just print the data
    Debug.Print "dlTotal=" & dlTotal & ", dlNow=" & dlNow & _
        ", ulTotal=" & ulTotal & ", ulNow=" & ulNow
    ProgressFunction = 0
End Function


