VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAsyncDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public url As String
Public toFile As String
Public isSSL As Long
Public hThread As Long


 'these can not be public or its converted to a get/let
 'these are written to directly by mem address in the C downloader thread
Private abort As Long
Dim m_Progress As Long
Dim dlSize As Long
Dim status_code As Long
Dim cnt As Long
Dim statusReported As Boolean
Dim retVal As Long

Private Declare Function BackgroundDownload Lib "bgdl.dll" ( _
    ByVal server As String, _
    ByVal webpath As String, _
    ByVal isSSL As Long, _
    ByVal path As String, _
    ByRef setValOnComplete As Long, _
    ByRef abortVar As Long, _
    ByRef StatusCode As Long, _
    ByRef dlSize As Long, _
    ByRef Progress As Long _
) As Long

Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Const STILL_ACTIVE = &H103

Enum eDownloadStatus
    ds_idle = -2
    ds_Downloading = -1
    ds_Complete = 1
    ds_DownloadError = 2
    ds_UserAborted = 3
End Enum

Property Get DownloadStatus() As eDownloadStatus
    If abort = 1 Then retVal = ds_UserAborted 'probably overwritten in cleanup didnt check just override now
    DownloadStatus = retVal
End Property

Property Get DownloadStatusStr() As String
    If abort = 1 Then retVal = ds_UserAborted
    If retVal = ds_UserAborted Then DownloadStatusStr = "User Aborted"
    If retVal = ds_idle Then DownloadStatusStr = "Idle"
    If retVal = ds_Complete Then DownloadStatusStr = "Complete"
    If retVal = ds_Downloading Then DownloadStatusStr = "Downloading"
    If retVal = ds_DownloadError Then DownloadStatusStr = "Download Error"
End Property

Property Get StatusCode() As Long 'http status code
    StatusCode = status_code
End Property

Property Get Progress() As Long 'how many bytes downloaded
    Progress = m_Progress
End Property

Property Get DownloadSize() As Long 'not all servers return this
    DownloadSize = dlSize
End Property

Sub AbortDownload()
    abort = 1
    retVal = ds_UserAborted
End Sub

Function isRunning() As Boolean
    Dim exitCode As Long, rv As Long
    rv = GetExitCodeThread(hThread, exitCode)
    If rv = 0 Then Exit Function 'failed
    If exitCode = STILL_ACTIVE Then isRunning = True
End Function

Function StartDownload(url, toFile, Optional forceFreshDownload As Boolean = True) As Boolean
    
    Dim tmp, a, started As Long, server As String, webpath As String
    
    Me.url = url
    Me.toFile = toFile
    
    If forceFreshDownload Then DeleteUrlCacheEntry url
    
    'quick cheesy url parse
    tmp = url
    If InStr(1, tmp, "https://", vbTextCompare) > 0 Then isSSL = True
    tmp = Replace(tmp, "https://", Empty)
    tmp = Replace(tmp, "http://", Empty)
    
    a = InStr(1, tmp, "/")
    If a > 1 Then
        server = Mid(tmp, 1, a - 1)
        webpath = Mid(tmp, a)
    Else
        Exit Function
    End If
    
    hThread = BackgroundDownload(server, webpath, isSSL, toFile, retVal, abort, status_code, dlSize, m_Progress)
    StartDownload = (hThread <> 0)
    
End Function

Private Sub Class_Initialize()
    retVal = ds_idle
End Sub

Private Sub Class_Terminate()

    'we can not allow this class this be released until we are sure the download thread has ended
    'since it does direct memory writes to class private variable memory addresses
    
    Dim cycles As Long
    
    If isRunning() Then
        Form1.lst Hex(ObjPtr(Me)) & " Class_Terminate while thread active waiting while aborting download", 2
    End If
    
    Do While isRunning()
        cycles = cycles + 1
        abort = 1
        Sleep 10
        DoEvents
        If cycles = 1000 Then
            Form1.lst "wtf force terminating downloader thread - shouldnt ever happen..."
            TerminateThread hThread, -4
        End If
        If cycles = 1500 Then
            MsgBox "ok i give up no soup for you", vbExclamation 'just to see if this would ever happen in testing probably not
            End
        End If
    Loop
    
    If cycles > 0 Then
         Form1.lst Hex(ObjPtr(Me)) & " Class_Terminate took " & cycles & " loops to abort download and terminate thread", 2
    End If
    
End Sub
