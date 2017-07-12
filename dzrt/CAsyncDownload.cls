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
Attribute VB_Exposed = True
Event DownloadComplete(fPath As String)
Event Progress(current As Long, total As Long)
Event Error(code As Long, msg As String)

Dim WithEvents dl As ucAsyncDownload
Attribute dl.VB_VarHelpID = -1
Dim myForm As frmAsyncDownload

Property Get LastUrl() As String
    LastUrl = dl.LastUrl
End Property

Sub AbortDownload()
    dl.AbortDownload
End Sub

Sub StartDownload(ByVal url As String, Optional opt As AsyncReadConstants)
    dl.StartDownload url, opt
End Sub

Private Sub Class_Initialize()
    Set myForm = New frmAsyncDownload
    Set dl = myForm.dl
End Sub

Private Sub dl_DownloadComplete(fPath As String)
    RaiseEvent DownloadComplete(fPath)
End Sub

Private Sub dl_Error(code As Long, msg As String)
    RaiseEvent Error(code, msg)
End Sub

Private Sub dl_Progress(current As Long, total As Long)
    RaiseEvent Progress(current, total)
End Sub
