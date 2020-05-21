VERSION 5.00
Begin VB.UserControl ctlDownload 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   Picture         =   "ctlDownload.ctx":0000
   ScaleHeight     =   960
   ScaleWidth      =   960
   ToolboxBitmap   =   "ctlDownload.ctx":324A
End
Attribute VB_Name = "ctlDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'EXAMPLES
'Private Sub ctlDownload1_Finished(x As AsyncProperty)
'    If x.StatusCode = vbAsyncStatusCodeEndDownloadData Then
'        'to get website text
'        MainText = Replace(StrConv(x.Value, vbUnicode), vbLf, vbNewLine)
'
'    End If
'End Sub
'
'
'
'Private Sub ctlDownloadPicture_Finished(X As AsyncProperty)
'
'    Dim Bytes() As Byte, fnum As Integer
'    If X.StatusCode = vbAsyncStatusCodeEndDownloadData Then
'
'        Bytes = X.Value
'        ' Save the file.
'        fnum = FreeFile
'        Open App.Path & "\Largeimage.jpg" For Binary As #fnum
'        Put #fnum, 1, Bytes()
'        Close fnum
'
'        Erase Bytes
'
'    End If
'
'End Sub




' This component doesn't require any external calls/apis/references
' Obtains an url to a byte array using native VB6 calls
' Asyncronous - no need to wait for data to arrive
' Multiple downloads are accepted at the same time (different URL's, etc)
' If you like this code, please VOTE for it
' You may use this code freely in your projects, but whenever possible,
' include my name 'Filipe Lage' on the 'Help->About' or something ;)
' Cheers :)
'
' Filipe Lage
' fclage@ezlinkng.com
'
Public Event Zero()
Public Event Progress(X As AsyncProperty, percent As Single)
Public Event Finished(X As AsyncProperty)
Public Event Cancelled(Cancel As Boolean)
Public Event CancelCountDown(Count As Integer)
Public CurrentDownloads As New Collection

Public Function Download(xurl As String) As Boolean
    On Error Resume Next
    UserControl.AsyncRead xurl, vbAsyncTypeByteArray, xurl, vbAsyncReadForceUpdate
    CurrentDownloads.Add xurl, xurl
    RefreshStatus
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    RaiseEvent Finished(AsyncProp)
    On Error Resume Next
    CurrentDownloads.Remove AsyncProp.PropertyName
    RefreshStatus
    On Error GoTo 0
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    Dim p As Single
    If AsyncProp.BytesMax > 0 Then p = 100 * (AsyncProp.BytesRead / AsyncProp.BytesMax) Else p = 0
    RaiseEvent Progress(AsyncProp, p)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 960
    UserControl.Height = 960
End Sub

Public Sub CancelAllDownloads()
    On Error Resume Next
    Do Until CurrentDownloads.Count = 0
        DoEvents
        CancelDownload CurrentDownloads(1)
        RaiseEvent CancelCountDown(CurrentDownloads.Count)
    Loop

    RaiseEvent Cancelled(True)

    RefreshStatus
    On Error GoTo 0
End Sub


Public Sub CancelDownload(xurl As String)
    On Error Resume Next
    UserControl.CancelAsyncRead CurrentDownloads(xurl)
    CurrentDownloads.Remove xurl
    On Error GoTo 0
End Sub

Private Sub UserControl_Show()
    If UIMode = True Then
    Else
        UserControl.Extender.Visible = False
    End If
End Sub

Private Sub UserControl_Terminate()
    Do Until CurrentDownloads.Count = 0
        CancelDownload CurrentDownloads(1)
    Loop
End Sub

Private Sub RefreshStatus()
    UserControl.Cls
    UserControl.CurrentX = 0
    UserControl.CurrentY = 0
    UserControl.Print CurrentDownloads.Count
    
        If CurrentDownloads.Count = 0 Then
    RaiseEvent Zero
    End If
    
End Sub

Private Function UIMode() As Boolean
    On Error Resume Next
    Err.Clear
    Debug.Print 1 / 0
    UIMode = (Err.Number <> 0)
    On Error GoTo 0
End Function


