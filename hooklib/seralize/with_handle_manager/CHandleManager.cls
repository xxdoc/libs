VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHandleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public handles As New Collection

Sub Add(fPath As String, sysHandle As Long)
    Dim h As New CFileHandle
    h.file = fPath
    h.sysHandle = sysHandle
    handles.Add h
End Sub

Function HandleExists(sysHandleorPath As Variant) As Boolean
    HandleExists = Not GetHandle(sysHandleorPath) Is Nothing
End Function

Function GetHandle(sysHandleorPath As Variant) As CFileHandle
    Dim h As CFileHandle
    Dim isHandle As Boolean
    
    isHandle = TypeName(sysHandleorPath) = "Long"
    
    For Each h In handles
        If isHandle Then
            If h.sysHandle = sysHandleorPath Then
                Set GetHandle = h
                Exit Function
            End If
        Else
            If h.file = sysHandleorPath Then
                Set GetHandle = h
                Exit Function
            End If
        End If
    Next
End Function

Sub Remove(sysHandle As Long)
    For i = 1 To handles.Count
        If handles(i).sysHandle = sysHandle Then
            handles.Remove i
            Exit Sub
        End If
    Next
End Sub

Function StartRedir(sysHandleorPath As Variant, redirTo As RedirTypes, Optional target As String) As Boolean
    Dim f As CFileHandle
    If Not HandleExists(sysHandleorPath) Then Exit Function
    Set f = GetHandle(sysHandleorPath)
    f.RedirectTo = redirTo
    f.target = target
    StartRedirection = True
End Function

Sub StopRedir(sysHandleorPath As Variant)
    Dim f As CFileHandle
    If Not HandleExists(sysHandleorPath) Then Exit Sub
    Set f = GetHandle(sysHandleorPath)
    f.RedirectTo = Empty
End Sub
