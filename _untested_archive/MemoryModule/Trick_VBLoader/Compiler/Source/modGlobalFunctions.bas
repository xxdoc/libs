Attribute VB_Name = "modGlobalFunctions"
Option Explicit

' // Set app icon to specified window
Public Sub SetWindowIcon( _
           ByVal hWnd As Long)
    Dim cx      As Long: Dim cy As Long
    Dim hIcon   As Long

    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)

    hIcon = LoadImage(App.hInstance, "WINICO", IMAGE_ICON, cx, cy, LR_SHARED)
    
    If hIcon Then
        SendMessage hWnd, WM_SETICON, ICON_SMALL, ByVal hIcon
    End If
    
End Sub

' // Show open file dialog and return selected file names
' // First item in the collection is path then file names
Public Function GetOpenFile( _
                ByVal hWnd As Long, _
                ByRef Title As String, _
                ByRef Filter As String, _
                ByVal Multiline As Boolean) As Collection
    Dim ofn As OPENFILENAME:    Dim out As String
    Dim ic  As Long:            Dim io  As Long
    Dim p   As Long
    
    ofn.Flags = IIf(Multiline, OFN_ALLOWMULTISELECT, 0) Or OFN_EXPLORER
    ofn.nMaxFile = 32767
    out = String(32767, vbNullChar)

    ofn.hwndOwner = hWnd
    ofn.lpstrTitle = StrPtr(Title)
    ofn.lpstrFile = StrPtr(out)
    ofn.lStructSize = Len(ofn)
    ofn.lpstrFilter = StrPtr(Filter)
    
    If GetOpenFileName(ofn) Then
    
        Set GetOpenFile = New Collection
        GetOpenFile.Add Left$(out, ofn.nFileOffset - 1)
        io = ofn.nFileOffset + 1: p = p + 1
        ic = InStr(io, out, vbNullChar)
        
        Do Until ic = io
        
            GetOpenFile.Add Mid$(out, io, ic - io)
            io = ic + 1: p = p + 1
            ic = InStr(io, out, vbNullChar)
            
        Loop
        
    End If
    
End Function

' // Show save dialog and return selected file name
Public Function GetSaveFile( _
                ByVal hWnd As Long, _
                ByRef Title As String, _
                ByRef Filter As String, _
                ByRef DefExt As String) As String
    Dim ofn As OPENFILENAME:    Dim out As String
    Dim i   As Long
    
    ofn.nMaxFile = 260
    ofn.Flags = OFN_OVERWRITEPROMPT Or OFN_EXPLORER
    out = String(260, vbNullChar)
    ofn.hwndOwner = hWnd
    ofn.lpstrTitle = StrPtr(Title)
    ofn.lpstrFile = StrPtr(out)
    ofn.lStructSize = Len(ofn)
    ofn.lpstrFilter = StrPtr(Filter)
    ofn.lpstrDefExt = StrPtr(DefExt)
    
    If GetSaveFileName(ofn) Then
    
        i = InStr(1, out, vbNullChar, vbBinaryCompare)
        If i Then GetSaveFile = Left$(out, i - 1)
        
    End If
    
End Function

' // Get file name by full path
Public Function GetFileTitle( _
                ByRef sPath As String, _
                Optional ByRef UseExtension As Boolean = False) As String
    Dim l As Long, p As Long
    
    l = InStrRev(sPath, "\")
    
    If UseExtension Then p = Len(sPath) + 1 Else p = InStrRev(sPath, ".")
    
    If p > l Then
        l = IIf(l = 0, 1, l + 1)
        GetFileTitle = Mid$(sPath, l, p - l)
    ElseIf p = l Then
        GetFileTitle = sPath
    Else
        GetFileTitle = Mid$(sPath, l + 1)
    End If
    
End Function

' // Get folder name from full path
Public Function GetFilePath( _
                ByRef sPath As String) As String
    Dim l As Long, p As Long
    
    l = InStrRev(sPath, "\")
    If l = Len(sPath) Or l = 0 Then GetFilePath = sPath: Exit Function
    GetFilePath = Mid$(sPath, 1, l)
    
End Function

' // Get file extension
Public Function GetFileExtension( _
                ByRef sPath As String) As String
    Dim i1  As Long, i2 As Long
    
    i1 = InStrRev(sPath, ".")
    i2 = InStrRev(sPath, "\")
    
    If i1 > 0 And i2 < i1 Then GetFileExtension = Mid$(sPath, i1)
    
End Function

