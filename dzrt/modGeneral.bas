Attribute VB_Name = "modGeneral"
Option Explicit

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'just a useful bonus method to add into the C dll might as well...
Declare Function ut_crc32 Lib "utypes.dll" Alias "crc32" (ByRef b As Byte, ByVal sz As Long) As Long
Declare Function ut_crc32w Lib "utypes.dll" Alias "crc32w" (ByVal b As Long, ByVal sz As Long) As Long

Public hUTypes As Long
 
Function ensureUTypes() As Boolean
    
    On Error Resume Next
    
    If hUTypes <> 0 Then
        ensureUTypes = True
        Exit Function
    End If
    
    Dim pth As String, b() As Byte, f As Long
    
    pth = App.path & "\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\..\UTypes.dll"
    
    If Not FileExists(pth) Then

        pth = App.path & "\UTypes.dll"
        b() = LoadResData("UTYPES", "DLLS")
        If AryIsEmpty(b) Then
            MsgBox "Failed to find UTypes.dll in resource?"
            Exit Function
        End If
        
        f = FreeFile
        Open pth For Binary As f
        Put f, , b()
        Close f
        
        'MsgBox "Dropped utypes.dll to: " & pth & " - Err: " & Err.Number
    End If
    
    If Not FileExists(pth) Then
        MsgBox "Failed to write UTypes.dll to disk from resource?"
        Exit Function
    End If
        
    hUTypes = LoadLibrary(pth)
        
    If hUTypes = 0 Then
        MsgBox "Failed to load UTypes.dll library?"
        Exit Function
    End If
    
    ensureUTypes = True
    
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x
  
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
