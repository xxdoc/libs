Attribute VB_Name = "modUtil"
Option Explicit
Global Const LANG_US = &H409
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function GetFreeFileName(ByVal folder As String, Optional extension = ".txt") As String
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(folder) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
again:
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetParentFolder(path) As String
    Dim tmp() As String
    Dim my_path
    Dim ub As String
    
    On Error GoTo hell
    If Len(path) = 0 Then Exit Function
    
    my_path = path
    While Len(my_path) > 0 And Right(my_path, 1) = "\"
        my_path = Mid(my_path, 1, Len(my_path) - 1)
    Wend
    
    tmp = Split(my_path, "\")
    tmp(UBound(tmp)) = Empty
    my_path = Replace(Join(tmp, "\"), "\\", "\")
    If VBA.Right(my_path, 1) = "\" Then my_path = Mid(my_path, 1, Len(my_path) - 1)
    
    GetParentFolder = my_path
    Exit Function
    
hell:
    GetParentFolder = Empty
    
End Function

Function objKeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    Set t = Nothing
    objKeyExistsInCollection = True
 Exit Function
nope: objKeyExistsInCollection = False
End Function

Function GetFolderFiles(folderPath As String, Optional filter As String = "*", Optional retFullPath As Boolean = True, Optional recursive As Boolean = False) As String()
   Dim fnames() As String
   Dim fs As String
   Dim folders() As String
   Dim i As Integer
   
   If Not FolderExists(folderPath) Then
        'returns empty array if fails
        GetFolderFiles = fnames()
        Exit Function
   End If
   
   folderPath = IIf(Right(folderPath, 1) = "\", folderPath, folderPath & "\")
   
   fs = Dir(folderPath & filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then push fnames(), IIf(retFullPath = True, folderPath & fs, fs)
     fs = Dir()
   Wend
   
   If recursive Then
        folders() = GetSubFolders(folderPath)
        If Not AryIsEmpty(folders) Then
            For i = 0 To UBound(folders)
                FolderEngine folders(i), fnames(), filter
            Next
        End If
        If Not retFullPath Then
            For i = 0 To UBound(fnames)
                fnames(i) = Replace(fnames(i), folderPath, Empty) 'make relative path from base
            Next
        End If
    End If
   
   GetFolderFiles = fnames()
End Function

Private Sub FolderEngine(fldrpath As String, ary() As String, Optional filter As String = "*")

    Dim files() As String
    Dim folders() As String
    Dim i As Long
     
    files = GetFolderFiles(fldrpath, filter)
    folders = GetSubFolders(fldrpath)
        
    If Not AryIsEmpty(files) Then
        For i = 0 To UBound(files)
            push ary, files(i)
        Next
    End If
    
    If Not AryIsEmpty(folders) Then
        For i = 0 To UBound(folders)
             FolderEngine folders(i), ary, filter
        Next
    End If
    
End Sub

Function GetSubFolders(folderPath As String, Optional retFullPath As Boolean = True) As String()
    Dim fnames() As String
    Dim fd As String
    
    If Not FolderExists(folderPath) Then
        'returns empty array if fails
        GetSubFolders = fnames()
        Exit Function
    End If
    
   If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

   fd = Dir(folderPath, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folderPath & fd) And vbDirectory) = vbDirectory Then
           push fnames(), IIf(retFullPath = True, folderPath & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
   GetSubFolders = fnames()
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function GetBaseName(path As String) As String
    Dim tmp() As String
    Dim ub As String
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function


