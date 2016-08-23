VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSystem3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Revision 4
'
'Info:     These are basically macros for VB's built in file processes
'            this should streamline your code quite a bit and hopefully
'            remove alot of redundant coding.
'
'Author:   dzzie@yahoo.com
'Sight:    http://sandsprite.com
    
'Variant version of clsFso for scripting languages, added recurse option to getfolderfiles
' 5.21.12 - bugfix in RandomNum sporotic overflow time of day related...
' 11.14.13- bugfix in FileExists,FolderExists, GetFolderFiles recursive filter, GetParentFolder trailing \ bugfix
'  1.21.14- bugfix deletefile detects readonly attribute and removes
'  7.8.16 - bugfixs synced with current CFileSystem2

Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public dlg As New CCmnDlg

Function GetFolderFiles(folderPath, Optional Filter = "*", Optional retFullPath = True, Optional recursive = False) As Collection
   Dim fnames As New Collection
   Dim folders As Collection
   Dim tmp As New Collection
   Dim i As Long
   Dim fs, f
   
   Set GetFolderFiles = fnames
   If Not FolderExists(folderPath) Then Exit Function
   folderPath = IIf(Right(folderPath, 1) = "\", folderPath, folderPath & "\")
   
   fs = Dir(folderPath & Filter, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   While fs <> ""
     If fs <> "" Then fnames.Add , IIf(retFullPath, folderPath & fs, fs)
     fs = Dir()
   Wend
   
   If recursive Then
        Set folders = GetSubFolders(folderPath)
        For Each f In folders
            FolderEngine f, tmp, Filter
        Next
        If Not retFullPath Then
            For Each f In tmp
                fnames.Add Replace(f, folderPath, Empty) 'make relative path from base
            Next
        End If
    End If
    
End Function

Private Sub FolderEngine(fldrpath, c As Collection, Optional Filter = "*")

    Dim files As Collection
    Dim folders As Collection
    Dim i As Long
    Dim f
    
    Set files = GetFolderFiles(fldrpath, Filter)
    Set folders = GetSubFolders(fldrpath)
        
    For Each f In files
        c.Add f
    Next
    
    For Each f In folders
        FolderEngine f, c, Filter
    Next
    
End Sub
 
    
Function GetSubFolders(folderPath, Optional retFullPath = True) As Collection
   Dim fnames As New Collection
   Dim fd
    
   Set GetSubFolders = fnames
    
   If Not FolderExists(folderPath) Then Exit Function
   If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

   fd = Dir(folderPath, vbDirectory)
   While fd <> ""
     If Left(fd, 1) <> "." Then
        If (GetAttr(folderPath & fd) And vbDirectory) = vbDirectory Then
           fnames.Add IIf(retFullPath = True, folderPath & fd, fd)
        End If
     End If
     fd = Dir()
   Wend
   
End Function

Function FolderExists(path)
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path)
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function GetParentFolder(path)
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

Function CreateFolder(path)
   On Error GoTo blah
   If FolderExists(path) Then Exit Function
   MkDir path
   If Not FolderExists(path) Then Exit Function
   CreateFolder = True
blah:
End Function

Function FileNameFromPath(fullpath)
    Dim tmp()
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function WebFileNameFromPath(fullpath)
    Dim tmp()
    If InStr(fullpath, "/") > 0 Then
        tmp = Split(fullpath, "/")
        WebFileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function DeleteFile(fpath)
 On Error GoTo hadErr
    
    Dim attributes As VbFileAttribute

    attributes = GetAttr(fpath)
    If (attributes And vbReadOnly) Then
        attributes = attributes - vbReadOnly
        SetAttr fpath, attributes
    End If
    
    Kill fpath
    DeleteFile = True
 Exit Function
hadErr:
'MsgBox "DeleteFile Failed" & vbCrLf & vbCrLf & fpath
DeleteFile = False
End Function

Sub Rename(fullpath, newName)
  Dim pf
  pf = GetParentFolder(fullpath)
  Name fullpath As pf & "\" & newName
End Sub

Sub SetAttribute(fpath, it As VbFileAttribute)
   SetAttr fpath, it
End Sub

Function GetExtension(path)
    Dim tmp()
    Dim ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetExtension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       GetExtension = ""
    End If
End Function

Function GetBaseName(path)
    Dim tmp()
    Dim ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    If InStr(1, ub, ".") > 0 Then
       GetBaseName = Mid(ub, 1, InStrRev(ub, ".") - 1)
    Else
       GetBaseName = ub
    End If
End Function

Function ChangeExt(path, ext)
    Dim bn
    ext = IIf(Left(ext, 1) = ".", ext, "." & ext)
    If FileExists(path) Then
        Rename path, GetBaseName(path) & ext
    Else
        'hack to just accept a file name might not be worth supporting
        bn = Mid(path, 1, InStr(1, path, ".") - 1)
        ChangeExt = bn & ext
    End If
End Function

Function SafeFileName(proposed)
  Dim badChars, bad(), i As Long
  badChars = ">,<,&,/,\,:,|,?,*,"""
  bad = Split(badChars, ",")
  For i = 0 To UBound(bad)
    proposed = Replace(proposed, bad(i), "")
  Next
  SafeFileName = CStr(proposed)
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

Function GetFreeFileName(parentDir, Optional extension = ".txt")
    
    
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String
    Dim folder As String
    
    folder = parentDir
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

Function GetFreeFolderName(ByVal parentFolder, Optional prefix = "")
    On Error GoTo handler 'can have overflow err once in awhile :(
    Dim i As Integer
    Dim tmp As String

    If Not FolderExists(parentFolder) Then Exit Function
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"
        
again:
    Do
      tmp = parentFolder & prefix & RandomNum()
    Loop Until Not FolderExists(tmp)
    
    GetFreeFolderName = tmp
    
Exit Function
handler:

    If i < 10 Then
        i = i + 1
        GoTo again
    End If
    
End Function

Function buildPath(folderPath)
    On Error GoTo oops
    
    If FolderExists(folderPath) Then buildPath = True: Exit Function
    
    Dim tmp(), build, i As Long
    
    tmp = Split(folderPath, "\")
    build = tmp(0)
    For i = 1 To UBound(tmp)
        build = build & "\" & tmp(i)
        If InStr(tmp(i), ".") < 1 Then
            If Not FolderExists(build) Then CreateFolder (build)
        End If
    Next
    buildPath = True
    Exit Function
oops: buildPath = False
End Function


Function ReadFile(filename) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub


Function Copy(fpath, toFolder)
   Dim baseName, newName
   If FolderExists(toFolder) Then
       baseName = FileNameFromPath(fpath)
       toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
       newName = toFolder & baseName
       FileCopy fpath, newName
       Copy = newName
   Else 'assume tofolder is actually new desired file path
       FileCopy fpath, toFolder
       Copy = toFolder
   End If
End Function

Function Move(fpath, toFolder)
    Dim fName As String
    fName = FileNameFromPath(fpath)
    toFolder = IIf(Right(toFolder, 1) = "\", toFolder, toFolder & "\")
    
    Name fpath As toFolder & fName
    Move = toFolder & fName
    
End Function

Function CreateFile(fpath)
    On Error GoTo hell
    Dim f As Long
    f = FreeFile
    If FileExists(fpath) Then Exit Function
    Open fpath For Binary As f
    Close f
    If FileExists(fpath) Then CreateFile = True
hell:
End Function


Function DeleteFolder(folderPath, Optional force = True)
 On Error GoTo failed
   Call delTree(folderPath, force)
   Call RmDir(folderPath)
   DeleteFolder = True
 Exit Function
failed:  DeleteFolder = False
End Function

Private Sub delTree(folderPath, Optional force = True)
   Dim sfi(), sfo(), i As Integer
   sfi() = GetFolderFiles(folderPath)
   sfo() = GetSubFolders(folderPath)
   If Not AryIsEmpty(sfi) And force = True Then
        For i = 0 To UBound(sfi)
            DeleteFile sfi(i)
        Next
   End If
   
   If Not AryIsEmpty(sfo) And force = True Then
        For i = 0 To UBound(sfo)
            Call DeleteFolder(sfo(i), True)
        Next
   End If
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function FolderName(folderPath)
    Dim ret, tmp()
    tmp = Split(folderPath, "\")
    If Not AryIsEmpty(tmp) Then
        If Len(tmp(UBound(tmp))) <> 0 Then ret = tmp(UBound(tmp)) _
        Else ret = tmp(UBound(tmp) - 1)
    Else
        ret = CStr(folderPath)
    End If
    FolderName = ret
End Function

Private Sub Class_Initialize()
    'If Not isRegistered And Not isInitalized Then TellThemAllAboutIt
End Sub


