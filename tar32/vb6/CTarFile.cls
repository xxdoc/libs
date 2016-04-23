VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTarFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'vb6 Class for tar32.dll by Yoshioka Tsuneo(tsuneo@rr.iij4u.or.jp)
'
'you can pack tar/tar.gz/gz/bz2 and unpack tar/tar.gz/tar.Z/tar.bz2/gz/Z/bz2.
'unpacking tar/tar.gz/tar.Z is auto-detect.
'
'homepage:
'   http://openlab.ring.gr.jp/tsuneo/tar32/index-e.html
'
'CTarFile class by David Zimmer <dzzie@yahoo.com>
'repo: https://github.com/dzzie/libs/tree/master/tar32
'
'we use byval so client can specify variants as well

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function TarGetVersion Lib "tar32" () As Integer
Private Declare Function tar Lib "tar32" Alias "Tar" (ByVal hwnd As Long, ByVal cmd As String, ByVal buf As String, ByVal bufLen As Long) As Long
Private Declare Function TarCheckArchive Lib "tar32" (ByVal szArcFile As String, ByVal mode As Long) As Long
Private Declare Function TarOpenArchive Lib "tar32" (ByVal hwnd As Long, ByVal fname As String, ByVal mode As Long) As Long
Private Declare Function TarCloseArchive Lib "tar32" (ByVal handle As Long) As Long
Private Declare Function TarFindFirst Lib "tar32" (ByVal hwnd As Long, ByVal wildCard_NOT_IMPLEMENTED As String, ByRef ii As INDIVIDUALINFO) As Long
Private Declare Function TarFindNext Lib "tar32" (ByVal hwnd As Long, ByRef ii As INDIVIDUALINFO) As Long
Private Declare Function TarGetFileCount Lib "tar32" (ByVal szArcFile As String) As Long

Private Type INDIVIDUALINFO
    dwOriginalSize As Long
    dwCompressedSize As Long
    dwCRC As Long
    uFlag As Long
    uOSType As Long
    wRatio As Integer
    wDate As Integer
    wTime As Integer
    szFileName As String * 513
    dummy1 As String * 3
    szAttribute As String * 8
    szMode As String * 8
    'safety As String * 10
End Type

Public Enum ArchiveTypes
    atTar = 0
    atTgz = 1
    atTarBZ2 = 2
End Enum

Dim hLib As Long
Public LastError As String
Public ShowDialog As Boolean
Public KeepDirectoryStructure As Boolean

Property Get isInitilized() As Boolean
    isInitilized = (hLib <> 0)
End Property

Property Get isValidArchive(ByVal archive As String) As Boolean

    If hLib = 0 Then Exit Property
    
    If Not FileExists(archive) Then
        LastError = "File not found"
        Exit Property
    End If
    
    isValidArchive = (TarCheckArchive(archive, 0) = 1)
    
    If Not isValidArchive Then
        LastError = "Not a valid archive"
        Exit Property
    End If
    
End Property

Property Get FileCount(ByVal pth As String) As Long
    If hLib = 0 Then Exit Property
    FileCount = TarGetFileCount(pth)
End Property

Property Get Version() As String
    If hLib = 0 Then Exit Property
    Version = TarGetVersion()
End Property

Function RawCmd(ByVal cmd As String, ByRef outputBuffer) As Long
    If hLib = 0 Then Exit Function
    outputBuffer = String(1000, Chr(0))
    RawCmd = tar(0, cmd, outputBuffer, LenB(outputBuffer))
End Function

Function ExtractTo(ByVal archive As String, ByVal targetDir As String) As Boolean
    
    If Not isValidArchive(archive) Then Exit Function
        
    Dim cmd As String
    
    cmd = "-xvf """ & archive & """ -o """ & targetDir & """"
    
    If Not ShowDialog Then cmd = cmd & " --display-dialog=0"
    If Not KeepDirectoryStructure Then cmd = cmd & " --use-directory=0"
    
    buf = String(1000, Chr(0))
    If tar(0, cmd, buf, LenB(buf)) = 0 Then
        ExtractTo = True
    Else
        LastError = Replace(buf, vbLf, vbCrLf)
    End If
    
End Function

'library is not honoring toDir...always extracts to app.path..chdir() no help, must manually rename..
Function ExtractFile(ByVal archive As String, ByVal file As String, ByVal toDir As String) As Boolean
        
        Dim buf As String
        Dim dst As String
        
        If Not isValidArchive(archive) Then Exit Function
        
        If Not FolderExists(toDir) Then
            LastError = "Invalid destination directory"
            Exit Function
        End If
        
        If Right(toDir, 1) <> "\" Then toDir = toDir & "\"
        
        cmd = "-xvf """ & archive & """ """ & file & """" ' -o """ & toDir & """"
        If Not ShowDialog Then cmd = cmd & " --display-dialog=0"
        If Not KeepDirectoryStructure Then cmd = cmd & " --use-directory=0"
        
        buf = String(1000, Chr(0))
        If tar(0, cmd, buf, LenB(buf)) = 0 Then
            buf = FileNameFromPath(file)
            dst = toDir & "\" & buf
            If FileExists(buf) Then
                 If FileExists(dst) Then Kill dst
                 Name buf As dst
                 ExtractFile = True
            End If
        Else
            LastError = Replace(buf, vbLf, vbCrLf)
        End If
        
End Function

Function EnumFiles(ByVal archive As String, Optional ByVal match As String = "*.*") As Collection
    
    Dim ii As INDIVIDUALINFO
    Dim h As Long, ret As Long, f As String
    Dim c As New Collection
    
    If Not isValidArchive(archive) Then Exit Function
    
    h = TarOpenArchive(0, archive, 0)
    If h = 0 Then
        LastError = "Failed to open archive"
        Exit Function
    End If
    
    ret = TarFindFirst(h, Empty, ii)
    While ret <> -1
        f = chomp(ii.szFileName)
        If match = "*.*" Then
            c.Add f
        Else
            If f Like match Then c.Add f
        End If
        ret = TarFindNext(h, ii)
    Wend
    
    TarCloseArchive h
    Set EnumFiles = c
    
End Function

Private Function chomp(x) As String
    a = InStr(x, Chr(0))
    If a = 1 Then Exit Function
    If a > 1 Then
        chomp = Mid(x, 1, a - 1)
    Else
        chomp = x
    End If
End Function

Function ArchiveFromFolder(ByVal folder As String, ByVal toArchive As String, Optional match As String = "*", Optional aType As ArchiveTypes = atTgz) As Boolean

    Dim buf As String, cmd As String
    Dim flags()
    
    flags = Array("-cvf", "-cvfz", "-cvfB")
    
    If FileExists(toArchive) Then Kill toArchive
    
    If aType < 0 Or aType > UBound(flags) Then
        LastError = "Invalid archive type"
        Exit Function
    End If
    
    If Not FolderExists(folder) Then
        LastError = "Folder not found"
        Exit Function
    End If
    
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    
    cmd = flags(aType)
    cmd = cmd & " """ & toArchive & """ """ & folder & match & """"
       
    If Not ShowDialog Then cmd = cmd & " --display-dialog=0"
    If Not KeepDirectoryStructure Then cmd = cmd & " --use-directory=0"

    buf = String(1000, Chr(0))
    If tar(0, cmd, buf, LenB(buf)) = 0 Then
        ArchiveFromFolder = True
    Else
        LastError = Replace(buf, vbLf, vbCrLf)
    End If
    
End Function

Function ArchiveFromFiles(fileAryorCol, ByVal toArchive As String, Optional aType As ArchiveTypes = atTgz) As Boolean

    Dim buf As String, cmd As String
    Dim flags(), files(), f
    
    flags = Array("-cvf", "-cvfz", "-cvfB")
    
    If TypeName(fileAryorCol) = "Collection" Or IsArray(fileAryorCol) Then
        For Each f In fileAryorCol
            If Not FileExists(f) Then
                LastError = "File not found: " & f
                Exit Function
            End If
            push files, f
        Next
    ElseIf TypeName(fileAryorCol) = "String" Or TypeName(fileAryorCol) = "Variant" Then
        If Not FileExists(fileAryorCol) Then
            LastError = "File not found: " & fileAryorCol
            Exit Function
        End If
        push files, fileAryorCol
    Else
        LastError = "Unsupported argument must specify single file path, array, or collection"
        Exit Function
    End If
    
    If FileExists(toArchive) Then Kill toArchive
    
    If aType < 0 Or aType > UBound(flags) Then
        LastError = "Invalid archive type"
        Exit Function
    End If
    
    If AryIsEmpty(files) Then
        LastError = "Must specify files to compress, this can be a single file path, array, or collection"
        Exit Function
    End If
        
    cmd = flags(aType)
    cmd = cmd & " """ & toArchive & """ "
    
    For Each f In files
        cmd = cmd & """" & f & """ "
    Next
       
    If Not ShowDialog Then cmd = cmd & " --display-dialog=0"
    If Not KeepDirectoryStructure Then cmd = cmd & " --use-directory=0"

    buf = String(1000, Chr(0))
    If tar(0, cmd, buf, LenB(buf)) = 0 Then
        ArchiveFromFiles = True
    Else
        LastError = Replace(buf, vbLf, vbCrLf)
    End If
    
End Function


Function FileExists(ByVal path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function



Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function FolderExists(ByVal path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Private Function FileNameFromPath(ByVal fullpath) As String
    fullpath = Replace(fullpath, "/", "\")
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub Class_Initialize()

    hLib = LoadLibrary("tar32.dll")
    If hLib = 0 Then h = LoadLibrary(App.path & "\tar32.dll")
    If hLib = 0 Then h = LoadLibrary(App.path & "\..\tar32.dll")
    If hLib = 0 Then LastError = "tar32.dll not found?"
    
    KeepDirectoryStructure = True
    ShowDialog = False
      
End Sub
