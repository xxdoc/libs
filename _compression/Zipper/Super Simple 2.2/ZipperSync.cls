VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipperSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'ZipperSync version 2.2
'
'A synchronous Class new with version 2.1 duplicating the functionality
'of Zipper, a UserControl wrapping ZipWriter to Zip multiple files into
'a new or existing archive.
'
'Usage:
'
'   Call the AddFile() method multiple times to build the list of
'   files to add to the ZIP archive.
'
'   Then call the Zip() method to add the files.
'
'   There are a number of properties you can use after failures to get
'   a description of the error.
'
'Changes:
'
'   Version 2.2:
'
'   Additions:
'
'       * AddBLOB() method, allows a Byte array to be added to the
'         Zip archive as a file without creating a temporary file.
'
'       Pretty unbelievable set of bugs due to a hurried and poorly
'       tested rewrite of the async Zipper.ctl!  Fixes:
'
'       * Added missing call to ZipWriter.CloseZip() method in the
'         Zip() method below.
'
'       * If UseLastModified = False on call to AddFile() method below,
'         use Now() value instead.  Was being left at the default value
'         (i.e. CDate(0.0)).
'
'       * In Zip() method below within the file data loop that calls:
'
'                           .WriteBytes(Buf)
'
'         ... on error make sure file FNum is still open before trying
'         to close it!
'
'       * GetSourceFileBuf() WAS NOT BEING SET = True at EOF!
'
'       * Change that same .WriteBytes() loop from Do While to Do Until!
'

'Performance tuning constants:
Private Const ZIP_CHUNK As Long = 131072  '2 * 65536, could be smaller or larger than this.
                                          'Larger is generally faster but eats more RAM.

Private ZipWriter As ZipWriter
Private FilesToZip As Collection
Private FNum As Integer
Private Buf() As Byte
Private BytesLeftInFile As Long
Private ZipDateTimes As ZipDateTimes
Private NeedCRC As Boolean
Private CRC As Long
Private BytesZipped As Long
Private CurrentFile As ZipFile

Private mFailed As String
Private mResult As ZIP_RESULTS

Public Property Get Failed() As String
    'Failed ZipWriter call.
    Failed = mFailed
End Property

Public Property Get Result() As ZIP_RESULTS
    'Result of failed ZipWriter call.
    Result = mResult
End Property

Public Sub AddBLOB( _
    ByRef BLOB() As Byte, _
    ByVal AsFile As String, _
    Optional ByVal ZMethod As Z_METHODS = Z_DEFLATED, _
    Optional ByVal ZLevel As Z_LEVELS = Z_DEFAULT_COMPRESSION, _
    Optional ByVal Comment As String = "", _
    Optional ByVal Password As String = "")
    
    Dim File As ZipFile
    
    Set File = New ZipFile
    With File
        .LetBLOB BLOB
        .AsFile = AsFile
        .ZMethod = ZMethod
        .ZLevel = ZLevel
        .Comment = Comment
        .FileDate = Now()
        If Len(Password) Then
            .Password = Password
            .NeedCRC = True
        End If
        .ByteCount = UBound(BLOB) - LBound(BLOB) + 1
    End With
    FilesToZip.Add File
End Sub

Public Function AddFile( _
    ByVal SourceFile As String, _
    Optional ByVal AsFile As String = "", _
    Optional ByVal ZMethod As Z_METHODS = Z_DEFLATED, _
    Optional ByVal ZLevel As Z_LEVELS = Z_DEFAULT_COMPRESSION, _
    Optional ByVal Attrs As VbFileAttribute = vbNormal, _
    Optional ByVal Comment As String = "", _
    Optional ByVal UseLastModified As Boolean, _
    Optional ByVal Password As String = "") As Boolean
    'Returns True on failure.
    
    Dim FileAttrs As VbFileAttribute
    Dim File As ZipFile
    
    On Error Resume Next
    FileAttrs = GetAttr(SourceFile)
    If Err Then
        Err.Clear
        mFailed = "File does not exist"
        AddFile = True
        Exit Function
    End If
    On Error GoTo 0
    
    If Len(AsFile) = 0 Then
        AsFile = Mid$(SourceFile, InStrRev(SourceFile, "\") + 1)
    End If
    Set File = New ZipFile
    With File
        .SourceFile = SourceFile
        .AsFile = AsFile
        .ZMethod = ZMethod
        .ZLevel = ZLevel
        If Attrs = 0 Then
            .Attrs = FileAttrs
        Else
            .Attrs = Attrs
        End If
        .Comment = Comment
        If UseLastModified Then
            On Error Resume Next
            ZipDateTimes.GetTimestamps SourceFile
            If Err Then
                Err.Clear
                mFailed = "GetTimestamps"
                AddFile = True
                Exit Function
            End If
            On Error GoTo 0
            .FileDate = ZipDateTimes.LastModified
        Else
            .FileDate = Now()
        End If
        If Len(Password) Then
            .Password = Password
            .NeedCRC = True
        End If
        .ByteCount = FileLen(SourceFile)
    End With
    FilesToZip.Add File
End Function

Public Sub ClearFiles()
    'Only needs to be called upon a failure of Zip() or to
    'abandon FilesToZip before calling Zip().
    Dim I As Long
    
    With FilesToZip
        For I = .Count To 1 Step -1
            .Remove I
        Next
    End With
End Sub

Public Function Zip( _
    ByVal FilePath As String, _
    Optional ByVal AppendMode As APPEND_MODES = APPEND_STATUS_CREATE) As Boolean
    'Returns True on failure.
    
    Dim BLOBMode As Boolean
    Dim BLOB() As Byte

    BytesZipped = 0
    With ZipWriter
        If .OpenZip(FilePath, AppendMode) Then
            mFailed = "OpenZip"
            GoTo ErrorExit '<---<< Early exit!
        Else
            Do While FilesToZip.Count > 0
                Set CurrentFile = FilesToZip.Item(1)
                FilesToZip.Remove 1
                If Len(CurrentFile.SourceFile) = 0 Then
                    BLOBMode = True
                    BLOB = CurrentFile.GetBLOB()
                Else
                    BLOBMode = False
                End If
                If CurrentFile.NeedCRC Then
                    .OpenCrc
                    If BLOBMode Then
                        .CrcBytes BLOB
                    Else
                        OpenSourceFile
                        Do While GetSourceFileBuf()
                            .CrcBytes Buf
                        Loop
                    End If
                    CRC = .CloseCrc()
                    CurrentFile.NeedCRC = False
                End If
                If .OpenFileInZip(CurrentFile.AsFile, _
                                  CurrentFile.ZMethod, _
                                  CurrentFile.ZLevel, _
                                  CurrentFile.Attrs, _
                                  CurrentFile.Comment, _
                                  CurrentFile.FileDate, _
                                  CurrentFile.Password, _
                                  CRC) Then
                    mFailed = "OpenFileInZip"
                    GoTo ErrorExit '<---<< Early exit!
                End If
                If BLOBMode Then
                    If .WriteBytes(BLOB) Then
                        mFailed = "WriteBytes"
                        GoTo ErrorExit '<---<< Early exit!
                    End If
                    Erase BLOB
                Else
                    OpenSourceFile
                    Do Until GetSourceFileBuf()
                        If .WriteBytes(Buf) Then
                            If FNum <> 0 Then
                                Close #FNum
                                FNum = 0
                            End If
                            mFailed = "WriteBytes"
                            GoTo ErrorExit '<---<< Early exit!
                        End If
                    Loop
                End If
                If .CloseFileInZip() Then
                    mFailed = "CloseFileInZip"
                    GoTo ErrorExit '<---<< Early exit!
                End If
            Loop
        End If
        If .CloseZip() Then
            mFailed = "CloseZip"
            GoTo ErrorExit '<---<< Early exit!
        End If
    End With
    Exit Function
    
ErrorExit:
    mResult = ZipWriter.Result
    ClearFiles
    Erase Buf
    Zip = True
End Function

Private Sub OpenSourceFile()
    FNum = FreeFile(0)
    Open CurrentFile.SourceFile For Binary Access Read As #FNum
    BytesLeftInFile = CurrentFile.ByteCount
    If BytesLeftInFile < ZIP_CHUNK Then
        ReDim Buf(BytesLeftInFile - 1)
    Else
        ReDim Buf(ZIP_CHUNK - 1)
    End If
End Sub

Private Function GetSourceFileBuf() As Boolean
    'Returns True on EOF, closes file.
    If BytesLeftInFile > 0 Then
        If BytesLeftInFile < ZIP_CHUNK Then
            If UBound(Buf) <> BytesLeftInFile - 1 Then
                ReDim Buf(BytesLeftInFile - 1)
            End If
            If Not NeedCRC Then BytesZipped = BytesZipped + BytesLeftInFile
            BytesLeftInFile = 0
        Else
            If Not NeedCRC Then BytesZipped = BytesZipped + ZIP_CHUNK
            BytesLeftInFile = BytesLeftInFile - ZIP_CHUNK
        End If
        Get #FNum, , Buf
    Else
        Close #FNum
        FNum = 0
        GetSourceFileBuf = True
    End If
End Function

Private Sub Class_Initialize()
    Set ZipWriter = New ZipWriter
    Set FilesToZip = New Collection
    Set ZipDateTimes = New ZipDateTimes
End Sub

