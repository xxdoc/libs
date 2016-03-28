VERSION 5.00
Begin VB.UserControl Zipper 
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Zipper.ctx":0000
   ScaleHeight     =   975
   ScaleWidth      =   1065
   ToolboxBitmap   =   "Zipper.ctx":03CD
   Begin VB.Timer tmrNext 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   420
      Top             =   420
   End
End
Attribute VB_Name = "Zipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Zipper
'
'A UserControl wrapping ZipWriter to Zip multiple files into a new
'or existing archive.
'
'Usage:
'
'   Call the AddFile() method multiple times to build the list of
'   files to add to the ZIP archive.
'
'   Then call the Zip() method to add the files.
'
'   Events are raised to report progress, ending with the Complete
'   event.  On completion the list is empty and you can build a
'   new list of files and then call Zip() again.
'
'   There are a number of properties you can use when the events
'   are raised to describe the error or current progress.
'

'Performance tuning constants:
Private Const ZIP_CHUNK As Long = 131072  '2 * 65536, could be smaller or larger than this.
                                          'Larger is generally faster but eats more RAM.

Private Const SUB_TICKS As Integer = 10   'Number of blocks to process in a Timer tick.
                                          'Up to a point, larger is faster but UI becomes
                                          'less responsive.

Private ZipWriter As ZipWriter
Private FilesToZip As Collection
Private FNum As Integer
Private Buf() As Byte
Private BytesLeftInFile As Long

Private mBytesToZip As Long
Private mBytesZipped As Long
Private mCancel As Boolean
Private mCurrentFile As ZipFile
Private mFailed As String
Private mResult As ZIP_RESULTS
Private mZipping As Boolean

Public Event Complete(ByVal Canceled As Boolean)
Public Event EndFile()
Public Event Error()
Public Event Progress()
Public Event StartFile()

Public Property Get BytesToZip() As Long
    'This value gets counted up by AddFile() calls and cleared
    'to 0 at the end of Zip().  It is valid before and during
    'Zip().
    BytesToZip = mBytesToZip
End Property

Public Property Get BytesZipped() As Long
    'This value gets counted up during Zip() runs and cleared
    'to 0 at the start of Zip().  It is valid during and after
    'a call to Zip().
    BytesZipped = mBytesZipped
End Property

Public Property Get CurrentFile() As ZipFile
    'Valid during and after Zip().
    Set CurrentFile = mCurrentFile
End Property

Public Property Get Failed() As String
    'Failed ZipWriter call.
    Failed = mFailed
End Property

Public Property Get Result() As ZIP_RESULTS
    'Result of failed ZipWriter call.
    Result = mResult
End Property

Public Property Get Zipping() As Boolean
    Zipping = mZipping
End Property

Public Function AddFile( _
    ByVal SourceFile As String, _
    Optional ByVal AsFile As String = "", _
    Optional ByVal ZMethod As Z_METHODS = Z_DEFLATED, _
    Optional ByVal ZLevel As Z_LEVELS = Z_DEFAULT_COMPRESSION, _
    Optional ByVal Attrs As VbFileAttribute = vbNormal, _
    Optional ByVal Comment As String = "") As Boolean
    'Returns True on failure.
    
    Dim FileAttrs As VbFileAttribute
    Dim File As ZipFile
    
    On Error Resume Next
    FileAttrs = GetAttr(SourceFile)
    If Err Then
        Err.Clear
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
        .ByteCount = FileLen(SourceFile)
        mBytesToZip = mBytesToZip + .ByteCount
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

Public Sub Cancel()
    If mZipping Then mCancel = True
End Sub

Public Function Zip( _
    ByVal FilePath As String, _
    Optional ByVal AppendMode As APPEND_MODES = APPEND_STATUS_CREATE) As Boolean
    'Returns True on failure.

    If mZipping Then
        'This error does not stop running Zip!
        mResult = 0
        mFailed = "Already mZipping"
        Zip = True
    Else
        mBytesZipped = 0
        With ZipWriter
            If .OpenZip(FilePath, AppendMode) Then
                mFailed = "OpenZip"
                mResult = .Result
                ClearFiles
                Zip = True
            Else
                mZipping = True
                tmrNext.Enabled = True
            End If
        End With
    End If
End Function

Private Sub tmrNext_Timer()
    Dim SubTicksLeft As Integer
    
    tmrNext.Enabled = False
    
    If FNum = 0 Then
        If FilesToZip.Count > 0 Then
            With ZipWriter
                Set mCurrentFile = FilesToZip.Item(1)
                RaiseEvent StartFile
                FilesToZip.Remove 1
                If .OpenFileInZip(mCurrentFile.AsFile, _
                                  mCurrentFile.ZMethod, _
                                  mCurrentFile.ZLevel, _
                                  mCurrentFile.Attrs, _
                                  mCurrentFile.Comment) Then
                    mFailed = "OpenFileInZip"
                    GoTo ErrorExit '<---<< Early exit!
                Else
                    FNum = FreeFile(0)
                    Open mCurrentFile.SourceFile For Binary Access Read As #FNum
                    BytesLeftInFile = mCurrentFile.ByteCount
                    If BytesLeftInFile < ZIP_CHUNK Then
                        ReDim Buf(BytesLeftInFile - 1)
                    Else
                        ReDim Buf(ZIP_CHUNK - 1)
                    End If
                End If
            End With
        Else
            If ZipWriter.CloseZip() Then
                mFailed = "CloseZip"
                GoTo ErrorExit '<---<< Early exit!
            End If
            Erase Buf
            mBytesToZip = 0
            mZipping = False
            RaiseEvent Complete(mCancel)
            mCancel = False
            Exit Sub '<---<< Early exit!
        End If
    End If
    
    If mCancel Then
        ClearFiles
        BytesLeftInFile = 0
    End If
    
    With ZipWriter
        SubTicksLeft = SUB_TICKS
        Do While SubTicksLeft > 0
            SubTicksLeft = SubTicksLeft - 1
            If BytesLeftInFile > 0 Then
                If BytesLeftInFile < ZIP_CHUNK Then
                    If UBound(Buf) <> BytesLeftInFile - 1 Then
                        ReDim Buf(BytesLeftInFile - 1)
                    End If
                    mBytesZipped = mBytesZipped + BytesLeftInFile
                    BytesLeftInFile = 0
                Else
                    mBytesZipped = mBytesZipped + ZIP_CHUNK
                    BytesLeftInFile = BytesLeftInFile - ZIP_CHUNK
                End If
                Get #FNum, , Buf
                If .WriteBytes(Buf) Then
                    Close #FNum
                    FNum = 0
                    mFailed = "WriteBytes"
                    GoTo ErrorExit '<---<< Early exit!
                End If
                RaiseEvent Progress
            Else
                Close #FNum
                FNum = 0
                If .CloseFileInZip() Then
                    mFailed = "CloseFileInZip"
                    GoTo ErrorExit '<---<< Early exit!
                End If
                RaiseEvent EndFile
                Exit Do '<---<< Early exit!
            End If
        Loop
    End With
    
    tmrNext.Enabled = True
    Exit Sub

ErrorExit:
    'Clean up, report error, exit with Timer disabled:
    mResult = ZipWriter.Result
    ClearFiles
    Erase Buf
    mBytesToZip = 0
    mZipping = False
    RaiseEvent Error
End Sub

Private Sub UserControl_Initialize()
    Set ZipWriter = New ZipWriter
    Set FilesToZip = New Collection
End Sub

Private Sub UserControl_Resize()
    Width = 420
    Height = 420
End Sub
