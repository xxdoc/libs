VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'ZipWriter
'
'A wrapper class for zlibwapi.dll, making use of the parts of
'the MiniZip project that it includes:
'
'   http://www.winimage.com/zLibDll/minizip.html
'
'This class can open and append to an existing Zip archive or
'create a new one and then open and write files into it with
'or without Deflate compression.
'

Private Const API_NULL As Long = 0

Private Enum Z_RESULTS
    Z_VERSION_ERROR = -6
    Z_BUF_ERROR = -5
    Z_MEM_ERROR = -4
    Z_DATA_ERROR = -3
    Z_STREAM_ERROR = -2
    Z_ERRNO = -1
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
End Enum

Public Enum ZIP_RESULTS
    'In typical "open source" C/C++ programming fashion there is
    'much that goes poorly documented, undocumented, or left
    'open to interpretation due to the crudity of the C and to
    'a lesser extent C++ languages themselves.
    '
    'I am hazarding a guess that these are the result codes
    'for calls to the entrypoints in zlibwapi.dll that we
    'are using here.

    ZIP_INTERNALERROR = -104
    ZIP_BADZIPFILE = -103
    ZIP_PARAMERROR = -102
    ZIP_OK = 0
    ZIP_EOF = 0
    ZIP_ERRNO = Z_ERRNO
    
    'Extensions for callers checking the Result property:
    
    ZIP_NOT_OPEN = -666
    ZIP_ALREADY_OPEN = -667
End Enum

Public Enum APPEND_MODES
    APPEND_STATUS_CREATE = 0
    [_APPEND_STATUS_CREATEAFTER] = 1 'Hide this member, seldom used.
    APPEND_STATUS_ADDINZIP = 2
End Enum

Private Declare Function zipOpen Lib "zlibwapi" ( _
    ByVal ppathname As Long, _
    ByVal append As APPEND_MODES) As Long

Private Type tm_zip
    'These fields seem absurdly long as Longs.  However in the
    'original C header file these are defined as "uInt" which
    'by all guesses ought to be unsigned 32-bit integer values.

    tm_sec As Long            'seconds after the minute - [0,59]
    tm_min As Long            'minutes after the hour - [0,59]
    tm_hour As Long           'hours since midnight - [0,23]
    tm_mday As Long           'day of the month - [1,31]
    tm_mon As Long            'months since January - [0,11]
    tm_year As Long           'years - [1980..2044]
End Type

Private Type zip_fileinfo
    'More typical "open source" C programming rears its ugly head
    'here.  The comment below:
    '
    '       if dos_date == 0, tmu_date is used
    '
    '... probably ought to say something like:
    '
    '       if dosDate = 0, tmz_date is used
    '
    'Oh well, that's the assumption I'm relying on and it seems
    'to work.

    tmz_date As tm_zip        'date in understandable format
    dosDate As Long           'if dos_date == 0, tmu_date is used

    'Next line (flag) is commented in the original C header file.
    'flag as long             'general purpose bit flag        2 bytes
    internal_fa As Long       'internal file attributes        2 bytes

    'Next line (external_fa) I have substituted the intrinsic VB
    'Enum as the type.
    'external_fa As Long      'external file attributes        4 bytes
    external_fa As VbFileAttribute
End Type

Public Enum Z_METHODS
    Z_STORE = 0               'Store without compressing.
    Z_DEFLATED = 8
End Enum

Public Enum Z_LEVELS
    'Your results may vary.   See these comments:
    Z_DEFAULT_COMPRESSION = -1
                              'On many files gives same compression as
                              'Z_BEST_COMPRESSION but runs faster.
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1          'On many files seems to take about the same time
                              'Z_DEFAULT_COMPRESSION but produces *better*
                              'compression than Z_BEST_COMPRESSION.
    Z_BEST_COMPRESSION = 9    'On many files slower than Z_DEFAULT_COMPRESSION
                              'but only gives same compression.
End Enum

Private Declare Function zipOpenNewFileInZip Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pfilename As Long, _
    ByRef zipfi As zip_fileinfo, _
    ByVal pextrafield_local As Long, _
    ByVal size_extrafield_local As Long, _
    ByVal pextrafield_global As Long, _
    ByVal size_extrafield_global As Long, _
    ByVal pcomment As Long, _
    ByVal method As Z_METHODS, _
    ByVal level As Z_LEVELS) As ZIP_RESULTS

Private Declare Function zipWriteInFileInZip Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pbuf As Long, _
    ByVal length As Long) As ZIP_RESULTS

Private Declare Function zipCloseFileInZip Lib "zlibwapi" ( _
    ByVal hfile As Long) As ZIP_RESULTS

Private Declare Function zipClose Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pglobal_comment As Long) As ZIP_RESULTS

Private Writing As Boolean 'True = "file in Zip" is open.
Private hZip As Long
Private mResult As ZIP_RESULTS

Public Property Get Result() As ZIP_RESULTS
    Result = mResult
End Property

Public Function OpenZip( _
    ByVal FilePath As String, _
    Optional ByVal AppendMode As APPEND_MODES = APPEND_STATUS_CREATE) As Boolean
    'Returns True on failure.
    '
    'FilePath can be a simple file name with .ZIP extension or
    'a complex path relative to CD or a full path.
    '
    'Examples:
    '
    '   D:\SomeArchive.zip
    '   SomeArchive.zip
    '   SomeFolder\SomeArchive.zip
    If hZip <> API_NULL Then
        mResult = ZIP_ALREADY_OPEN
        OpenZip = True
    Else
        FilePath = StrConv(FilePath, vbFromUnicode)
        hZip = zipOpen(StrPtr(FilePath), AppendMode)
        If hZip = API_NULL Then
            mResult = ZIP_ERRNO
            OpenZip = True
        Else
            mResult = ZIP_OK
        End If
    End If
End Function

Public Function OpenFileInZip( _
    ByVal FileName As String, _
    Optional ByVal ZMethod As Z_METHODS = Z_DEFLATED, _
    Optional ByVal ZLevel As Z_LEVELS = Z_DEFAULT_COMPRESSION, _
    Optional ByVal Attrs As VbFileAttribute = vbNormal, _
    Optional ByVal Comment As String = "") As Boolean
    'Returns True on failure.
    '
    'FileName may consist of a simple file name or a relative
    'path, e.g. "SomeFolder\SomeFile.txt" creates the file
    'within the folder.
    Dim ZFI As zip_fileinfo
    Dim FileDate As Date
    
    If hZip = API_NULL Then
        mResult = ZIP_NOT_OPEN
        OpenFileInZip = True
    ElseIf Writing Then
        mResult = ZIP_ALREADY_OPEN
        OpenFileInZip = True
    Else
        FileName = StrConv(FileName, vbFromUnicode)
        With ZFI
            With .tmz_date
                FileDate = Now()
                .tm_sec = Second(FileDate)
                .tm_min = Minute(FileDate)
                .tm_hour = Hour(FileDate)
                .tm_mday = Day(FileDate)
                .tm_mon = Month(FileDate) - 1 'Odd but true.
                .tm_year = Year(FileDate)
            End With
            .external_fa = Attrs And (vbNormal _
                                   Or vbReadOnly _
                                   Or vbHidden _
                                   Or vbSystem _
                                   Or vbArchive)
        End With
        Comment = StrConv(Comment, vbFromUnicode)
        mResult = zipOpenNewFileInZip(hZip, _
                                      StrPtr(FileName), _
                                      ZFI, _
                                      0, _
                                      0, _
                                      0, _
                                      0, _
                                      StrPtr(Comment), _
                                      ZMethod, _
                                      ZLevel)
        If mResult = ZIP_OK Then
            Writing = True
        Else
            OpenFileInZip = True
        End If
    End If
End Function

Public Function WriteBytes(ByRef Bytes() As Byte) As Boolean
    'Returns True on failure.
    
    If (hZip = API_NULL) Or (Not Writing) Then
        mResult = ZIP_NOT_OPEN
        WriteBytes = True
    Else
        mResult = zipWriteInFileInZip(hZip, _
                                      VarPtr(Bytes(LBound(Bytes))), _
                                      UBound(Bytes) - LBound(Bytes) + 1)
        If mResult <> ZIP_OK Then
            WriteBytes = True
        End If
    End If
End Function

Public Function CloseFileInZip() As Boolean
    'Returns True on failure.
    
    If Not Writing Then
        mResult = ZIP_NOT_OPEN
        CloseFileInZip = True
    Else
        mResult = zipCloseFileInZip(hZip)
        If mResult = ZIP_OK Then
            Writing = False
        End If
    End If
End Function

Public Function CloseZip() As Boolean
    'Returns True on failure.
    
    If hZip = API_NULL Then
        mResult = ZIP_NOT_OPEN
        CloseZip = True
    Else
        If Writing Then
            'Caller forgot to call CloseFileInZip() first.  Do it here
            'implicitly.  We could return an error result but why make
            'life hard?
            mResult = zipCloseFileInZip(hZip)
            If mResult = ZIP_OK Then
                Writing = False
            Else
                CloseZip = True
                Exit Function
            End If
        End If
        
        mResult = zipClose(hZip, API_NULL)
        If mResult = ZIP_OK Then
            hZip = API_NULL
        End If
    End If
End Function

Private Sub Class_Terminate()
    If Writing Then
        zipCloseFileInZip hZip
    End If
    If hZip <> API_NULL Then
        zipClose hZip, API_NULL
    End If
End Sub
