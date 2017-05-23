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
'ZipWriter version 2
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
'Version 2
'
'Adds password encyption.  Requires precalculating a CRC before
'you can call OpenFileInZip().
'
'Adds file date specification instead of default Now() value.
'
'
'Note:
'
'   While Unicode (UTF-16LE) String values are accepted here
'   they are converted to ANSI for zlibwapi.dll calls.  By
'   rights though only printable ASCII character range values
'   (32-127) should be used even though values from 1-31 and
'   128-255 generally work.  This is for file name, comment,
'   and password values.
'

Private Const Z_NULL As Long = 0

Private Const MAX_WBITS As Long = 15 '32K LZ77 window.  Normally negated on use.
Private Const DEF_MEM_LEVEL As Long = 8

Private Enum Z_STRATEGIES  'Compression strategies:
    Z_DEFAULT_STRATEGY = 0 'Normal data.
    Z_FILTERED = 1         'Data produced by a filter (or predictor).
    Z_HUFFMAN_ONLY = 2     'Force Huffman encoding only (no string match).
    Z_RLE = 3              'Limit match distances to one (run-length encoding).
                           'Often used in PNG compression.
    Z_FIXED = 4            'Prevents the use of dynamic Huffman codes, allowing
                           'for a simpler decoder for special applications.
End Enum

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
    tmz_date As tm_zip        'date in understandable format
    dosDate As Long           'if 0, tmz_date is used

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

Private Declare Function zipOpenNewFileInZip3 Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pfilename As Long, _
    ByRef zipfi As zip_fileinfo, _
    ByVal pextrafield_local As Long, _
    ByVal size_extrafield_local As Long, _
    ByVal pextrafield_global As Long, _
    ByVal size_extrafield_global As Long, _
    ByVal pcomment As Long, _
    ByVal method As Z_METHODS, _
    ByVal level As Z_LEVELS, _
    ByVal raw As Long, _
    ByVal windowBits As Long, _
    ByVal memLevel As Long, _
    ByVal strategy As Z_STRATEGIES, _
    ByVal ppassword As Long, _
    ByVal crcForCrypting As Long) As ZIP_RESULTS

Private Declare Function zipWriteInFileInZip Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pbuf As Long, _
    ByVal Length As Long) As ZIP_RESULTS

Private Declare Function zipCloseFileInZip Lib "zlibwapi" ( _
    ByVal hfile As Long) As ZIP_RESULTS

Private Declare Function zipClose Lib "zlibwapi" ( _
    ByVal hfile As Long, _
    ByVal pglobal_comment As Long) As ZIP_RESULTS

'If buf is Z_NULL, this function returns the required initial value for the crc.
Private Declare Function crc32 Lib "zlibwapi" ( _
    ByVal crcValue As Long, _
    ByVal pbuf As Long, _
    ByVal Length As Long) As Long
                   
Private Writing As Boolean 'True = "file in Zip" is open.
Private hZip As Long
Private CRC As Long 'Running value.
Private mResult As ZIP_RESULTS

Public Property Get Result() As ZIP_RESULTS
    Result = mResult
End Property

Public Sub OpenCrc()
    CRC = crc32(0, Z_NULL, 0)
End Sub

Public Sub CrcBytes(ByRef Bytes() As Byte)
    CRC = crc32(CRC, VarPtr(Bytes(LBound(Bytes))), UBound(Bytes) - LBound(Bytes) + 1)
End Sub

Public Function CloseCrc() As Long
    CloseCrc = CRC
End Function

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
    If hZip <> Z_NULL Then
        mResult = ZIP_ALREADY_OPEN
        OpenZip = True
    Else
        FilePath = StrConv(FilePath, vbFromUnicode)
        hZip = zipOpen(StrPtr(FilePath), AppendMode)
        If hZip = Z_NULL Then
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
    Optional ByVal Comment As String = "", _
    Optional ByVal FileDate As Date = #12:00:00 AM#, _
    Optional ByVal Password As String = "", _
    Optional ByVal CRC As Long = 0) As Boolean
    'Returns True on failure.
    '
    'FileName may consist of a simple file name or a relative
    'path, e.g. "SomeFolder\SomeFile.txt" creates the file
    'within the folder.
    '
    'Pass a "zero" FileDate i.e. #12:00:00 AM# to use Now().
    '
    'If Len(Password) > 0 then CRC is required.
    Dim ZFI As zip_fileinfo
    
    If hZip = Z_NULL Then
        mResult = ZIP_NOT_OPEN
        OpenFileInZip = True
    ElseIf Writing Then
        mResult = ZIP_ALREADY_OPEN
        OpenFileInZip = True
    Else
        FileName = StrConv(FileName, vbFromUnicode)
        With ZFI
            With .tmz_date
                If FileDate = #12:00:00 AM# Then
                    FileDate = Now()
                End If
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
        If Len(Password) Then
            Password = StrConv(Password, vbFromUnicode)
            mResult = zipOpenNewFileInZip3(hZip, _
                                          StrPtr(FileName), _
                                          ZFI, _
                                          0, _
                                          0, _
                                          0, _
                                          0, _
                                          StrPtr(Comment), _
                                          ZMethod, _
                                          ZLevel, _
                                          0, _
                                          -MAX_WBITS, _
                                          DEF_MEM_LEVEL, _
                                          Z_DEFAULT_STRATEGY, _
                                          StrPtr(Password), _
                                          CRC)
        Else
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
        End If
        If mResult = ZIP_OK Then
            Writing = True
        Else
            OpenFileInZip = True
        End If
    End If
End Function

Public Function WriteBytes(ByRef Bytes() As Byte) As Boolean
    'Returns True on failure.
    
    If (hZip = Z_NULL) Or (Not Writing) Then
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
    
    If hZip = Z_NULL Then
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
        
        mResult = zipClose(hZip, Z_NULL)
        If mResult = ZIP_OK Then
            hZip = Z_NULL
        End If
    End If
End Function

Private Sub Class_Terminate()
    If Writing Then
        zipCloseFileInZip hZip
    End If
    If hZip <> Z_NULL Then
        zipClose hZip, Z_NULL
    End If
End Sub
