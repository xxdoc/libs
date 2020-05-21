VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MiniZipDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========
'MiniZipDir
'==========
'
'Minimal interface to minilib (in zlibwapi.dll) for querying the directory in a Zip archive.
'
'Put zlibwapi.dll into the Project's (EXE's) folder.
'
'NOTES
'
'   o Paths to Zip files to be queried must be absolute, full-qualified paths.  It doesn't
'     find them if you try to look up by relative path.
'
'   o The String values in and out are in theory ANSI but may well actually be codepage 437
'     DOS OEM extended ASCII.  However a lot of Zip archive creation software probably just
'     jams ANSI in there under the current codepage.
'
'     In theory there is also UTF-8 support but that probably is only true for leter versions
'     of the Zip format.  Or perhaps only for Zip64 format?
'
'   o Directory separators come back as "/" and not "\" but we could easily replace those
'     here.  I'm not sure why they do this though.  I had assumed Zip format originated on
'     MS-DOS, but true or not perhaps "/" was put in back then to be universal, or perhaps
'     this is something that "minilib" is doing?
'

Private Const WIN32_NULL As Long = 0

Private Const MAX_PATH As Long = 260

Private Const Z_ERRNO As Long = -1

Private Const UNZ_NULL As Long = 0

Public Enum UnzipResults
    UNZ_OK = 0
    UNZ_END_OF_LIST_OF_FILE = -100
    UNZ_ERRNO = Z_ERRNO
    UNZ_EOF = 0
    UNZ_PARAMERROR = -102
    UNZ_BADZIPFILE = -103
    UNZ_INTERNALERROR = -104
    UNZ_CRCERROR = -105
    UNZ_ALREADYOPEN = -500 'Private extension.
    UNZ_OPENFAILED = -501 'Private extension.
    UNZ_NOTOPEN = -502 'Private extension.
End Enum
#If False Then
Dim UNZ_OK, UNZ_END_OF_LIST_OF_FILE, UNZ_ERRNO, UNZ_EOF, UNZ_PARAMERROR, UNZ_BADZIPFILE
Dim UNZ_INTERNALERROR, UNZ_CRCERROR, UNZ_ALREADYOPEN
#End If

'These UDTs have not been verified.  They should be correct according to unzip.h but
'the byte count values and item sizes are weird.  Perhaps those are how many bytes
'are significant?

Private Type tm_unz
    tm_sec As Long                  'seconds after the minute - [0,59]
    tm_min As Long                  'minutes after the hour - [0,59]
    tm_hour As Long                 'hours since midnight - [0,23]
    tm_mday As Long                 'day of the month - [1,31]
    tm_mon As Long                  'months since January - [0,11]
    tm_year As Long                 'years - [1980..2044]
End Type

Private Type unz_file_info
    version As Long                 'version made by                 2 bytes
    version_needed As Long          'version needed to extract       2 bytes
    flag As Long                    'general purpose bit flag        2 bytes
    compression_method As Long      'compression method              2 bytes
    dosDate As Long                 'last mod file date in Dos fmt   4 bytes
    crc As Long                     'crc-32                          4 bytes
    compressed_size As Long         'compressed size                 4 bytes
    uncompressed_size As Long       'uncompressed size               4 bytes
    size_filename As Long           'filename length                 2 bytes
    size_file_extra As Long         'extra field length              2 bytes
    size_file_comment As Long       'file comment length             2 bytes
    disk_num_start As Long          'disk number start               2 bytes
    internal_fa As Long             'internal file attributes        2 bytes
    external_fa As Long             'external file attributes        4 bytes
    tmu_date As tm_unz
End Type

Private Declare Function unzOpen Lib "zlibwapi" ( _
    ByVal Path As String) As Long
Private Declare Function unzClose Lib "zlibwapi" ( _
    ByVal unzhFile As Long) As UnzipResults
Private Declare Function unzGoToFirstFile Lib "zlibwapi" ( _
    ByVal unzhFile As Long) As UnzipResults
Private Declare Function unzGoToNextFile Lib "zlibwapi" ( _
    ByVal unzhFile As Long) As UnzipResults
Private Declare Function unzGetCurrentFileInfo Lib "zlibwapi" ( _
    ByVal unzhFile As Long, _
    ByVal pFileInfo As Long, _
    ByVal FileName As String, _
    ByVal FileNameBufferSize As Long, _
    ByVal pExtraField As Long, _
    ByVal ExtraFieldBufferSize As Long, _
    ByVal Comment As String, _
    ByVal CommentBufferSize As Long) As UnzipResults

Private unzhFile As Long 'Unzip handle to an open archive.

Private mEndOfFiles As Boolean
Private mUnzResults As UnzipResults

Public Property Get EndOfFiles() As Boolean
    EndOfFiles = mEndOfFiles
End Property

Public Property Get UnzResults() As UnzipResults
    UnzResults = mUnzResults
End Property

Public Sub CloseZip()
    Dim NewUnzResults As UnzipResults

    If unzhFile <> UNZ_NULL Then
        NewUnzResults = unzClose(unzhFile)
        If NewUnzResults <> UNZ_OK Then
            mUnzResults = NewUnzResults
            Err.Raise &H80045500, TypeName(Me), "Close failed"
        Else
            mUnzResults = UNZ_OK
            unzhFile = UNZ_NULL
            mEndOfFiles = True
        End If
    Else
        'Wasn't open, just quietly return:
        mUnzResults = UNZ_OK
    End If
End Sub

Public Function FirstFile() As String
    Dim NewUnzResults As UnzipResults

    If unzhFile = UNZ_NULL Then
        mUnzResults = UNZ_NOTOPEN
        Err.Raise &H80045508, TypeName(Me), "Not open"
    End If

    NewUnzResults = unzGoToFirstFile(unzhFile)
    If NewUnzResults = UNZ_END_OF_LIST_OF_FILE Then
        mEndOfFiles = True
    ElseIf NewUnzResults = UNZ_OK Then
        mUnzResults = UNZ_OK
        FirstFile = GetCurrentFileInfo()
    Else
        mUnzResults = NewUnzResults
        Err.Raise &H8004550C, TypeName(Me), "FirstFile failed"
    End If
End Function

Public Function NextFile() As String
    Dim NewUnzResults As UnzipResults

    If unzhFile = UNZ_NULL Then
        mUnzResults = UNZ_NOTOPEN
        Err.Raise &H80045510, TypeName(Me), "Not open"
    End If

    NewUnzResults = unzGoToNextFile(unzhFile)
    If NewUnzResults = UNZ_END_OF_LIST_OF_FILE Then
        mEndOfFiles = True
    ElseIf NewUnzResults = UNZ_OK Then
        mUnzResults = UNZ_OK
        NextFile = GetCurrentFileInfo()
    Else
        mUnzResults = NewUnzResults
        Err.Raise &H80045514, TypeName(Me), "NextFile failed"
    End If
End Function

Public Sub OpenZip(ByRef Path As String)
    Dim NewUnzhFile As Long

    If unzhFile <> UNZ_NULL Then
        mUnzResults = UNZ_ALREADYOPEN
        Err.Raise &H80045520, TypeName(Me), "Already open"
    End If

    NewUnzhFile = unzOpen(Path)
    If NewUnzhFile = UNZ_NULL Then
        mUnzResults = UNZ_OPENFAILED
        Err.Raise &H80045524, TypeName(Me), "Open failed"
    Else
        mUnzResults = UNZ_OK
        unzhFile = NewUnzhFile
    End If
End Sub

Private Function GetCurrentFileInfo() As String
    Dim Buffer As String
    Dim NewUnzResults As UnzipResults
    Dim EndPos As Long

    Buffer = Space$(MAX_PATH - 1)
    NewUnzResults = unzGetCurrentFileInfo(unzhFile, _
                                          WIN32_NULL, _
                                          Buffer, _
                                          MAX_PATH, _
                                          WIN32_NULL, _
                                          0, _
                                          vbNullString, _
                                          0)
    If NewUnzResults <> UNZ_OK Then
        mUnzResults = NewUnzResults
        Err.Raise &H80045504, TypeName(Me), "GetCurrentFileInfo failed"
    Else
        EndPos = InStr(Buffer, vbNullChar)
        If EndPos = 0 Then
            GetCurrentFileInfo = Buffer
        Else
            GetCurrentFileInfo = Left$(Buffer, EndPos - 1)
        End If
    End If
End Function

Private Sub Class_Initialize()
    mEndOfFiles = True
End Sub

Private Sub Class_Terminate()
    If unzhFile <> UNZ_NULL Then CloseZip
End Sub
