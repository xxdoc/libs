VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HugeBinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'HugeBinaryFile
'==============
'
'A class for doing simple binary I/O on very large disk files
'(well over the usual 2GB limit).  It only does I/O using Byte
'arrays, and makes use of Currency values that are scaled to
'whole numbers in places:
'
'   For a file of one byte the FileLen property returns 1.0000 as
'   its value.
'
'Operation is similar in many ways to native VB Get#/Put# I/O, for
'example the EOF property must be checked after a ReadBytes() call.
'You must also Dim/Redim buffers to desired sizes before calling
'ReadBytes() or WriteBytes().
'
'Short (signed Long) relative seeks and long (unsigned Currency)
'absolute seeks from 0 may be done.
'
'AutoFlush may be set True to force buffer flushes on every write.
'The Flush() method may be called explicitly if necessary.
'

Public Enum HBF_Errors
  HBF_UNKNOWN_ERROR = 45600
  HBF_FILE_ALREADY_OPEN
  HBF_OPEN_FAILURE
  HBF_SEEK_FAILURE
  HBF_FILELEN_FAILURE
  HBF_READ_FAILURE
  HBF_WRITE_FAILURE
  HBF_FILE_ALREADY_CLOSED
End Enum

Private Const HBF_SOURCE = "HugeBinaryFile"

Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1
Private Const INVALID_SET_FILE_POINTER = -1
Private Const INVALID_FILE_SIZE = -1

Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2

Private Type MungeCurr
    Value As Currency
End Type

Private Type Munge2Long
    LowVal As Long
    HighVal As Long
End Type

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Any) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpFileSizeHigh As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Private hFile As Long
Private sFName As String
Private fAutoFlush As Boolean
Private fEOF As Boolean
Private C As MungeCurr
Private L As Munge2Long

Public Property Get AutoFlush() As Boolean
    RaiseErrorIfClosed
    AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    fAutoFlush = NewVal
End Property

Public Property Get FileHandle() As Long
    RaiseErrorIfClosed
    FileHandle = hFile
End Property

Public Property Get FileLen() As Currency
    RaiseErrorIfClosed
    L.LowVal = GetFileSize(hFile, L.HighVal)
    If L.LowVal = INVALID_FILE_SIZE Then
        If Err.LastDllError Then RaiseError HBF_FILELEN_FAILURE
    End If
    LSet C = L
    FileLen = C.Value * 10000@
End Property

Public Property Get FileName() As String
    RaiseErrorIfClosed
    FileName = sFName
End Property

Public Property Get EOF() As Boolean
    RaiseErrorIfClosed
    EOF = fEOF
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Sub CloseFile()
    RaiseErrorIfClosed
    CloseHandle hFile
    sFName = ""
    fAutoFlush = False
    fEOF = False
    hFile = INVALID_HANDLE_VALUE
End Sub

Public Sub Flush()
    RaiseErrorIfClosed
    FlushFileBuffers hFile
End Sub

Public Sub OpenFile(ByVal OpenFileName As String)
    If hFile <> INVALID_HANDLE_VALUE Then
        RaiseError HBF_FILE_ALREADY_OPEN
    End If
    hFile = CreateFile(OpenFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
                       0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError HBF_OPEN_FAILURE
    End If
    sFName = OpenFileName
End Sub

Public Function ReadBytes(ByRef Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If ReadFile(hFile, _
                Buffer(LBound(Buffer)), _
                UBound(Buffer) - LBound(Buffer) + 1, _
                ReadBytes, _
                0) Then
        If ReadBytes = 0 Then
            fEOF = True
        End If
    Else
        RaiseError HBF_READ_FAILURE
    End If
End Function

Public Sub SeekAbsolute(ByVal Position As Currency)
    RaiseErrorIfClosed
    C.Value = Position / 10000@
    LSet L = C
    If SetFilePointer(hFile, L.LowVal, L.HighVal, FILE_BEGIN) _
        = INVALID_SET_FILE_POINTER Then
            If Err.LastDllError Then RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekEnd()
    RaiseErrorIfClosed
    If SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) _
        = INVALID_SET_FILE_POINTER Then
            RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
    'Offset is signed.
    RaiseErrorIfClosed
    If SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT) _
        = INVALID_SET_FILE_POINTER Then
            RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Function WriteBytes(Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If WriteFile(hFile, _
                 Buffer(LBound(Buffer)), _
                 UBound(Buffer) - LBound(Buffer) + 1, _
                 WriteBytes, _
                 0) Then
        If fAutoFlush Then Flush
    Else
        RaiseError HBF_WRITE_FAILURE
    End If
End Function

Private Sub Class_Initialize()
    hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

Private Sub RaiseError(ByVal ErrorCode As HBF_Errors)
    Dim Win32Err As Long, Win32Text As String

    Win32Err = Err.LastDllError
    If Win32Err Then
        Win32Text = vbNewLine & "Error " & Win32Err & vbNewLine _
                  & DecodeAPIErrors(Win32Err)
    End If
    If IsOpen Then CloseFile
    Select Case ErrorCode
        Case HBF_FILE_ALREADY_OPEN
            Err.Raise HBF_FILE_ALREADY_OPEN, HBF_SOURCE, _
                "File already open."
        Case HBF_OPEN_FAILURE
            Err.Raise HBF_OPEN_FAILURE, HBF_SOURCE, _
                "Error opening file." & Win32Text
        Case HBF_SEEK_FAILURE
            Err.Raise HBF_SEEK_FAILURE, HBF_SOURCE, _
                "Seek Error." & Win32Text
        Case HBF_FILELEN_FAILURE
            Err.Raise HBF_FILELEN_FAILURE, HBF_SOURCE, _
                "GetFileSize Error." & Win32Text
        Case HBF_READ_FAILURE
            Err.Raise HBF_READ_FAILURE, HBF_SOURCE, _
                "Read failure." & Win32Text
        Case HBF_WRITE_FAILURE
            Err.Raise HBF_WRITE_FAILURE, HBF_SOURCE, _
                "Write failure." & Win32Text
        Case HBF_FILE_ALREADY_CLOSED
            Err.Raise HBF_FILE_ALREADY_CLOSED, HBF_SOURCE, _
                "File must be open for this operation."
        Case Else
            Err.Raise HBF_UNKNOWN_ERROR, HBF_SOURCE, _
               "Unknown error." & Win32Text
    End Select
End Sub

Private Sub RaiseErrorIfClosed()
    If hFile = INVALID_HANDLE_VALUE Then RaiseError HBF_FILE_ALREADY_CLOSED
End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Dim strMsg As String, lngMsgLen As Long

    strMsg = Space$(256)
    lngMsgLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                              ErrorCode, 0&, strMsg, 256&, 0&)
    If lngMsgLen > 0 Then
        DecodeAPIErrors = Left(strMsg, lngMsgLen)
    Else
        DecodeAPIErrors = "Unknown Error."
    End If
End Function
