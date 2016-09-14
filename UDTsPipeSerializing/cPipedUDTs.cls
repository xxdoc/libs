VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPipedUDTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Pipe-based helper to serialize/deserialize VB-UDTs InMemory ... [Olaf Schmidt 2015]

Private Declare Function CreateNamedPipeW& Lib "kernel32" (ByVal lpName As Long, ByVal dwOpenMode&, ByVal dwPipeMode&, ByVal nMaxInstances&, ByVal nOutBufferSize&, ByVal nInBufferSize&, ByVal nDefaultTimeOut&, ByVal lpSecurityAttributes&)
Private Declare Function WriteFile& Lib "kernel32" (ByVal hFile&, lpBuffer As Any, ByVal nNumberOfBytesToWrite&, lpNumberOfBytesWritten&, ByVal lpOverlapped&)
Private Declare Function ReadFile& Lib "kernel32" (ByVal hFile&, lpBuffer As Any, ByVal nNumberOfBytesToRead&, lpNumberOfBytesRead&, ByVal lpOverlapped&)
Private Declare Function PeekNamedPipe& Lib "kernel32" (ByVal hNamedPipe&, lpBuffer As Any, ByVal nBufferSize&, lpBytesRead&, lpTotalBytesAvail&, lpBytesLeftThisMessage&)
Private Declare Function DisconnectNamedPipe& Lib "kernel32" (ByVal hPipe&)
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject&)
 
Private hPipe As Long, FNr As Long, PipeName As String

Property Get handle() As Long
    handle = FNr
End Property

Public Function Init(Optional PipeNameSuffix$ = "vbPipedUDTs") As Boolean
Const PipeNamePrefix$ = "\\.\pipe\"
  CleanUp
  PipeName = PipeNamePrefix & PipeNameSuffix
  hPipe = CreateNamedPipeW(StrPtr(PipeName), 3, 0, 255, -1, -1, 0, 0)
  If hPipe = -1 Then hPipe = 0 'reset from InvalidHandleValue to "no Handle"
  If hPipe Then FNr = FreeFile 'only try to find a free VB-FileNumber when hPipe is valid
  If FNr Then Open PipeName For Binary As FNr 'open only, when we got an FNr
  Init = hPipe <> 0 And FNr <> 0
End Function

Public Function ReadBytes() As Byte()
    Dim Bytes As Long
    Dim b() As Byte
    If FNr = 0 Then Exit Function
    PeekNamedPipe hPipe, ByVal 0&, 0, ByVal 0&, Bytes, 0 'peek if there's bytes available at all
    If Bytes <= 0 Then Exit Function
    ReDim Preserve b(0 To Bytes - 1)
    ReadFile hPipe, b(0), Bytes, Bytes, ByVal 0&
    ReadBytes = b()
End Function

Public Function WriteBytes(b() As Byte)
Dim Bytes As Long
    If FNr = 0 Then Exit Function
    WriteFile hPipe, b(0), UBound(b) + 1, Bytes, 0
    If Bytes <> UBound(b) + 1 Then Exit Function
    WriteBytes = Loc(FNr) 'report the amount of deserialized Bytes
End Function

Private Sub CleanUp()
  If FNr Then Close FNr: FNr = 0
  If hPipe Then DisconnectNamedPipe hPipe
  If hPipe Then CloseHandle hPipe: hPipe = 0
End Sub
 
Private Sub Class_Terminate()
  CleanUp
End Sub

