VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCmdOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'NOTE: to use the async data out event the console app you are reading from will have to
'either manually call flush or setvbuf(stdout, NULL, _IONBF, 0) once at beginning of app..
Event OutPutAvailable(Data As String)

Public result As String
Public exitCode As Long
Public hideWindow As Boolean
Public blnStdOut As Boolean
Public blnStdErr As Boolean
Public blnOEMConvert As Boolean
Public maxExecSecs As Long

Private m_abort As Boolean
Private m_terminateOnAbort As Boolean

''''''''''''''''''''''''''''''''''''''''
' Joacim Andersson, Brixoft Software
' http://www.brixoft.net
''''''''''''''''''''''''''''''''''''''''

' STARTUPINFO flags
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100

' ShowWindow flags
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const SW_SHOWNORMAL = 1

' DuplicateHandle flags
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

' Error codes
Private Const ERROR_BROKEN_PIPE = 109

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CreatePipe _
 Lib "kernel32" ( _
 phReadPipe As Long, _
 phWritePipe As Long, _
 lpPipeAttributes As Any, _
 ByVal nSize As Long) As Long

Private Declare Function ReadFile _
 Lib "kernel32" ( _
 ByVal hFile As Long, _
 lpBuffer As Any, _
 ByVal nNumberOfBytesToRead As Long, _
 lpNumberOfBytesRead As Long, _
 lpOverlapped As Any) As Long

Private Declare Function CreateProcess _
 Lib "kernel32" Alias "CreateProcessA" ( _
 ByVal lpApplicationName As String, _
 ByVal lpCommandLine As String, _
 lpProcessAttributes As Any, _
 lpThreadAttributes As Any, _
 ByVal bInheritHandles As Long, _
 ByVal dwCreationFlags As Long, _
 lpEnvironment As Any, _
 ByVal lpCurrentDriectory As String, _
 lpStartupInfo As STARTUPINFO, _
 lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetCurrentProcess _
 Lib "kernel32" () As Long

Private Declare Function DuplicateHandle _
 Lib "kernel32" ( _
 ByVal hSourceProcessHandle As Long, _
 ByVal hSourceHandle As Long, _
 ByVal hTargetProcessHandle As Long, _
 lpTargetHandle As Long, _
 ByVal dwDesiredAccess As Long, _
 ByVal bInheritHandle As Long, _
 ByVal dwOptions As Long) As Long

Private Declare Function CloseHandle _
 Lib "kernel32" ( _
 ByVal hObject As Long) As Long

Private Declare Function OemToCharBuff _
 Lib "user32" Alias "OemToCharBuffA" ( _
 lpszSrc As Any, _
 ByVal lpszDst As String, _
 ByVal cchDstLength As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long

Private Const WAIT_TIMEOUT = &H102

Sub abort(Optional terminate As Boolean = True)
    m_terminateOnAbort = terminate
    m_abort = True
End Sub

Private Sub Class_Initialize()
    hideWindow = True
    blnStdOut = True
    blnStdErr = True
    blnOEMConvert = True
    m_terminateOnAbort = True
End Sub

Sub CfgHandles(Optional stdOut As Boolean = True, Optional stdErr As Boolean = True, Optional oemConvert As Boolean = True)
    blnStdOut = stdOut
    blnStdErr = stdErr
    blnOEMConvert = oemConvert
End Sub

Sub CfgOpts(Optional hideWin As Boolean = True, Optional maxExecTimeSecs As Long = 0)
    hideWindow = hideWin
    maxExecSecs = maxExecTimeSecs
End Sub


'you can specify executable or command line or both.
'if the executable path has any spaces in it you must specify exe path
'if you do specify the executable, the command line is now only the args
'and argv[0] must be manually specified to this api, normally it is exe path, and normal programs expect their args to start with argv[1]
'because of this, if exe is not null, and command line is not null we automatically prefix the command line with the exe path as arg0 to
'maintain expected behavior..

Public Function GetCommandOutput( _
    Optional executable As String = vbNullString, _
    Optional sCommandLine As String = vbNullString _
) As Boolean

    Dim hPipeRead As Long, hPipeWrite1 As Long, hPipeWrite2 As Long
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim baOutput() As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim fTwoHandles As Boolean
    Dim success As Boolean
    Dim lRet As Long
    Dim startTime As Date
    
    Const BUFSIZE = 1024      ' pipe buffer size
    exitCode = 0
    result = Empty
    startTime = Now
    m_abort = False
    
    'see comment above..
    If Len(executable) > 0 And Len(sCommandLine) > 0 Then
        sCommandLine = """" & executable & """ " & sCommandLine
    End If
    
    ' At least one of them should be True, otherwise there's no point in calling the function
    If (Not blnStdOut) And (Not blnStdErr) Then
        Err.Raise 5         ' Invalid Procedure call or Argument
        'GetCommandOutput = LaunchProcess(executable, sCommandLine, True, hideWindow, exitCode)
        'Exit Function
    End If

    ' If both are true, we need two write handles. If not, one is enough.
    fTwoHandles = blnStdOut And blnStdErr

    ReDim baOutput(BUFSIZE - 1) As Byte

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With

    If CreatePipe(hPipeRead, hPipeWrite1, sa, BUFSIZE) = 0 Then
        Exit Function
    End If

    hCurProcess = GetCurrentProcess()

    ' Replace our inheritable read handle with an non-inheritable. Not that it
    ' seems to be necessary in this case, but the docs say we should.
    Call DuplicateHandle(hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, 0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)

    ' If both STDOUT and STDERR should be redirected, get an extra handle.
    If fTwoHandles Then
        Call DuplicateHandle(hCurProcess, hPipeWrite1, hCurProcess, hPipeWrite2, 0&, 1&, DUPLICATE_SAME_ACCESS)
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = IIf(hideWindow, SW_HIDE, SW_SHOW)          ' hide the window

        If fTwoHandles Then
            .hStdOutput = hPipeWrite1
            .hStdError = hPipeWrite2
        ElseIf blnStdOut Then
            .hStdOutput = hPipeWrite1
        Else
            .hStdError = hPipeWrite1
        End If
    End With

    If CreateProcess(executable, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, pi) Then

        ' Close thread handle - we don't need it
        Call CloseHandle(pi.hThread)

        ' Also close our handle(s) to the write end of the pipe. This is important, since
        ' ReadFile will *not* return until all write handles are closed or the buffer is full.
        Call CloseHandle(hPipeWrite1)
        hPipeWrite1 = 0
        If hPipeWrite2 Then
            Call CloseHandle(hPipeWrite2)
            hPipeWrite2 = 0
        End If

        Dim i As Long
        Do
            ' Add a DoEvents to allow more data to be written to the buffer for each call.
            ' This results in fewer, larger chunks to be read.
            'DoEvents
            
            'on win7 x64, if the process has exited, a readFile of hPipeRead will cause complete hang..
            'so we must check is process is still active..
            If WaitForSingleObject(pi.hProcess, 0) <> WAIT_TIMEOUT Then 'process has terminated..
                Exit Do
            End If
            
            Dim avail As Long, bleft As Long
            If PeekNamedPipe(hPipeRead, ByVal 0&, 0, lBytesRead, avail, bleft) <> 0 Then
                If avail > 0 Then
                    If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                        Exit Do
                    End If
                End If
            End If

            If blnOEMConvert Then
                ' convert from "DOS" to "Windows" characters
                sNewOutput = String$(lBytesRead, 0)
                Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
            Else
                ' perform no conversion (except to Unicode)
                sNewOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
            End If

            result = result & sNewOutput

            ' If you are executing an application that outputs data during a long time,
            ' and don't want to lock up your application, it might be a better idea to
            ' wrap this code in a class module in an ActiveX EXE and execute it asynchronously.
            ' Then you can raise an event here each time more data is available.
            RaiseEvent OutPutAvailable(sNewOutput)
            
            i = i + 1
            If i = 50 Then
                DoEvents
                i = 0
            End If
            
            If i Mod 3 = 0 Then
                If maxExecSecs > 0 Then
                    If DateDiff("s", startTime, Now) >= maxExecSecs Then
                        If TerminateProcess(pi.hProcess, -22) <> 0 Then Exit Do
                    End If
                End If
            End If
            
            If m_abort Then
                If m_terminateOnAbort Then TerminateProcess pi.hProcess, -22
                Exit Do
            End If
        Loop

        ' When the process terminates successfully, Err.LastDllError will be
        ' ERROR_BROKEN_PIPE (109). Other values indicates an error.

        If GetExitCodeProcess(pi.hProcess, exitCode) = 0 Then exitCode = -1
        Call CloseHandle(pi.hProcess)
        success = True
    Else
        result = "Failed to create process, check the path of the command line."
    End If

    ' clean up
    Call CloseHandle(hPipeRead)
    
    If hPipeWrite1 Then Call CloseHandle(hPipeWrite1)
    If hPipeWrite2 Then Call CloseHandle(hPipeWrite2)
    
    GetCommandOutput = success
    
End Function



Public Function LaunchProcess( _
        Optional ByVal executable As String = vbNullString, _
        Optional ByVal sCommandLine As String = vbNullString, _
        Optional wait As Boolean = True _
) As Boolean

    Dim hPipeRead As Long, hPipeWrite1 As Long, hPipeWrite2 As Long
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim lRet As Long
    Dim startTime As Date
    Dim i As Long
    
    Const STILL_ACTIVE = 259
 
    Const BUFSIZE = 1024      ' pipe buffer size
    
    m_abort = False
    exitCode = 0
    result = Empty
    startTime = Now
    
    'see comment above..
    If Len(executable) > 0 And Len(sCommandLine) > 0 Then
        sCommandLine = """" & executable & """ " & sCommandLine
    End If
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 0    ' get inheritable pipe handles
    End With

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW 'Or STARTF_USESTDHANDLES
        .wShowWindow = IIf(hideWindow, SW_HIDE, SW_SHOW)
    End With

    If CreateProcess(executable, sCommandLine, ByVal 0&, ByVal 0&, 0, 0&, ByVal 0&, vbNullString, si, pi) Then

        ' Close thread handle - we don't need it
        Call CloseHandle(pi.hThread)

        If wait Then
            Do
                i = i + 1
                WaitForSingleObject pi.hProcess, 100
                If GetExitCodeProcess(pi.hProcess, exitCode) = 0 Then Exit Do
                If exitCode <> STILL_ACTIVE Then Exit Do
                
                If i Mod 3 = 0 Then
                    If maxExecSecs > 0 Then
                        If DateDiff("s", startTime, Now) >= maxExecSecs Then
                            If TerminateProcess(pi.hProcess, -22) <> 0 Then Exit Do
                        End If
                    End If
                End If
                
                Sleep 10
                DoEvents
                
                If m_abort Then
                    If m_terminateOnAbort Then TerminateProcess pi.hProcess, -22
                    Exit Do
                End If
            Loop
        End If

         If GetExitCodeProcess(pi.hProcess, exitCode) = 0 Then exitCode = -1
        Call CloseHandle(pi.hProcess)
        LaunchProcess = True
    
    End If

    
    
End Function




