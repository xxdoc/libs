Attribute VB_Name = "Sample"
Option Explicit

'**************************************************
'* NT Service sample                              *
'* © 2000-2004 Sergey Merzlikin                   *
'* http://www.smsoft.ru                           *
'* e-mail: sm@smsoft.ru                           *
'* The code is freeware. It may be used           *
'* in programs of any kind without permission     *
'**************************************************

Private Const Service_Name = "SampleVB6Service"
Public Const INFINITE = -1&      '  Infinite timeout
Private Const WAIT_TIMEOUT = 258&

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 128) As Byte
End Type

Public Const VER_PLATFORM_WIN32_NT = 2&
Private Const STATUS_TIMEOUT = &H102&
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT _
        Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON _
        Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Declare Function MsgWaitForMultipleObjects Lib "user32" _
        (ByVal nCount As Long, pHandles As Long, _
        ByVal fWaitAll As Long, ByVal dwMilliseconds _
        As Long, ByVal dwWakeMask As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public hStopEvent As Long, hStartEvent As Long, hStopPendingEvent As Long
Public IsNT As Boolean, IsNTService As Boolean
Public ServiceNamePtr As Long

Private Sub Main()
    Dim hnd As Long
    Dim h(0 To 1) As Long
    ' Only one instance
    If App.PrevInstance Then Exit Sub
    ' Check OS type
    IsNT = CheckIsNT()
    ' Creating events
    hStopEvent = CreateEventW(0&, 1&, 0&, 0&)
    hStopPendingEvent = CreateEventW(0&, 1&, 0&, 0&)
    hStartEvent = CreateEventW(0&, 1&, 0&, 0&)
    ServiceNamePtr = StrPtr(Service_Name)
    If IsNT Then
        ' Trying to start service
        hnd = StartAsService
        h(0) = hnd
        h(1) = hStartEvent
        ' Waiting for one of two events: sucsessful service start (1) or
        ' terminaton of service thread (0)
        IsNTService = MsgWaitObj(INFINITE, h(0), 2&) = 1&
        If Not IsNTService Then
            CloseHandle hnd
            MessageBox 0&, "This program must be started as a service.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
        End If
    Else
        MessageBox 0&, "This program is only for Windows NT/2000/XP/2003.", App.Title, vbInformation Or vbOKOnly Or vbMsgBoxSetForeground
    End If
    
    If IsNTService Then
        ' ******************
        ' Here you may initialize and start service's objects
        ' These objects must be event-driven and must return control
        ' immediately after starting.
        ' ******************
        SetServiceState SERVICE_RUNNING
        App.LogEvent "VB6 Service Sample started"
        Do
            ' ******************
            ' It is main service loop. Here you may place statements
            ' which perform useful functionality of this service.
            ' ******************
            ' Loop repeats every second. You may change this interval.
        Loop While MsgWaitObj(1000&, hStopPendingEvent, 1&) = WAIT_TIMEOUT
        ' ******************
        ' Here you may stop and destroy service's objects
        ' ******************
        SetServiceState SERVICE_STOPPED
        App.LogEvent "VB6 Service Sample stopped"
        SetEvent hStopEvent
        ' Waiting for service thread termination
        MsgWaitObj INFINITE, hnd, 1&
        CloseHandle hnd
    End If
    CloseHandle hStopEvent
    CloseHandle hStartEvent
    CloseHandle hStopPendingEvent
End Sub

' CheckIsNT() returns True, if the program runs
' under Windows NT or Windows 2000, and False
' otherwise.
Public Function CheckIsNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    OSVer.dwOSVersionInfoSize = LenB(OSVer)
    GetVersionEx OSVer
    CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT
End Function

' The MsgWaitObj function replaces Sleep,
' WaitForSingleObject, WaitForMultipleObjects functions.
' Unlike these functions, it
' doesn't block thread messages processing.
' Using instead Sleep:
'     MsgWaitObj dwMilliseconds
' Using instead WaitForSingleObject:
'     retval = MsgWaitObj(dwMilliseconds, hObj, 1&)
' Using instead WaitForMultipleObjects:
'     retval = MsgWaitObj(dwMilliseconds, hObj(0&), n),
'     where n - wait objects quantity,
'     hObj() - their handles array.

Public Function MsgWaitObj(Interval As Long, _
            Optional hObj As Long = 0&, _
            Optional nObj As Long = 0&) As Long
    Dim T As Long, T1 As Long
    If Interval <> INFINITE Then
        T = GetTickCount()
        On Error Resume Next
        T = T + Interval
        ' Overflow prevention
        If Err <> 0& Then
            If T > 0& Then
                T = ((T + &H80000000) _
                + Interval) + &H80000000
            Else
                T = ((T - &H80000000) _
                + Interval) - &H80000000
            End If
        End If
        On Error GoTo 0
        ' T contains now absolute time of the end of interval
    Else
        T1 = INFINITE
    End If
    Do
        If Interval <> INFINITE Then
            T1 = GetTickCount()
            On Error Resume Next
         T1 = T - T1
            ' Overflow prevention
            If Err <> 0& Then
                If T > 0& Then
                    T1 = ((T + &H80000000) _
                    - (T1 - &H80000000))
                Else
                    T1 = ((T - &H80000000) _
                    - (T1 + &H80000000))
                End If
            End If
            On Error GoTo 0
            ' T1 contains now the remaining interval part
            If IIf((T1 Xor Interval) > 0&, _
                T1 > Interval, T1 < 0&) Then
                ' Interval expired
                ' during DoEvents
                MsgWaitObj = STATUS_TIMEOUT
                Exit Function
            End If
        End If
        ' Wait for event, interval expiration
        ' or message appearance in thread queue
        MsgWaitObj = MsgWaitForMultipleObjects(nObj, _
                hObj, 0&, T1, QS_ALLINPUT)
        ' Let's message be processed
        DoEvents
        If MsgWaitObj <> nObj Then Exit Function
        ' It was message - continue to wait
    Loop
End Function


