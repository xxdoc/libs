VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Mutex Example
'http://sandsprite.com
'dzzie@yahoo.com

Private Declare Function CreateMutex Lib "Kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "Kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const WAIT_FAILED = -1&
Private Const WAIT_TIMEOUT = &H102&

Dim hMutex As Long

Private Sub Class_Terminate()
     Release
End Sub
 
Function Release()
    ReleaseMutex hMutex
    CloseHandle hMutex
    hMutex = 0
End Function

Function Create(name As String, Optional owned As Boolean = True) As Boolean
    
    'If hMutex <> 0 Then
    '    MsgBox "Only Use one Mutex at a time perClass", vbExclamation
    '    Exit Function
    'End If
    
    Dim m As Long
    If owned Then m = 1 Else m = 0
    
    hMutex = CreateMutex(0&, m, name)
    Create = True
    If Err.LastDllError = 183& Then Create = False
    
End Function
 
Function Wait(Optional pollInterval As Long = 100, Optional maxCycles As Long = -1) As Boolean
    
    Dim lret As Long
    Dim Unlocked As Boolean
    Dim cycles As Long
    
    Do Until Unlocked = True
           
           lret = WaitForSingleObject(hMutex, pollInterval)   'request ownership
           DoEvents
           
           If maxCycles > 0 Then
                cycles = cycles + 1
                If cycles >= maxCycles Then Exit Function 'returns false
           End If
           
           If (lret <> WAIT_TIMEOUT) And (lret <> WAIT_FAILED) Then Unlocked = True
           
    Loop
    
    Wait = True
    
End Function



 


