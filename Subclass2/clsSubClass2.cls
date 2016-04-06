VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)

Public ErrorMessage As String
Public ErrorNumber As Long

Function AttachMessage(hwnd As Variant, wMsg As Long) As Boolean
On Error GoTo hell

    Dim h As Long
    
    'allow them to just pass in an object by reference or its explicit hwnd
    If TypeName(hwnd) = "Long" Or TypeName(hwnd) = "Integer" Then
        h = hwnd
    Else
        h = hwnd.hwnd
    End If
    
    modSubclass.MonitorWindowMessage Me, h, wMsg
    AttachMessage = True
    ErrorMessage = ""
    ErrorNumber = 0

Exit Function
hell:
      ErrorMessage = Err.Description
      ErrorNumber = Err.Number

End Function

Sub DetatchMessage(hwnd As Variant, wMsg As Long)
    
    Dim h As Long
    
    'allow them to just pass in an object by reference or its explicit hwnd
    If TypeName(hwnd) = "Long" Or TypeName(hwnd) = "Integer" Then
        h = hwnd
    Else
        h = hwnd.hwnd
    End If
    
    modSubclass.DetachWindowMessage h, wMsg, Me
    
End Sub

Private Sub Class_Initialize()
    modSubclass.RegisterClassActive Me
End Sub

Private Sub Class_Terminate()
    modSubclass.RemoveActiveClass Me
End Sub

Friend Sub ForwardMessage(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    
    'this sub is only called from the module,
    'friend methods are not externally visible in compiled dll interface
    
    'these arguments are purposly byref so that you can change parameters
    RaiseEvent MessageReceived(hwnd, wMsg, wParam, lParam, Cancel)
    
    
End Sub

