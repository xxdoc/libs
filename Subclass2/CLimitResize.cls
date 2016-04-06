VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLimitResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents subclass As CSubclass2
Attribute subclass.VB_VarHelpID = -1

'key= "hwnd:" & frm.hwnd , data = configured minmaxinfo as bytearray
Private cForms As New Collection

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Const WM_GETMINMAXINFO = &H24
Private Type POINTAPI
     x As Long
     y As Long
End Type
Private Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Sub Add(frm As Object, minWidth As Long, minHeight As Long, Optional maxWidth = 10000, Optional maxHeight = 10000)
     
    On Error Resume Next
    Dim b() As Byte
    Dim tmp As MINMAXINFO
    Dim alreadyExists  As Boolean
    Dim key As String
    
    key = "hwnd:" & frm.hwnd
    
    ReDim b(Len(tmp))
    b = cForms(key)
    If Err.Number <> 0 Then alreadyExists = False
    
    tmp.ptMinTrackSize.x = minWidth
    tmp.ptMinTrackSize.y = minHeight
    tmp.ptMaxTrackSize.x = maxWidth
    tmp.ptMaxTrackSize.y = maxHeight
    
    CopyMemory b(0), tmp, Len(tmp)
    
    If alreadyExists Then
        cForms.Remove key
        cForms.Add b, key
    Else
        cForms.Add b, key
        subclass.AttachMessage frm.hwnd, WM_GETMINMAXINFO
    End If
        
End Sub

Sub Remove(frm As Object)
    Dim key As String
    
    key = "hwnd:" & frm.hwnd
    
    On Error GoTo hell
    cForms.Remove key
    subclass.DetatchMessage frm.hwnd, WM_GETMINMAXINFO
    
hell:
End Sub

Private Sub Class_Initialize()
    Set subclass = New CSubclass2
End Sub

Private Sub Class_Terminate()
    'subclass_terminate will release all subclasses
    Set subclass = Nothing
End Sub

Private Sub subclass_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
        Dim minMax As MINMAXINFO
        Dim b() As Byte
        Dim key As String
        
        ReDim b(Len(minMax))
        
        On Error Resume Next
        key = "hwnd:" & hwnd
        b() = cForms(key)
        'If Err.Number <> 0 Then subclass.DetatchMessage hwnd, WM_GETMINMAXINFO
        If Err.Number = 0 Then
                CopyMemory ByVal lParam, b(0), Len(minMax)
                Cancel = True
        End If
        
End Sub
