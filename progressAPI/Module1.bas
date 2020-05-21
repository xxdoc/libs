Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const ICC_PROGRESS_CLASS As Long = &H20 ' Load progress bar control class. https://msdn.microsoft.com/en-us/library/windows/desktop/bb775507(v=vs.85).aspx

Public qwProgressBar As Long

Public Const WS_VISIBLE As Long = &H10000000 ' Creates a window that is initially visible. This applies to overlapped, child, and pop-up windows. For overlapped windows, the y parameter is used as a ShowWindow function parameter. https://support.microsoft.com/en-us/kb/111011
Public Const WS_CHILD As Long = &H40000000 ' The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style. https://support.microsoft.com/en-us/kb/111011
Public Const CCM_FIRST = &H2000
Public Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Public Const WM_USER As Long = &H400 ' Microsoft: Win32api.txt

Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR ' Sets the Progress Bar background color
Public Const PBM_SETRANGE = (WM_USER + 1) ' Sets the total low & high of the Progress Bar
Public Const PBM_SETPOS = (WM_USER + 2) ' Sets the current % position of the Progress Bar
Public Const PBM_DELTAPOS = (WM_USER + 3)
Public Const PBM_SETSTEP = (WM_USER + 4)
Public Const PBM_STEPIT = (WM_USER + 5)
Public Const PBM_SETRANGE32 = (WM_USER + 6)
Public Const PBM_GETRANGE = (WM_USER + 7)
Public Const PBM_GETPOS = (WM_USER + 8)
Public Const PBM_SETBARCOLOR = (WM_USER + 9) ' Sets the color of the Progress Bar
Public Const PBM_SETMARQUEE = (WM_USER + 10)
Public Const PBM_GETSTEP = (WM_USER + 13)
Public Const PBM_GETBKCOLOR = (WM_USER + 14) ' Get Progress Bar Background Color
Public Const PBM_GETBARCOLOR = (WM_USER + 15) ' Get Progress Bar bar Color
Public Const PBM_SETSTATE = (WM_USER + 16) ' NORMAL, ERROR, PAUSED
Public Const PBM_GETSTATE = (WM_USER + 17)
Public Const PBS_SMOOTH = &H1 ' Smooth looking Progress Bar
Public Const PBS_VERTICAL = &H4 ' Vertical Progress Bar
Public Const PROGRESS_CLASS = "msctls_progress32"

Public Type tagINITCOMMONCONTROLSEX
   dwSize As Long
   dwICC As Long
End Type

Public Function InitComctl32(dwFlags As Long) As Boolean

   Dim icc As tagINITCOMMONCONTROLSEX
   
   On Error GoTo Err_OldVersion
  
   With icc
      .dwSize = Len(icc)
      .dwICC = dwFlags
   End With
     
   InitComctl32 = InitCommonControlsEx(icc)
   Exit Function

Err_OldVersion:
   InitCommonControls
   
End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  'Combines two integers into a long
   MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
   MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Public Function LoWord(dwValue As Long) As Integer
   CopyMemory LoWord, dwValue, 2
End Function
