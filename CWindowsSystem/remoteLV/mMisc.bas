Attribute VB_Name = "Module5"
'========mMisc.bas=======
'Some support functions
Option Explicit

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal nStyle As Long)
Public Const GWL_STYLE = (-16)

Public Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function
