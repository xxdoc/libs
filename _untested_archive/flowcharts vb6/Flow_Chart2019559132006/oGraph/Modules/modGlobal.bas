Attribute VB_Name = "modGlobal"
Option Explicit


Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Public Const WINDING = 2 ' constants for FillMode.
Public Const RGN_AND As Long = 1
Public Const RGN_COPY As Long = 5
Public Const RGN_DIFF As Long = 4
Public Const RGN_ERROR As Long = 0
Public Const RGN_MAX As Long = RGN_COPY
Public Const RGN_MIN As Long = RGN_AND
Public Const RGN_OR As Long = 2
Public Const RGN_XOR As Long = 3

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046


Public Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'----------------------------------------------------------------------
' Type Defs.
'----------------------------------------------------------------------
'

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

