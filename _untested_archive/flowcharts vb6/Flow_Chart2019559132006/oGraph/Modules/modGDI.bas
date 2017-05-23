Attribute VB_Name = "modGDI"
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
      ByVal hObject As Long, _
      ByVal nCount As Long, _
      lpObject As Any _
   ) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
      ByVal lpDriverName As String, _
      lpDeviceName As Any, _
      lpOutput As Any, _
      lpInitData As Any _
    ) As Long
Private Declare Function DeleteDC Lib "gdi32" ( _
       ByVal hdc As Long _
    ) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
       ByVal hdc As Long, ByVal hObj As Long _
    ) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
       ByVal hObj As Long _
    ) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
       ByVal hdc As Long, _
       ByVal nWidth As Long, _
       ByVal nHeight As Long _
    ) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
       ByVal hdc As Long _
    ) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
       ByVal hDestDC As Long, _
       ByVal X As Long, ByVal Y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, _
       ByVal hSrcDC As Long, _
       ByVal xSrc As Long, ByVal ySrc As Long, _
       ByVal dwRop As Long) As Long



Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal xOffset As Long, ByVal yOffset As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


Private Enum EDrawTextFormat
   DT_BOTTOM = &H8
   DT_CALCRECT = &H400
   DT_CENTER = &H1
   DT_EXPANDTABS = &H40
   DT_EXTERNALLEADING = &H200
   DT_INTERNAL = &H1000
   DT_LEFT = &H0
   DT_NOCLIP = &H100
   DT_NOPREFIX = &H800
   DT_RIGHT = &H2
   DT_SINGLELINE = &H20
   DT_TABSTOP = &H80
   DT_TOP = &H0
   DT_VCENTER = &H4
   DT_WORDBREAK = &H10
   DT_EDITCONTROL = &H2000&
   DT_PATH_ELLIPSIS = &H4000&
   DT_END_ELLIPSIS = &H8000&
   DT_MODIFYSTRING = &H10000
   DT_RTLREADING = &H20000
   DT_WORD_ELLIPSIS = &H40000
End Enum
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type DrawObjectData
   iDirX As Integer
   iDirY As Integer
   ForeColor As Long
   BackColor As Long
   sCaption As String
   tP As POINTAPI
   lSize As Long
End Type


Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PictDesc, riid As GUID, _
    ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Function HBitmapFromPicture(picThis As IPicture) As Long

   ' Create a copy of the bitmap:
   Dim lhDC As Long
   Dim lhDCCopy As Long
   Dim lhBmpCopy As Long
   Dim lhBmpCopyOld As Long
   Dim lhBmpOld As Long
   Dim lhDCC As Long
   Dim tBM As BITMAP

   GetObjectAPI picThis.handle, Len(tBM), tBM
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDC = CreateCompatibleDC(lhDCC)
   lhBmpOld = SelectObject(lhDC, picThis.handle)

   lhDCCopy = CreateCompatibleDC(lhDCC)
   lhBmpCopy = CreateCompatibleBitmap(lhDCC, tBM.bmWidth, tBM.bmHeight)
   lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)

   BitBlt lhDCCopy, 0, 0, tBM.bmWidth, tBM.bmHeight, lhDC, 0, 0, vbSrcCopy

   If Not (lhDCC = 0) Then
      DeleteDC lhDCC
   End If
   If Not (lhBmpOld = 0) Then
      SelectObject lhDC, lhBmpOld
   End If
   If Not (lhDC = 0) Then
      DeleteDC lhDC
   End If
   If Not (lhBmpCopyOld = 0) Then
      SelectObject lhDCCopy, lhBmpCopyOld
   End If
   If Not (lhDCCopy = 0) Then
      DeleteDC lhDCCopy
   End If

   HBitmapFromPicture = lhBmpCopy

End Function

Public Function HBitmapFromDC( _
      ByVal lhDC As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   ) As Long

   ' Copy the bitmap in lHDC:
   Dim lhDCCopy As Long
   Dim lhBmpCopy As Long
   Dim lhBmpCopyOld As Long
   Dim lhDCC As Long
   Dim tBM As BITMAP
   
   lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDCCopy = CreateCompatibleDC(lhDCC)
   lhBmpCopy = CreateCompatibleBitmap(lhDCC, lWidth, lHeight)
   lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)

   BitBlt lhDCCopy, 0, 0, lWidth, lHeight, lhDC, 0, 0, vbSrcCopy

   If Not (lhDCC = 0) Then
      DeleteDC lhDCC
   End If
   If Not (lhBmpCopyOld = 0) Then
      SelectObject lhDCCopy, lhBmpCopyOld
   End If
   If Not (lhDCCopy = 0) Then
      DeleteDC lhDCCopy
   End If

   HBitmapFromDC = lhBmpCopy

End Function
  
  
Public Function IconToTPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As GUID
    
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .picType = vbPicTypeIcon
       .hImage = hIcon
    End With
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    
    Set IconToTPicture = oNewPic
    
End Function



Public Sub drawBackground(ByVal lhDC As Long, ByVal pBackColour As Long, tDrawR As RECT)
    Dim hBr As Long
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJunk As POINTAPI
   
   
   hBr = GetSysColorBrush(pBackColour) 'And &H1F&)
   FillRect lhDC, tDrawR, hBr
   Call SetBkColor(lhDC, pBackColour)
   Call SetBkMode(lhDC, TRANSPARENT)
   DeleteObject hBr
   'hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbHighlight And &H1F&))
   'hPenOld = SelectObject(lhDC, hPen)
   'MoveToEx lhDC, tDrawR.Left, tDrawR.Top, tJunk
   'LineTo lhDC, tDrawR.Right - 1, tDrawR.Top
   'LineTo lhDC, tDrawR.Right - 1, tDrawR.Bottom - 1
   'LineTo lhDC, tDrawR.Left, tDrawR.Bottom - 1
   'LineTo lhDC, tDrawR.Left, tDrawR.Top
   'SelectObject lhDC, hPenOld
   'DeleteObject hPen
   
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Private Sub GetRBGFromOLEColour(ByVal dwOleColour As Long, r As Long, g As Long, b As Long)
    
  'pass a hex colour, return the rgb components
   Dim clrref As Long
   
  'translate OLE color to valid color if passed
   OleTranslateColor dwOleColour, 0, clrref
  
   b = (clrref \ 65536) And &HFF
   g = (clrref \ 256) And &HFF
   r = clrref And &HFF
   
End Sub

Public Function GetGDIColorFromOLE(ByVal OleColor As OLE_COLOR, Optional ByVal AlphaValue As Byte = 255) As Long
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim bytestruct       As COLORBYTES
    Dim result           As COLORLONG
    
    Call GetRBGFromOLEColour(OleColor, r, g, b)
    bytestruct.RedByte = r
    bytestruct.GreenByte = g
    bytestruct.BlueByte = b
    bytestruct.AlphaByte = AlphaValue
    LSet result = bytestruct
    GetGDIColorFromOLE = result.longval
End Function

 Public Sub HLSToRGB( _
     ByVal h As Single, ByVal s As Single, ByVal l As Single, _
     r As Long, g As Long, b As Long _
     )
 Dim rR As Single, rG As Single, rB As Single
 Dim Min As Single, Max As Single

     If s = 0 Then
     ' Achromatic case:
     rR = l: rG = l: rB = l
     Else
     ' Chromatic case:
     ' delta = Max-Min
     If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
     Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
     End If
     ' Get the Max value:
     Max = 2 * l - Min
     
     ' Now depending on sector we can evaluate the h,l,s:
     If (h < 1) Then
         rR = Max
         If (h < 0) Then
             rG = Min
             rB = rG - h * (Max - Min)
         Else
             rB = Min
             rG = h * (Max - Min) + rB
         End If
     ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
             rB = Min
             rR = rB - (h - 2) * (Max - Min)
         Else
             rR = Min
             rB = (h - 2) * (Max - Min) + rR
         End If
     Else
         rB = Max
         If (h < 4) Then
             rR = Min
             rG = rR - (h - 4) * (Max - Min)
         Else
             rG = Min
             rR = (h - 4) * (Max - Min) + rG
         End If
         
     End If
             
     End If
     r = rR * 255: g = rG * 255: b = rB * 255
 End Sub


Public Sub RGBToHLS( _
     ByVal r As Long, ByVal g As Long, ByVal b As Long, _
     h As Single, s As Single, l As Single _
     )
 Dim Max As Single
 Dim Min As Single
 Dim delta As Single
 Dim rR As Single, rG As Single, rB As Single

     rR = r / 255: rG = g / 255: rB = b / 255

 '{Given: rgb each in [0,1].
 ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
         Max = Maximum(rR, rG, rB)
         Min = Minimum(rR, rG, rB)
             l = (Max + Min) / 2 '{This is the lightness}
         '{Next calculate saturation}
         If Max = Min Then
             'begin {Acrhomatic case}
             s = 0
             h = 0
             'end {Acrhomatic case}
         Else
             'begin {Chromatic case}
                 '{First calculate the saturation.}
             If l <= 0.5 Then
                 s = (Max - Min) / (Max + Min)
             Else
                 s = (Max - Min) / (2 - Max - Min)
             End If
             '{Next calculate the hue.}
             delta = Max - Min
             If rR = Max Then
                     h = (rG - rB) / delta '{Resulting color is between yellow and magenta}
             ElseIf rG = Max Then
                 h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
             ElseIf rB = Max Then
                 h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
             End If
         'end {Chromatic Case}
     End If
 End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     If (rR > rG) Then
     If (rR > rB) Then
         Maximum = rR
     Else
         Maximum = rB
     End If
     Else
     If (rB > rG) Then
         Maximum = rB
     Else
         Maximum = rG
     End If
     End If
 End Function
 Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     If (rR < rG) Then
     If (rR < rB) Then
         Minimum = rR
     Else
         Minimum = rB
     End If
     Else
     If (rB < rG) Then
         Minimum = rB
     Else
         Minimum = rG
     End If
 End If
 End Function
