Attribute VB_Name = "modPrint"
Const PT_LINETO = &H2
Const PT_BEZIERTO = &H4
Const PT_CLOSEFIGURE = &H1
Const DI_APPBANDING = &H1
Const DI_ROPS_READ_DESTINATION = &H2
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type DOCINFO
cbSize As Long
lpszDocName As String
lpszOutput As String
lpszDatatype As String
fwType As Long
End Type
Private Declare Function StretchBlt Lib "GDI32.DLL" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long

Public Sub PrintMemoryToPrinter(ByVal pCWidth As Single, ByVal pCHeight As Single, ByVal Prn As Printer, ByRef pMemDc As cDibSection, Optional ByVal pDocName As String)
        Dim DI As DOCINFO
        Dim vbHiMetric As Integer
        Dim PicRatio As Double
        Dim PrnWidth As Double
        Dim PrnHeight As Double
        Dim PrnRatio As Double
        Dim PrnPicWidth As Double
        Dim PrnPicHeight As Double
        Dim pHeight As Single
        Dim pWidth As Single
        Dim pWidthRatio As Double
        Dim pHeightRatio As Double
        Dim pCopyIndex As Single
'        On Error Resume Next
        pHeight = pCHeight ''pMemDc.Height
        pWidth = pCWidth  'pMemDc.Width
        
        vbHiMetric = ScaleModeConstants.vbHiMetric
        Prn.ScaleMode = vbPixels
        'pWidth = Prn.ScaleX(pWidth, ScaleModeConstants.vbTwips, Prn.ScaleMode)
        'pHeight = Prn.ScaleY(pHeight, ScaleModeConstants.vbTwips, Prn.ScaleMode)
    
        DI.cbSize = Len(DI)
        DI.lpszDocName = pDocName
        DI.lpszOutput = vbNullString
        DI.lpszDatatype = vbNullString
        
    
         ' Determine if picture should be printed in landscape or portrait
         ' and set the orientation.
         
         If pHeight >= pWidth Then
            Prn.Orientation = vbPRORPortrait   ' Taller than wide.
         Else
            Prn.Orientation = vbPRORLandscape  ' Wider than tall.
         End If
        
         PrnWidth = Prn.ScaleWidth
         PrnHeight = Prn.ScaleHeight
         
         ' Calculate device independent Width to Height ratio for printer.
         'PrnRatio = PrnWidth / PrnHeight
         
         pWidthRatio = PrnWidth / pWidth
         pHeightRatio = PrnHeight / pHeight
         If pWidthRatio < pHeightRatio Then
            PrnPicWidth = pWidth * (IIf(CLng(pWidthRatio) <= 1, 1, CLng(pWidthRatio) - 1))
            PrnPicHeight = pHeight * (IIf(CLng(pWidthRatio) <= 1, 1, CLng(pWidthRatio) - 1))
         Else
            PrnPicWidth = pWidth * (IIf(CLng(pHeightRatio) <= 1, 1, CLng(pHeightRatio) - 1))
            PrnPicHeight = pHeight * (IIf(CLng(pHeightRatio) <= 1, 1, CLng(pHeightRatio) - 1))
         End If
        On Error GoTo 0
        'starts a print job
      
        Call StartDoc(Prn.hdc, DI)
        'prepare the printer driver to accept data
        For pCopyIndex = 1 To Prn.Copies
            Call StartPage(Prn.hdc)
            
            Call StretchBlt(Prn.hdc, 0, 0, PrnPicWidth, PrnPicHeight, pMemDc.hdc, 0, 0, pMemDc.Width, pMemDc.Height, vbSrcCopy)
        'inform the device that the application has finished writing to a page
            Call EndPage(Prn.hdc)
        Next
        'end the print job
        Call EndDoc(Prn.hdc)
End Sub


