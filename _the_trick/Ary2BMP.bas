
Option Explicit

Private Type GUID
    Data1           As Long
    data2           As Integer
    Data3           As Integer
    Data4(7)        As Byte
End Type
Private Type PicBmp
    size            As Long
    Type            As Long
    hbmp            As Long
    hpal            As Long
    Reserved        As Long
End Type
Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Declare Function SetDIBits Lib "gdi32" ( _
                         ByVal hDC As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO, _
                         ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" ( _
                         ByVal aHDC As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO, _
                         ByVal wUsage As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal Length As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" ( _
                         ByVal hDC As Long, _
                         ByRef pBitmapInfo As BITMAPINFO, _
                         ByVal un As Long, _
                         ByRef lplpVoid As Long, _
                         ByVal handle As Long, _
                         ByVal dw As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" ( _
                         ByRef PicDesc As PicBmp, _
                         ByRef RefIID As GUID, _
                         ByVal fPictureOwnsHandle As Long, _
                         ByRef IPic As IPicture) As Long
Private Declare Function IIDFromString Lib "ole32.dll" ( _
                         ByVal lpsz As Long, _
                         ByRef lpiid As GUID) As Long

Private Sub BMP2Array(bmp As IPicture, data() As Byte)
    Dim bi  As BITMAPINFO
    Dim l   As Long
    
    bi.bmiHeader.biSize = Len(bi.bmiHeader)
    GetDIBits Me.hDC, bmp.handle, 0, 0, ByVal 0, bi, 0
    
    ReDim data(bi.bmiHeader.biWidth * Abs(bi.bmiHeader.biHeight) * 3 - 1)
    bi.bmiHeader.biBitCount = 24
    bi.bmiHeader.biCompression = 0
    bi.bmiHeader.biSizeImage = 0
    
    GetDIBits Me.hDC, bmp.handle, 0, Abs(bi.bmiHeader.biHeight), data(0), bi, 0
    
    CopyMemory l, data(0), 4
    CopyMemory data(0), data(4), l
    
    ReDim Preserve data(l - 1)
    
End Sub

Private Function Array2BMP(data() As Byte) As IPicture
    Dim sWidth      As Long
    Dim sHeight     As Long
    Dim pixels      As Long
    Dim bi          As BITMAPINFO
    Dim hbmp        As Long
    Dim lpPtr       As Long
    
    pixels = -Int(-(UBound(data) + 5) / 3)
    sWidth = -Int(Int(-Sqr(pixels)) / 4) * 4
    sHeight = -Int(-pixels / sWidth)
    
    With bi.bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biBitCount = 24
        .biHeight = sHeight
        .biPlanes = 1
        .biWidth = sWidth
    End With
    
    hbmp = CreateDIBSection(Me.hDC, bi, 0, lpPtr, 0, 0)
    CopyMemory ByVal lpPtr, UBound(data) + 1, 4
    CopyMemory ByVal lpPtr + 4, data(0), UBound(data) + 1
    
    Dim IID_IPictureDisp    As GUID
    Dim pic                 As PicBmp
    
    IIDFromString StrPtr("{7BF80981-BF32-101A-8BBB-00AA00300CAB}"), IID_IPictureDisp
    
    With pic
       .size = Len(pic)
       .Type = vbPicTypeBitmap
       .hbmp = hbmp
    End With

    OleCreatePictureIndirect pic, IID_IPictureDisp, True, Array2BMP
    
End Function

Private Sub Form_Click()
    Dim fnum    As Integer
    Dim data()  As Byte
    Dim data2() As Byte
    Dim lIndex  As Long
    
    fnum = FreeFile
    Open "C:\Temp\videoplayback.mp4" For Binary As fnum
    ReDim data(LOF(fnum) - 1)
    Get fnum, , data()
    Close fnum
    
    Set Picture1.Picture = Array2BMP(data)
    ' Correct?
    BMP2Array Picture1.Picture, data2()
    
    Debug.Assert UBound(data) = UBound(data2)
    
    For lIndex = 0 To UBound(data)
        Debug.Assert data(lIndex) = data2(lIndex)
    Next
    
End Sub