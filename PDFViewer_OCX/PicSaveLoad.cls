VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PicSaveLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************
' Thanks to Dilettante ++++++++++++
'**********************************


'===========
'PicSaveLoad version 1.1
'===========
'
'A global predeclared object providing simple LoadPicture() and SavePicture()
'methods based on GDI+ for loading and saving bitmap StdPicture objects in BMP,
'GIF, PNG, and JPEG formats from and to disk files.

'Added BMP and corrected EncoderParams passing to GdipSaveImageToStream() for
'GDI+ version 1.1, the default in later versions of Windows.
'

Private Const WIN32_NULL As Long = 0
Private Const WIN32_FALSE As Long = 0
Private Const WIN32_TRUE As Long = 1

Private Enum HRESULT
    S_OK = 0
    E_OUTOFMEMORY = &H8007000E
    E_INVALIDARG = &H80070057
End Enum

Private Enum STGM
    STGM_READ = &H0&
    STGM_WRITE = &H1&
    STGM_READWRITE = &H2&
    
    STGM_SHARE_EXCLUSIVE = &H10&
    STGM_SHARE_DENY_WRITE = &H20&
    STGM_SHARE_DENY_READ = &H30&
    STGM_SHARE_DENY_NONE = &H40&
    
    STGM_FAILIFTHERE = &H0&
    STGM_CREATE = &H1000&
End Enum

Private Declare Function SHCreateStreamOnFile Lib "shlwapi" Alias "SHCreateStreamOnFileW" ( _
    ByVal pszFile As Long, _
    ByVal grfMode As STGM, _
    ByRef stm As IUnknown) As HRESULT

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal pDestination As Long, _
    ByVal pSource As Long, _
    ByVal Length As Long)

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hBitmap As Long
    hPal As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef clsid As IID) As Long

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PICTDESC As PICTDESC, _
    ByRef IID As IID, _
    ByVal fOwn As Long, _
    ByRef Obj As Object) As Long

Private Const BI_RGB As Long = 0&

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Const DIB_RGB_COLORS As Long = 0&

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef Bits As Byte, _
    ByRef bmi As BITMAPINFO, _
    ByVal wUsage As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    ByRef Obj As Any) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Enum GdiplusStatus
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    'Boolean value that specifies whether to suppress the GDI+ background
    'thread. If you set this member to TRUE, GdiplusStartup returns (in
    'its output parameter) a pointer to a hook function and a pointer to
    'an unhook function. You must call those functions appropriately to
    'replace the background thread. If you do not want to be responsible
    'for calling the hook and unhook functions, set this member to FALSE:
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    EncoderGUID As IID
    NumberOfValues As Long
    Type As EncoderParameterValueType
    pValue As Long
End Type

Private Type EncoderParameters
    Count As Long 'Must always be set to 0 or 1 here, we have just one declared below.
    Parameter As EncoderParameter
End Type

Private Enum PixelFormat
    [_PixelFormatIndexed] = &H10000
    [_PixelFormatGDI] = &H20000
    PixelFormat1bppIndexed = &H1& Or &H100& Or [_PixelFormatIndexed] Or [_PixelFormatGDI]
                           '(1 | ( 1 << 8) | PixelFormatIndexed | PixelFormatGDI)
    PixelFormat4bppIndexed = &H2& Or &H400& Or [_PixelFormatIndexed] Or [_PixelFormatGDI]
                           '(2 | ( 4 << 8) | PixelFormatIndexed | PixelFormatGDI)
    PixelFormat8bppIndexed = &H3& Or &H800& Or [_PixelFormatIndexed] Or [_PixelFormatGDI]
                           '(3 | ( 8 << 8) | PixelFormatIndexed | PixelFormatGDI)
    PixelFormat16bppRGB555 = &H5& Or &H1000& Or [_PixelFormatGDI]
                           '(5 | (16 << 8) | PixelFormatGDI)
    PixelFormat24bppRGB = &H8& Or &H1800& Or [_PixelFormatGDI]
                           '(8 | (24 << 8) | PixelFormatGDI)
    PixelFormat32bppRGB = &H9& Or &H2000& Or [_PixelFormatGDI]
                           '(9 | (32 << 8) | PixelFormatGDI)
End Enum

Private Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3
    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7
    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX
    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
    ByRef gdipToken As Long, _
    ByRef StartupInput As GdiplusStartupInput, _
    ByVal pStartupOutput As Long) As GdiplusStatus

Private Declare Function GdiplusShutdown Lib "gdiplus" ( _
    ByVal gdipToken As Long) As GdiplusStatus

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" ( _
    ByVal WidthPixels As Long, _
    ByVal HeightPixels As Long, _
    ByVal Stride As Long, _
    ByVal FormatPixels As PixelFormat, _
    ByRef Scan0 As Byte, _
    ByRef gdipBitmap As Long) As GdiplusStatus

Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" ( _
    ByVal Stream As IUnknown, _
    ByRef hBitmap As Long) As GdiplusStatus

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
    ByVal gdipBitmap As Long, _
    ByRef hBitmap As Long, _
    ByVal background As Long) As GdiplusStatus

Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
    ByVal gdipImage As Long) As GdiplusStatus

Private Declare Function GdipImageRotateFlip Lib "gdiplus" ( _
    ByVal gdipImage As Long, _
    ByVal rfType As RotateFlipType) As GdiplusStatus

Private Declare Function GdipSaveImageToStream Lib "gdiplus" ( _
    ByVal gdipImage As Long, _
    ByVal Stream As IUnknown, _
    ByRef clsidEncoder As IID, _
    ByVal pEncoderParams As Long) As GdiplusStatus

Private IID_IPicture As IID
Private ImageFormatBMP As IID
Private ImageFormatGIF As IID
Private ImageFormatJPEG As IID
Private ImageFormatPNG As IID
Private EncoderQuality As IID
Private gdipStartupInput As GdiplusStartupInput
Private gdipToken As Long

Public Enum SavePictureFormats
    [_fmtMin] = 0
    fmtBMP = 0
    fmtGIF
    fmtJPEG
    fmtPNG
    [_fmtMax] = fmtPNG
End Enum
#If False Then
Dim fmtBMP, fmtGIF, fmtJPEG, fmtPNG
#End If

Public Function LoadPicture( _
    ByVal FileName As String, _
    Optional ByVal BackARGB As Long = &HFFFFFF) As StdPicture
    
    Dim Stream As IUnknown
    Dim HRESULT As HRESULT
    
    If gdipToken = 0 Then Err.Raise &H80044800, TypeName(Me), "GDI+ failed to start"
    Set Stream = Nothing 'Force creation on 64-bit Windows.  Not sure why this is
                         'required or why it works.
    HRESULT = SHCreateStreamOnFile(StrPtr(FileName), _
                                   STGM_READ Or STGM_SHARE_DENY_NONE, _
                                   Stream)
    If HRESULT <> S_OK Then
        Err.Raise &H80044810, _
                  TypeName(Me), _
                  "SHCreateStreamOnFile error " & Hex$(HRESULT)
    End If
    LoadPictureStream LoadPicture, Stream, BackARGB
End Function

Public Sub SavePicture( _
    ByVal StdPicture As StdPicture, _
    ByVal FileName As String, _
    ByVal Format As SavePictureFormats, _
    Optional ByVal JpegQualityPct As Long = 70)

    Dim Stream As IUnknown
    Dim HRESULT As HRESULT

    If (StdPicture Is Nothing) _
        Or ([_fmtMin] > Format Or Format > [_fmtMax]) _
        Or (1 > JpegQualityPct Or JpegQualityPct > 100) Then
            Err.Raise 5, TypeName(Me) 'Invalid procedure call or argument.
    End If
    Set Stream = Nothing 'Force creation on 64-bit Windows.  Not sure why this is
                         'required or why it works.
    HRESULT = SHCreateStreamOnFile(StrPtr(FileName), _
                                   STGM_CREATE Or STGM_WRITE Or STGM_SHARE_EXCLUSIVE, _
                                   Stream)
    If HRESULT <> S_OK Then
        Err.Raise &H80044900, _
                  TypeName(Me), _
                  "SHCreateStreamOnFile error " & Hex$(HRESULT)
    End If
    SavePictureStream StdPicture, Stream, Format, JpegQualityPct
End Sub

Private Sub LoadPictureStream( _
    ByRef StdPicture As StdPicture, _
    ByVal Stream As IUnknown, _
    ByVal BackARGB As Long)
    
    Dim gdipStatus As GdiplusStatus
    Dim gdipBitmap As Long
    Dim hBitmap As Long
    Dim PICTDESC As PICTDESC
    Dim HRESULT As HRESULT

    gdipStatus = GdipCreateBitmapFromStream(Stream, gdipBitmap)
    If gdipStatus <> ok Then
        Err.Raise &H80044A00, _
                  TypeName(Me), _
                  "GdipCreateBitmapFromStream error " & CStr(gdipStatus)
    End If
    gdipStatus = GdipCreateHBITMAPFromBitmap(gdipBitmap, hBitmap, BackARGB)
    If gdipStatus <> ok Then
        GdipDisposeImage gdipBitmap
        Err.Raise &H80044A10, _
                  TypeName(Me), _
                  "GdipCreateHBITMAPFromBitmap error " & CStr(gdipStatus)
    End If
    GdipDisposeImage gdipBitmap
    With PICTDESC
       .cbSizeOfStruct = LenB(PICTDESC)
       .picType = vbPicTypeBitmap
       .hBitmap = hBitmap
       '.hPal = 0
    End With
    HRESULT = OleCreatePictureIndirect(PICTDESC, IID_IPicture, WIN32_TRUE, StdPicture)
    If HRESULT <> S_OK Then
        DeleteObject hBitmap
        Err.Raise &H80044A20, _
                  TypeName(Me), _
                  "OleCreatePictureIndirect error " & Hex$(HRESULT)
    End If
    'Do not: "DeleteObject hBitmap" here because at this point the
    'LoadPicture result (a StdPicture object) owns it and will take
    'care of cleaning it up.
End Sub

Private Sub SavePictureStream( _
    ByVal StdPicture As StdPicture, _
    ByVal Stream As IUnknown, _
    ByVal Format As SavePictureFormats, _
    ByVal JpegQualityPct As Long)

    Dim BITMAP As BITMAP
    Dim BITMAPINFO As BITMAPINFO
    Dim Stride As Long
    Dim Bits() As Byte
    Dim hDC As Long
    Dim LastDllError As Long
    Dim gdipStatus As GdiplusStatus
    Dim gdipBitmap As Long
    Dim EncoderGUID As IID
    Dim Params As EncoderParameters
    Dim pEncoderParams As Long

    If gdipToken = 0 Then Err.Raise &H80044800, TypeName(Me), "GDI+ failed to start"
    
    'It is unsafe to call GdipCreateBitmapFromHBITMAP() on a GDI bitmap or a GDI palette
    'that is currently (or was previously) selected into a device context.  Thus we
    'need to be less direct.  Here we'll use GdipCreateBitmapFromScan0() instead.
    
    If GetObject(StdPicture.Handle, LenB(BITMAP), BITMAP) = 0 Then
        Err.Raise &H80044808, TypeName(Me), "GetObject error " & CStr(Err.LastDllError)
    End If
    With BITMAP
        BITMAPINFO.bmiHeader.biSize = LenB(BITMAPINFO.bmiHeader)
        BITMAPINFO.bmiHeader.biWidth = .bmWidth
        BITMAPINFO.bmiHeader.biHeight = .bmHeight
        BITMAPINFO.bmiHeader.biPlanes = .bmPlanes
        BITMAPINFO.bmiHeader.biCompression = BI_RGB
        BITMAPINFO.bmiHeader.biBitCount = 24
        Stride = (((.bmWidth * 3) + 3) \ 4&) * 4& 'Rounding up to multiple of 4 (i.e. DWORD boundary).
        ReDim Bits(Stride * .bmHeight - 1)
        hDC = GetDC(WIN32_NULL) 'We are cheating a little by using the default monitor hDC here.
        If GetDIBits(hDC, _
                     StdPicture.Handle, _
                     0, _
                     .bmHeight, _
                     Bits(0), _
                     BITMAPINFO, DIB_RGB_COLORS) = 0 Then
            LastDllError = Err.LastDllError
            ReleaseDC WIN32_NULL, hDC
            Err.Raise &H80044810, TypeName(Me), "GetDIBits error " & CStr(LastDllError)
        End If
        ReleaseDC WIN32_NULL, hDC
        gdipStatus = GdipCreateBitmapFromScan0(.bmWidth, _
                                               .bmHeight, _
                                               Stride, _
                                               PixelFormat24bppRGB, _
                                               Bits(0), _
                                               gdipBitmap)
        If gdipStatus <> ok Then
            Err.Raise &H80044818, _
                      TypeName(Me), _
                      "GdipCreateBitmapFromScan0 error " & CStr(gdipStatus)
        End If
    End With
    gdipStatus = GdipImageRotateFlip(gdipBitmap, RotateNoneFlipY)
    If gdipStatus <> ok Then
        GdipDisposeImage gdipBitmap
        Err.Raise &H80044820, _
                  TypeName(Me), _
                  "GdipImageRotateFlip error " & CStr(gdipStatus)
    End If
    Select Case Format
        Case fmtBMP
            EncoderGUID = ImageFormatBMP
        Case fmtJPEG
            EncoderGUID = ImageFormatJPEG
            Params.Count = 1
            With Params.Parameter
                .EncoderGUID = EncoderQuality
                .NumberOfValues = 1
                .Type = EncoderParameterValueTypeLong
                .pValue = VarPtr(JpegQualityPct)
            End With
            pEncoderParams = VarPtr(Params)
        Case fmtGIF
            EncoderGUID = ImageFormatGIF
        Case fmtPNG
            EncoderGUID = ImageFormatPNG
    End Select
    gdipStatus = GdipSaveImageToStream(gdipBitmap, Stream, EncoderGUID, pEncoderParams)
    GdipDisposeImage gdipBitmap
    If gdipStatus <> ok Then
        Err.Raise &H80044858, _
                  TypeName(Me), _
                  "GdipSaveImageToStream error " & CStr(gdipStatus)
    End If
End Sub

Private Sub Class_Initialize()
    CLSIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture
    CLSIDFromString StrPtr("{557CF400-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatBMP
    CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatGIF
    CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatJPEG
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), ImageFormatPNG
    CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), EncoderQuality
    gdipStartupInput.GdiplusVersion = 1
    If GdiplusStartup(gdipToken, gdipStartupInput, WIN32_NULL) <> ok Then
        'On an error gdipToken should be left at its original value (0) anyway, but
        'let's be sure:
        gdipToken = 0
        'We do that because we can't raise exceptions during Class_Initialize and
        'so we'll have to test gdipToken within any procedure that uses GDI+ calls.
        'The 0 value is our indication that GdiplusStartup failed.
    End If
End Sub

Private Sub Class_Terminate()
    If gdipToken <> 0 Then GdiplusShutdown gdipToken
End Sub

