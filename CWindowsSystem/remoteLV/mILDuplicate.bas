Attribute VB_Name = "Module3"
'=========mILDuplicate===========
'*****************************************************************
' Module to duplicate imagelist from remote process
' using remote API call.
' Written by Arkadiy Olovyannikov (ark@msun.ru)
' Copyright 2005 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'*****************************************************************
Option Explicit

Private Type PictDesc
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As _
PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Private Type IMAGEINFO
   hbmImage As Long
   hbmMask As Long
   Unused1 As Long
   Unused2 As Long
   rcImage As RECT
End Type

Private Type BITMAP '14 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'********************************************************************************
'Main function to duplicate remote image list and fill VB ImageList with
'appropriate icons.
'Note: VB ImageList doesn't support new XP style 24/32 bpp icons,
'it uses old 8 bpp icons, so some colors are corrupted.
'If you want keep XP style, use API ImageList_XXX functions to create and fill
'image list and bind this list to ListView with API calls.
'********************************************************************************
Public Function IL_Duplicate(ByVal hProcess As Long, ByVal hIml As Long, il As ImageList) As Long
   
   If (hIml = 0) Or (hProcess = 0) Or (il Is Nothing) Then Exit Function
   Dim nCount As Long, i As Long, ret As Long, hImlNew As Long, hIcon As Long
   Dim ii As IMAGEINFO
   Dim bmp As BITMAP, bmMask As BITMAP
   
   Dim dt() As API_DATA
   Dim abBitmap() As Byte, abMask() As Byte
   Dim dcTemp As Long, dcSrc As Long, dcDest As Long, dcBitmap As Long, dcMask As Long
   Dim hBitmap As Long, hMask As Long, hImage As Long, hTemp As Long
   Dim hOld1 As Long, hOld2 As Long, hOld3 As Long
   Dim cx As Long, cy As Long
   
'Get image count
   ReDim dt(0)
   dt(0).lpData = hIml
   dt(0).argType = arg_Value
   dt(0).dwDataLength = 4
   nCount = CallAPIRemote(hProcess, "comctl32", "ImageList_GetImageCount", 1, dt, 5000)
   If nCount = 0 Then Exit Function
   
'Get bitmap and mask handles from image list, using
'ImageList_GetImageInfo API for the first image_list item
   ReDim dt(2)
   dt(0).lpData = hIml
   dt(0).argType = arg_Value
   dt(0).dwDataLength = 4
   
   dt(1).lpData = 0
   dt(1).argType = arg_Value
   dt(1).dwDataLength = 4
   
   dt(2).lpData = VarPtr(ii)
   dt(2).argType = arg_Pointer
   dt(2).dwDataLength = Len(ii)
   dt(2).bOut = True
   
   ret = CallAPIRemote(hProcess, "comctl32", "ImageList_GetImageInfo", 3, dt, 5000)
   If ret = 0 Then Exit Function
   
   cx = ii.rcImage.Right - ii.rcImage.Left
   cy = ii.rcImage.Bottom - ii.rcImage.Top

'Retrieve BITMAP objects for hBitmap and pass it trough process boundaries
   dt(0).lpData = ii.hbmImage
   dt(0).argType = arg_Value
   dt(0).dwDataLength = 4
   
   dt(1).lpData = Len(bmp)
   dt(1).argType = arg_Value
   dt(1).dwDataLength = 4
   
   dt(2).lpData = VarPtr(bmp)
   dt(2).argType = arg_Pointer
   dt(2).dwDataLength = Len(bmp)
   dt(2).bOut = True
   
   ret = CallAPIRemote(hProcess, "gdi32", "GetObjectA", 3, dt, 5000)
   If ret = 0 Then Exit Function

'Get bitmap bits for hBitmap and pass them trough process boundaries
   ReDim abBitmap(bmp.bmHeight * bmp.bmWidthBytes - 1)
   ret = ReadProcessMemory(hProcess, ByVal bmp.bmBits, abBitmap(0), UBound(abBitmap) + 1, ret)
   If ret = 0 Then Exit Function
   
   bmp.bmBits = VarPtr(abBitmap(0))
'Create new bitmap from HBITMAT structure
   hBitmap = CreateBitmapIndirect(bmp)
   
'Repeat above steps fo hMask bitmap. Actually, it's not nessesarry for VB ImageList
'since VB doesn't support Mask Image, MaskColor only
   If ii.hbmMask Then
      dt(0).lpData = ii.hbmMask
      dt(2).lpData = VarPtr(bmMask)
      ret = CallAPIRemote(hProcess, "gdi32", "GetObjectA", 3, dt, 5000)
      If ret = 0 Then GoTo CleanUp
      If bmMask.bmBits Then
         ReDim abMask(bmMask.bmHeight * bmMask.bmWidthBytes - 1)
         ret = ReadProcessMemory(hProcess, ByVal bmMask.bmBits, abMask(0), UBound(abMask) + 1, ret)
         If ret = 0 Then Exit Function
         bmMask.bmBits = VarPtr(abMask(0))
      Else
         bmMask.bmBits = 0
      End If
      hMask = CreateBitmapIndirect(bmMask)
   End If
   
   dcTemp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   dcSrc = CreateCompatibleDC(dcTemp)
   dcDest = CreateCompatibleDC(dcTemp)
   
   hOld1 = SelectObject(dcSrc, hBitmap)
   hImage = CreateCompatibleBitmap(dcTemp, cx, cy)
   
'Prepare VB ImageList
   il.ListImages.Clear
   il.UseMaskColor = hMask
   If bmMask.bmBits Then il.MaskColor = abMask(0) Else il.MaskColor = 0
   il.ImageWidth = cx
   il.ImageHeight = cy

'Get images from hBitmap one by one and add them to VB ImageList
   For i = -1 To nCount - 1
       hOld2 = SelectObject(dcDest, hImage)
       Call StretchBlt(dcDest, 0, 0, cx, cy, dcSrc, 0, bmp.bmHeight - (i + 1) * cy - 1, cx, -cy, vbSrcCopy)
       hImage = SelectObject(dcDest, hOld2)
       il.ListImages.Add , , BitmapToPicture(hImage)
   Next i
   IL_Duplicate = nCount
'Free objects
CleanUp:
   SelectObject dcSrc, hOld1
   SelectObject dcDest, hOld2
   If hBitmap Then DeleteObject hBitmap
   If hMask Then DeleteObject hMask
   If hImage Then DeleteObject hImage
   DeleteDC dcTemp
   DeleteDC dcSrc
   DeleteDC dcDest
End Function

Private Function IconToPicture(ByVal hIcon As Long) As StdPicture
    If hIcon = 0 Then Exit Function
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As Guid
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .PicType = vbPicTypeIcon
       .hImage = hIcon
    End With
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
    Set IconToPicture = oNewPic
End Function

Private Function BitmapToPicture(ByVal hBmp As Long) As StdPicture
    Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .PicType = vbPicTypeBitmap
       .hImage = hBmp
    End With
    With IGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    Set BitmapToPicture = oNewPic
End Function
