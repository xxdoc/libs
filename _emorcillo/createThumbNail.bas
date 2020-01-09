CreateThumbnail

' ----==== GDI+ Declarations ====----

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" ( _
   token As Long, _
   inputbuf As GdiplusStartupInput, _
   Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "GDIPlus" ( _
   ByVal token As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
   ByVal hbm As Long, _
   ByVal hpal As Long, _
   Bitmap As Long) As Long

Private Declare Function GdipGetImageThumbnail Lib "GDIPlus" ( _
   ByVal Image As Long, _
   ByVal thumbWidth As Long, _
   ByVal thumbHeight As Long, _
   thumbImage As Long, _
   ByVal callback As Long, _
   ByVal callbackData As Long) As Long
   
Private Declare Function GdipDisposeImage Lib "GDIPlus" ( _
   ByVal Image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" ( _
   ByVal Bitmap As Long, _
   hbmReturn As Long, _
   ByVal background As Long) As Long

' ----==== OLE API Declarations ====----

Private Type PICTDESC
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As Long
   hPalOrXYExt As Long
End Type

Private Type IID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7)  As Byte
End Type

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" ( _
   lpPictDesc As PICTDESC, _
   riid As IID, _
   ByVal fOwn As Boolean, _
   lplpvObj As Object)

'----------------------------------------------------------
' Procedure : CreateThumbnail
' Purpose   : Creates a thumbnail of a picture
'----------------------------------------------------------
'
Function CreateThumbnail( _
   ByVal Image As StdPicture, _
   ByVal Width As Long, _
   ByVal Height As Long) As StdPicture
Dim tSI As GdiplusStartupInput
Dim lGDIP As Long
Dim lRes As Long
Dim lBitmap As Long

   ' Initialize GDI+
   tSI.GdiplusVersion = 1
   lRes = GdiplusStartup(lGDIP, tSI)
   
   If lRes = 0 Then
   
      ' Create a GDI+ Bitmap from the image handle
      lRes = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBitmap)
   
      If lRes = 0 Then
      
         Dim lThumb As Long
         Dim hBitmap As Long
         
         ' Create the thumbnail
         lRes = GdipGetImageThumbnail(lBitmap, Width, Height, _
                                      lThumb, 0, 0)
      
         If lRes = 0 Then
            
            ' Create a GDI bitmap from the thumbnail
            lRes = GdipCreateHBITMAPFromBitmap(lThumb, hBitmap, 0)
      
            ' Create the StdPicture object
            Set CreatheThumbnail = HandleToPicture(hBitmap, _
                                      vbPicTypeBitmap)
         
            ' Dispose the thumbnail image
            GdipDisposeImage lThumb
         
         End If
         
         ' Dispose the image
         GdipDisposeImage lBitmap
      
      End If
      
      ' Shutdown GDI+
      GdiplusShutdown lGDIP
      
   End If
   
   If lRes Then Err.Raise 5, , "Cannot load file"
   
End Function

'----------------------------------------------------------
' Procedure : HandleToPicture
' Purpose   : Creates a StdPicture object to wrap a GDI
'             image handle
'----------------------------------------------------------
'
Public Function HandleToPicture( _
   ByVal hGDIHandle As Long, _
   ByVal ObjectType As PictureTypeConstants, _
   Optional ByVal hpal As Long = 0) As StdPicture
Dim tPictDesc As PICTDESC
Dim IID_IPicture As IID
Dim oPicture As IPicture
    
   ' Initialize the PICTDESC structure
   With tPictDesc
      .cbSizeOfStruct = Len(tPictDesc)
      .picType = ObjectType
      .hgdiObj = hGDIHandle
      .hPalOrXYExt = hpal
   End With
    
   ' Initialize the IPicture interface ID
   With IID_IPicture
      .Data1 = &H7BF80981
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(3) = &HAA
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
    
   ' Create the object
   OleCreatePictureIndirect tPictDesc, IID_IPicture, _
                            True, oPicture
    
   ' Return the picture object
   Set HandleToPicture = oPicture
        
End Function
