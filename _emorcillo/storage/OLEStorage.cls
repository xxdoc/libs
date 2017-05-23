VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OLEStorage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum OpenFlags

   ' Access
   sfRead = STGM_READ
   sfWrite = STGM_WRITE
   sfReadWrite = STGM_READWRITE
   
   ' Sharing
   sfShareDenyNone = STGM_SHARE_DENY_NONE
   sfShareDenyRead = STGM_SHARE_DENY_READ
   sfShareDenyWrite = STGM_SHARE_DENY_WRITE
   sfShareExclusive = STGM_SHARE_EXCLUSIVE
   sfPriority = STGM_PRIORITY

   ' Creation
   sfCreate = STGM_CREATE
   sfConvert = STGM_CONVERT
   sfFailIfThere = STGM_FAILIFTHERE

   ' Transactioning
   sfDirect = STGM_DIRECT
   sfTransacted = STGM_TRANSACTED

   ' Transactioning Performance
   sfNoScratch = STGM_NOSCRATCH
   sfNoSnapshot = STGM_NOSNAPSHOT

   ' Direct SWMR and Simple
   sfSimple = STGM_SIMPLE
   sfDirectSWMR = STGM_DIRECT_SWMR

   ' Delete On Release
   sfDeleteOnRelease = STGM_DELETEONRELEASE
   
End Enum

'*********************************************************************************************
'
' Creates a new structured storage file and opens it
'
' Parameters:
'
' FileName        - Optional. Full path and name of the file. If omitted a file will be created
'                   in the TEMP folder.
' OpenFlags       - Optional. Flags used to open the file. Default is read/write access without
'                   sharing.
' CreatePropSets  - If true the function creates the standard property sets in the file.
' UseANSIProps    - If true the property set will use ANSI strings.
'
'*********************************************************************************************
Public Function CreateStorageFile( _
      Optional ByVal Filename As String, _
      Optional ByVal OpenFlags As OpenFlags = sfReadWrite Or sfShareExclusive, _
      Optional CreatePropSets As Boolean = True, _
      Optional ByVal UseANSIProps As Boolean = True) As Storage

   On Error GoTo ShowError
   
   ' Create a new Storage object
   Set CreateStorageFile = New Storage
       
   ' Create storage file
   CreateStorageFile.frInitialize StgCreateDocfile(Filename, OpenFlags), Nothing
   
   ' Create the property sets
   If CreatePropSets Then CreateStorageFile.CreatePropertySets UseANSIProps

   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number
   
   If ErrNum > 0 Then
      Err.Raise ErrNum, "CreateStorageFile"
   Else
      Err.Raise ErrNum, "CreateStorageFile", ErrorMessage(ErrNum)
   End If
   
End Function
'*********************************************************************************************
'
' Creates a new stream object from a file
'
' Parameters:
'
' FileName        - Full path and name of the file. If omitted a file will be created
'                   in the TEMP folder.
' OpenFlags       - Optional. Flags used to open the file. Default is read/write access without
'                   sharing.
'
'*********************************************************************************************
Public Function CreateStreamFile( _
      ByVal Filename As String, _
      Optional ByVal OpenFlags As OpenFlags = sfReadWrite Or sfShareExclusive) As Stream

   On Error GoTo ShowError
   
   ' Create a new Storage object
   Set CreateStreamFile = New Storage
       
   ' Create storage file
   CreateStreamFile.frInitialize SHCreateStreamOnFile(Filename, OpenFlags), Nothing

   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number
   
   If ErrNum > 0 Then
      Err.Raise ErrNum, "CreateStorageFile"
   Else
      Err.Raise ErrNum, "CreateStorageFile", ErrorMessage(ErrNum)
   End If
   
End Function

'*********************************************************************************************
'
' Creates a new stream object in memory
'
' Parameters:
'
' hGlobal         - Handle to global memory. If omitted the function will alloc memory for the
'                   stream.
' DeleteOnRelease - If True the memory is released when the stream is destroyed.
'
'*********************************************************************************************
Public Function CreateStreamMemory( _
      Optional ByVal hGlobal As Long = 0, _
      Optional ByVal DeleteOnRelease As Boolean = True) As Stream

   On Error GoTo ShowError
   
   ' Create a new Storage object
   Set CreateStreamMemory = New Storage
       
   ' Create storage file
   CreateStreamMemory.frInitialize CreateStreamOnHGlobal(hGlobal, DeleteOnRelease), Nothing

   Exit Function
   
ShowError:
Dim ErrNum As Long

   ErrNum = Err.Number
   
   If ErrNum > 0 Then
      Err.Raise ErrNum, "CreateStorageFile"
   Else
      Err.Raise ErrNum, "CreateStorageFile", ErrorMessage(ErrNum)
   End If
   
End Function


'*********************************************************************************************
'
' Returns the global memory handle of a memory based stream
'
'*********************************************************************************************
Public Function GetHGlobalFromStream(ByVal Stream As Stream) As Long

   GetHGlobalFromStream = olelib.GetHGlobalFromStream(Stream.Stream)

End Function

'*********************************************************************************************
'
' Returns whether a file is a structured storage file or not
'
'*********************************************************************************************
Public Function IsStorageFile(ByVal Filename As String) As Boolean

   IsStorageFile = olelib.StgIsStorageFile(Filename) = S_OK

End Function


'*********************************************************************************************
' Loads a picture from a file to a StdPicture object
' without converting to BMP
'*********************************************************************************************
Function LoadPicture(ByVal Filename As String) As IPictureDisp
Dim oStrm As IStream
Dim IID_IPicture As UUID

   ' Open the file
   Set oStrm = SHCreateStreamOnFile(Filename, STGM_READ Or STGM_SHARE_EXCLUSIVE)
   
   CLSIDFromString IIDSTR_IPicture, IID_IPicture
   
   ' Load the picture
   Set LoadPicture = olelib.OleLoadPicture(oStrm, 0, False, IID_IPicture)

End Function

'*********************************************************************************************
'
' Opens an storage file
'
' Parameters:
'
' Filename: full path & name of the file
'
'*********************************************************************************************
Public Function OpenStorageFile(ByVal Filename As String, Optional ByVal OpenFlags As OpenFlags = sfReadWrite Or sfShareExclusive) As Storage

   ' Create the object
   Set OpenStorageFile = New Storage
   
   ' Open the file
   OpenStorageFile.frInitialize StgOpenStorage(Filename, Nothing, OpenFlags, vbNullString), Nothing
    
End Function



