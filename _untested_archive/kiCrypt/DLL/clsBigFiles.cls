VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBigFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'  Module:       cBigFiles  (clsBigFiles.cls)
'
'  Purpose:      This module calculates the size of a file.  Can handle
'                file sizes greater than 2gb.
'
' Reference:     Richard Newcombe  22-Jan-2007
'                Getting Past the 2 Gb File Limit
'                http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
'
'                How To Seek Past VBA's 2GB File Limit
'                http://support.microsoft.com/kb/189981
'
'                How To Write Data to a File Using WriteFile API
'                http://support.microsoft.com/kb/165942/EN-US/
'
'                File Caching
'                http://msdn.microsoft.com/en-us/library/aa364218(VS.85).aspx
'
'                CreateFile Function
'                http://msdn.microsoft.com/en-us/library/aa363858(VS.85).aspx
'
'                INFO: FILE_FLAG_WRITE_THROUGH and FILE_FLAG_NO_BUFFERING
'                http://support.microsoft.com/kb/99794
'
' Description:   The descriptions in this module are excerts from Richard
'                Newcombe's article.
'
'                When working in the IDE, any numbers that are entered are
'                limited to a Long variable type. Actually, as far I've
'                found, the IDE uses Longs for most numeric storage within
'                the projects that you write.
'
'                Okay, so what's the problem with Longs? Well, by definition
'                they are a signed 4-byte variable, in hex &H7FFFFFFF, with a
'                lower limit of -2,147,483,648 and an upper limit of
'                2,147,483,647 (2 Gb). &H80000000 stores the sign of the
'                value. Even when you enter values in Hex, they are stored in
'                a Long.
'
'                Working with random access files, you quite often use a Long
'                to store the filesize and current position, completely
'                unaware that if the file you access is just one byte over
'                the 2 Gb size, you can cause your application to corrupt the
'                file when writing to it.
'
'                Unfortunately, there is no quick fix for this. To get around
'                the problem, you need to write your own file handling
'                module, one that uses windows APIs to open, read, write, and
'                close any file.
'
'                The API's expect the Low and High 32-bit values in unsigned
'                format. Also, the APIs return unsigned values. So, the first
'                thing you have to do is decide on a variable type that you
'                can use to store values higher than 2 Gb. After some serious
'                thought, I decided to use a Currency type (64-bit scaled
'                integer) this gives you a 922,337 gig upper file limit, way
'                bigger that the largest hard drive available today.
'
' AddIn tools    Callers Add-in v3.6 dtd 04-Sep-2016 by RD Edwards (RDE)
' for VB6:       Fantastic VB6 add-in to indentify if a routine calls
'                another routine or is called by other routines within
'                a project.  A must have tool for any VB6 programmer.
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74734&lngWId=1
'
'                NOTE:  Under Windows 10, if you have problems recognizing
'                a VB6 addin, try recompiling it directly into the System32
'                folder.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
' 03-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Documented and modified
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added flag to OpenReadWrite() operations that will not go
'              through any intermediate cache but will go directly to disk.
'              Thanks to Lorin for pointing this out.
' 15-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() routine with IsPathValid() routine.
' 25-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added boolean flag to parameters.  See CreateBigFile().
' 29-Nov-2014  Kenneth Ives  kenaso@tx.rr.com
'              - Corrected calculations in Size2Long(), Long2Size() routines
'              - Updated NumberToHex() routine
'              - Added NumberToBinary(), BinaryToHex(), CalcProgress() routines
'              - Added two Events() and one Property
'              - Minor tweaks and fixes
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
  Private Const MODULE_NAME             As String = "clsBigFiles"
  Private Const FILE_ATTRIBUTE_NORMAL   As Long = &H80
  Private Const FILE_BEGIN              As Long = 0
  Private Const FILE_SHARE_READ         As Long = &H1
  Private Const FILE_SHARE_WRITE        As Long = &H2
  Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
  Private Const GENERIC_READ            As Long = &H80000000
  Private Const GENERIC_WRITE           As Long = &H40000000
  Private Const OPEN_EXISTING           As Long = 3
  Private Const OPEN_ALWAYS             As Long = 4
  Private Const INVALID_HANDLE_VALUE    As Long = -1
  Private Const MAX_LONG                As Long = 2147483647

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  ' The CreateFile function creates or opens the following objects and
  ' returns a handle that can be used to access the object:
  '
  '       Files
  '       pipes
  '       mailslots
  '       communications resources
  '       disk devices (Windows NT only)
  '       consoles
  '       directories (open only)
  Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
          (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, _
          ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long) As Long

  ' CloseHandle invalidates the specified object handle, decrements
  ' the object’s handle count, and performs object retention checks.
  ' Once the last handle to an object is closed, the object is removed
  ' from the operating system. In other words, closes an open file.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' This API is used to set the current position in the open file.
  ' This function is very important because any reads or writes to
  ' the file do not automatically forward the file position.
  Private Declare Function SetFilePointer Lib "kernel32" _
          (ByVal hFile As Long, ByVal lDistanceToMove As Long, _
          lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

  ' Use SetEndOfFile to set the new end of the file marker. This
  ' function moves the end-of-file position for the specified file
  ' to the current position of the file pointer.
  Private Declare Function SetEndOfFile Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' This function writes data to a file, starting at the position
  ' indicated by the file pointer. After the write operation has
  ' been completed, the file pointer is adjusted by the number of
  ' bytes written.
  Private Declare Function WriteFile Lib "kernel32" _
          (ByVal hFile As Long, lpBuffer As Any, _
          ByVal nNumberOfBytesToWrite As Long, _
          lpNumberOfBytesWritten As Long, _
          ByVal lpOverlapped As Any) As Long

  ' This function reads data from a file, starting at the position
  ' indicated by the file pointer. After the read operation has
  ' been completed, the file pointer is adjusted by the number of
  ' bytes read.
  Private Declare Function ReadFile Lib "kernel32" _
          (ByVal hFile As Long, lpBuffer As Any, _
          ByVal nNumberOfBytesToRead As Long, _
          lpNumberOfBytesRead As Long, _
          ByVal lpOverlapped As Any) As Long

  ' GetFileSize determines the size of the file. The file size is
  ' given in a 64-bit value that is split into two 32-bit values.
  ' The high-order half is put into the variable passed as
  ' lpFileSizeHigh; the low-order half is returned by the function.
  ' To get the size, you can either put the binary or hexadecimal
  ' values of the two variables side-by-side, or use the formula
  ' filesize = lpFileSizeHigh * 2^32 + return value. If an error
  ' occurs, the function instead returns -1.
  Private Declare Function GetFileSize Lib "kernel32.dll" _
          (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

  ' Forces Windows to flush the write cache.
  Private Declare Function FlushFileBuffers Lib "kernel32" _
          (ByVal hFile As Long) As Long

  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" _
          Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Currency)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m cur FileSize
' Variable name:     mcurFileSize
' ***************************************************************************
  Private mcurFilesize As Currency


' ***************************************************************************
' ****                         Events                                    ****
' ***************************************************************************
  ' Raises current percentage for the progress bar display
  Public Event FileProgress(ByVal lngPercent As Long)

  ' Raises ongoing pointer status
  Public Event CurrentProgress(ByVal curAmount As Currency, _
                               ByVal curMaxAmount As Currency)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Let FileSize(ByVal curData As Currency)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.FileSize = 2147483647
    mcurFilesize = curData
End Property


' ***************************************************************************
' ***                           Methods                                   ***
' ***************************************************************************

' ***************************************************************************
' Routine:       OpenReadOnly
'
' Description:   Open a file to be used as input.  The file must already
'                exist.  If the file does not exist, an error will occur.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName - Fully qualified path and file name
'                hFile - Numeric value designating an open file
'
' Returns:       TRUE - Successfully opened file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function OpenReadOnly(ByRef strFileName As String, _
                             ByRef hFile As Long) As Boolean

    On Error Resume Next

    OpenReadOnly = False   ' Preset to bad ending

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify file exist
    If IsPathValid(strFileName) Then

        ' Open a file that is to be used as input (read only).
        hFile = CreateFile(strFileName & Chr$(0), _
                           GENERIC_READ, FILE_SHARE_READ, _
                           0&, OPEN_EXISTING, 0, 0)

        ' Error opening file
        If hFile = INVALID_HANDLE_VALUE Then
            gblnStopProcessing = True
            GoTo OpenReadOnly_CleanUp
        End If

        ' Set pointer to start at the beginning of file
        SetFilePointer hFile, 0, 0, FILE_BEGIN
        OpenReadOnly = True

    Else
        InfoMsg "Cannot locate" & vbNewLine & strFileName
    End If

OpenReadOnly_CleanUp:
    If gblnStopProcessing Then
        API_CloseFile hFile
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       OpenReadWrite
'
' Description:   Open a file to update.  If the file exist, if will be
'                opened.  If the file does not exist, it will be created.
'                Use carefully.  If you open an existing file and something
'                goes wrong, the file may become a zero byte file.  There
'                is no recovery of the data available.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName - Fully qualified path and file name
'                hFile - Numeric value designating an open file
'
' Returns:       TRUE - Successfully opened file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 01-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added flag when the file is opened that will allow data to not
'              go through any intermediate cache but will go directly to disk.
'              Thanks to Lorin for pointing this out.
' ***************************************************************************
Public Function OpenReadWrite(ByRef strFileName As String, _
                              ByRef hFile As Long) As Boolean

    On Error Resume Next

    OpenReadWrite = False        ' Preset to bad ending
    RaiseEvent FileProgress(0)   ' Initialize progressbar display

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Open a file that is to be updated.
    '
    ' The flag FILE_FLAG_WRITE_THROUGH is specified so that
    ' system caching is not in effect, then the data is
    ' immediately flushed to disk without going through the
    ' system cache. The operating system also requests a
    ' write-through of the hard disk cache to persistent media.
    '
    ' Note:  Not all hardware supports this write-through capability
    '        with both flags being used at the same time.
    '        FILE_FLAG_WRITE_THROUGH and FILE_FLAG_NO_BUFFERING
    '
    hFile = CreateFile(strFileName & Chr$(0), _
                       GENERIC_READ Or GENERIC_WRITE, _
                       FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                       0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL Or _
                       FILE_FLAG_WRITE_THROUGH, 0)

    ' Error opening file
    If hFile = INVALID_HANDLE_VALUE Then
        gblnStopProcessing = True
        GoTo OpenReadWrite_CleanUp
    End If

    ' Set pointer to start at the beginning of file
    SetFilePointer hFile, 0, 0, FILE_BEGIN
    OpenReadWrite = True

OpenReadWrite_CleanUp:
    If gblnStopProcessing Then
        API_CloseFile hFile
    End If

    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       API_ReadFile
'
' Description:   This routine is used to read data from an opened file.
'
'                This works much like the standard Read # command. You pass
'                the function the file handle you got when opening the file,
'                the byte position you want to start our read from, the size
'                of the data block you require, and a buffer. For the buffer,
'                you use an array of bytes because this seems to be the best
'                method to ensure that any type of data can be read and
'                written with it.
'
'                You use your conversion function to split the position info
'                into the two 32-bit variables needed for the SetFilePointer
'                API. Set the location in your file and then call your
'                ReadFile API. And, for when you read at the end of the file,
'                you pass the number of bytes that have been read.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'                abytData() - Byte array to hold the data that was read
'
' Returns:       TRUE - Successfully read the file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function API_ReadFile(ByVal hFile As Long, _
                             ByVal curPosition As Currency, _
                             ByRef abytData() As Byte) As Boolean

    Dim lngLowOrder  As Long
    Dim lngHighOrder As Long
    Dim lngBytesRead As Long
    Dim lngBlockSize As Long

    Const ROUTINE_NAME As String = "API_ReadFile"

    On Error GoTo API_ReadFile_Error

    API_ReadFile = False   ' Preset to bad ending
    lngLowOrder = 0
    lngHighOrder = 0

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Calculate block of data to be read
    lngBlockSize = (UBound(abytData) + 1) * LenB(abytData(0))

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' Set pointer to start at specific position within file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN

    ' Read this portion of file
    If ReadFile(hFile, abytData(0), lngBlockSize, lngBytesRead, 0&) = 0 Then

        InfoMsg "Failed to read file." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
    Else
        API_ReadFile = True   ' Set flag for successful finish
    End If

API_ReadFile_CleanUp:
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
        API_ReadFile = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

API_ReadFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume API_ReadFile_CleanUp

End Function

' ***************************************************************************
' Routine:       API_WriteFile
'
' Description:   This routine is used to write data to the file.
'
'                This works much like the standard Write # command. You pass
'                the function the file handle you got when opening the file,
'                the byte position you want to start our read from, the size
'                of the data block you require, and a buffer. For the buffer,
'                you use an array of bytes because this seems to be the best
'                method to ensure that any type of data can be read and
'                written with it.
'
'                You use your conversion function to split the position info
'                into the two 32-bit variables needed for the SetFilePointer
'                API. Set the location in your file and then call your
'                ReadFile API. And, for when you read at the end of the file,
'                you pass the number of bytes that have been written.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'                abytData() - Byte array that holds the data to be written
'
' Returns:       TRUE - Successfully updated the file
'                FALSE - An error occurred while accessing the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function API_WriteFile(ByVal hFile As Long, _
                              ByVal curPosition As Currency, _
                              ByRef abytData() As Byte) As Boolean

    Dim lngLowOrder     As Long
    Dim lngHighOrder    As Long
    Dim lngBlockSize    As Long
    Dim lngBytesWritten As Long

    Const ROUTINE_NAME As String = "API_WriteFile"

    On Error GoTo API_WriteFile_Error

    API_WriteFile = False   ' Preset to bad ending
    lngLowOrder = 0
    lngHighOrder = 0

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Calculate block of data to write
    lngBlockSize = (UBound(abytData) + 1) * LenB(abytData(0))

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' Set pointer to start at specific position within file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN

    ' Write to file
    If WriteFile(hFile, abytData(0), lngBlockSize, lngBytesWritten, 0&) = 0 Then

        InfoMsg "Failed to write to file." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True

    Else
        FlushFileBuffers hFile     ' Flush file buffers to force writing
        DoEvents                   '     of data immediately
        CalcProgress curPosition   ' Calculate progress
        API_WriteFile = True       ' Set flag for successful finish
    End If

API_WriteFile_CleanUp:
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        API_CloseFile hFile     ' Verify file handle has been released
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

API_WriteFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume API_WriteFile_CleanUp

End Function

' ***************************************************************************
' Routine:       API_SetEndOfFile
'
' Description:   Sets the pointer to the end of the file designating that
'                we are now finished with this file.
'
' Parameters:    hFile - Numeric value designating an open file
'                curPosition - Current position within the file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Nov-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub API_SetEndOfFile(ByVal hFile As Long, _
                            ByVal curPosition As Currency)

    Dim lngLowOrder  As Long
    Dim lngHighOrder As Long

    lngLowOrder = 0
    lngHighOrder = 0

    ' Calculate current position within file
    Size2Long curPosition, lngLowOrder, lngHighOrder

    ' Set pointer to beginning of file
    SetFilePointer hFile, lngLowOrder, lngHighOrder, FILE_BEGIN

    ' Set end of file marker
    SetEndOfFile hFile

End Sub

' ***************************************************************************
' Routine:       API_CloseFile
'
' Description:   Closes an open file.
'
' Parameters:    hFile - Numeric value designating an open file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub API_CloseFile(ByRef hFile As Long)

    ' Always close a file when not in use or
    ' undesired consequences may happen.
    If hFile > 0 Then
        CloseHandle hFile  ' Release file handle
        hFile = 0          ' Reset handle value
    End If

End Sub

' ***************************************************************************
' Routine:       CalcFileSize
'
' Description:   This routine is used to open a file as read only and
'                calculate it's size.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName  - Name of file
'                curFileSize  - Returned file size in bytes
'                strBitsInHex - OPTIONAL - Return the file size calculated
'                               into bits and in hex format without leading
'                               zeroes. Used for calculating a hash.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated file size calculations.
' ***************************************************************************
Public Sub CalcFileSize(ByRef strFileName As String, _
                        ByRef curFileSize As Currency, _
               Optional ByRef strBitsInHex As String = vbNullString)

    Dim hFile        As Long      ' receives a handle to the file
    Dim lngLowOrder  As Long      ' receive the low-order half of the file size
    Dim lngHighOrder As Long      ' receive the high-order half of the file size
    Dim curHexTemp   As Currency  ' Holding area if greater than 2gb

    Const ROUTINE_NAME As String = "CalcFileSize"

    On Error GoTo CalcFileSize_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    curFileSize = 0@
    lngHighOrder = 0&

    ' Open file as read only and get the file handle
    If OpenReadOnly(strFileName, hFile) Then

        SetFilePointer hFile, 0, 0, FILE_BEGIN              ' Set pointer to beginning of file
        lngLowOrder = GetFileSize(hFile, lngHighOrder)      ' Get API file size values
        API_CloseFile hFile                                 ' Close file handle

        curFileSize = Long2Size(lngLowOrder, lngHighOrder)  ' Calculate actual file size

        '************************
        '* Optional return data *
        '************************
        ' Test for overflow when calculating the
        ' number of bits. This information is
        ' used for generating file hashes.
        If (curFileSize * 8) > MAX_LONG Then

            ' File size is greater than 2gb
            curHexTemp = (curFileSize * 8)
            strBitsInHex = NumberToHex(CStr(curHexTemp))

        Else
            ' File size is up to 2gb
            strBitsInHex = Hex$(curFileSize * 8)
        End If

    Else
        curFileSize = CCur(INVALID_HANDLE_VALUE)
    End If

CalcFileSize_CleanUp:
    If hFile > 0 Then
        API_CloseFile hFile
    End If

    FileSize = curFileSize   ' Update FileSize() property
    On Error GoTo 0          ' Nullify this error trap
    Exit Sub

CalcFileSize_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    curFileSize = 0
    strBitsInHex = vbNullString
    Resume CalcFileSize_CleanUp

End Sub

' **************************************************************************
' Routine:       CreateBigFile
'
' Description:   Creates a null filled file.  These can be greater than
'                2gb.
'
' Parameters:    strFilename - Name of file to be created
'                curFileSize - Size of new file
'                strData - Optional - Data to be used to fill file.
'                          Default = vbNullString
'
' Returns:       TRUE - File was successfully created
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' 25-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added boolean flag to parameters
' ***************************************************************************
Public Function CreateBigFile(ByRef strFileName As String, _
                              ByVal curFileSize As Currency, _
                     Optional ByRef strData As String = vbNullString) As Boolean

    Dim hFile As Long

    Const ROUTINE_NAME As String = "CreateBigFile"

    On Error GoTo CreateBigFile_Error

    CreateBigFile = False   ' Preset to FALSE

    ' If a file by this name already
    ' exist then make sure it is empty
    If IsPathValid(strFileName) Then

        ' Reset file attributes to normal
        SetFileAttributes strFileName, FILE_ATTRIBUTE_NORMAL

        ' Verify file is empty
        hFile = FreeFile
        Open strFileName For Output As #hFile
        Close #hFile
        DoEvents

    End If

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Open file with read/write access
    If OpenReadWrite(strFileName, hFile) Then

        FillFileWithData hFile, curFileSize, strData   ' Load file with null values
        API_CloseFile hFile                            ' Always release file handle

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            GoTo CreateBigFile_CleanUp
        End If

        CreateBigFile = True   ' Good finish
    End If

CreateBigFile_CleanUp:
    If hFile > 0 Then
        API_CloseFile hFile  ' Verify file handle has been released
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CreateBigFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CreateBigFile_CleanUp

End Function

' **************************************************************************
' Routine:       FillFileWithData
'
' Description:   Updates a file from 1 byte to greater than 2gb with null
'                values or optionally user defined data.
'
' Parameters:    hFile - Numeric designator of file to be updated
'                curFileSize - Amount of data to be written to file
'                strData - Optional - Data to be used to fill file.
'                          Default = vbNullString
'
' Returns:       TRUE - File was successfully updated
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Original routine
' 03-Dec-2014  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function FillFileWithData(ByVal hFile As Long, _
                                 ByVal curFileSize As Currency, _
                        Optional ByRef strData As String = vbNullString) As Boolean

    Dim lngIndex    As Long
    Dim curAmtLeft  As Currency
    Dim curPosition As Currency
    Dim abytData()  As Byte

    Const FIVE_MIL     As Long = 5000000   ' 5,000,000 bytes
    Const ROUTINE_NAME As String = "FillFileWithData"

    On Error GoTo FillFileWithData_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    curPosition = 0@             ' Always start at the beginning
    curAmtLeft = curFileSize     ' Save file size for countdown
    RaiseEvent FileProgress(0)   ' Initialize progressbar display

    ' If no file handle then leave
    If hFile < 1 Then
        Exit Function
    End If

    If curAmtLeft >= FIVE_MIL Then

        ReDim abytData(FIVE_MIL)   ' File is larger than five million bytes

        ' See if any user defined data is to be written
        If Len(strData) > 0 Then

            If Len(strData) = 1 Then
                strData = Replicate(FIVE_MIL, strData)
            Else
                strData = Replicate((FIVE_MIL \ Len(strData)) + 1, strData)
                strData = Mid$(strData, 1, FIVE_MIL)
            End If

            For lngIndex = 0 To FIVE_MIL
                abytData(lngIndex) = Mid$(strData, lngIndex + 1, 1)
            Next lngIndex

            strData = vbNullString    ' Empty string variable
        End If

    Else

        ReDim abytData(0 To curAmtLeft - 1)   ' File is less than five million bytes

        ' See if any user defined data is to be written
        If Len(strData) > 0 Then

            If Len(strData) = 1 Then
                strData = Replicate(curAmtLeft, strData)
            Else
                strData = Replicate((curAmtLeft \ Len(strData)) + 1, strData)
                strData = Mid$(strData, 1, curAmtLeft)
            End If

            For lngIndex = 0 To (curAmtLeft - 1)
                abytData(lngIndex) = Mid$(strData, lngIndex + 1, 1)
            Next lngIndex

            strData = vbNullString    ' Empty string variable
        End If

    End If

    Do
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit Do   ' exit Do..Loop
        End If

        API_WriteFile hFile, curPosition, abytData()          ' Write to file
        curPosition = curPosition + CCur(UBound(abytData))    ' Update position in file
        curAmtLeft = curAmtLeft - CCur(UBound(abytData))      ' Calculate amount left to process

        ' See how much data is left to process
        If curAmtLeft < 1 Then
            Exit Do   ' Exit Do..Loop

        ' If amount left to process is less then
        ' five million bytes then resize array to
        ' correct amount for one last write
        ElseIf curAmtLeft < UBound(abytData) Then
            ReDim Preserve abytData(curAmtLeft)
        End If
    Loop

FillFileWithData_CleanUp:
    Erase abytData()   ' Always empty array when not needed
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

FillFileWithData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume FillFileWithData_CleanUp

End Function

' **************************************************************************
' Routine:       CalcProgress
'
' Description:   Calculates current amount of completion
'
' Parameters:    curCurrAmt   - Current value
'                curMaxAmount - Optional - Maximum value (Default = 0)
'
' Returns:       Percentage of progression
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 02-Dec-2014  Kenneth Ives  kenaso@tx.rr.com
'              Modified for this module
' ***************************************************************************
Public Function CalcProgress(ByVal curCurrAmt As Currency, _
                    Optional ByVal curMaxAmount As Currency = 0@) As Long

    Dim lngPercent As Long

    Const MAX_PERCENT As Long = 100

    ' See if any data has been passed
    If curMaxAmount < 1 Then

        ' See if value has been entered
        ' into FileSize() Property
        If mcurFilesize < 1 Then
            curMaxAmount = curCurrAmt     ' Use current amount
        Else
            curMaxAmount = mcurFilesize   ' Use data from FileSize() Property
        End If
    End If

    ' If current amount exceeds
    ' max amount then set to 100%
    If (curCurrAmt >= curMaxAmount) Then
        lngPercent = MAX_PERCENT
    Else
        ' Calculate percentage based
        ' on current and maximum value
        lngPercent = CLng(Round(curCurrAmt / curMaxAmount, 3) * MAX_PERCENT)
    End If

    ' Validate percentage so we
    ' do not exceed maximum bounds
    If lngPercent > MAX_PERCENT Then
        lngPercent = MAX_PERCENT
    End If

    ' Will only pass data if user has this class
    ' defined in their application as WithEvents
    RaiseEvent FileProgress(lngPercent)                    ' Update progressbar display
    RaiseEvent CurrentProgress(curCurrAmt, curMaxAmount)   ' Update current progress status

    ' User may want to manipulate
    ' percentage separately
    CalcProgress = lngPercent

End Function


' ***************************************************************************
' ****                Internal Procedures and Functions                  ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Size2Long
'
' Description:   This routine will work out the higher 32 bits. This code
'                looks like it could be done with a simple division, but
'                you have the problem of the IDE using longs. So, rather
'                than running the risk of the IDE using a long somewhere
'                in the calculations (tests have returned varied results
'                where the division is often 1 out on certain file sizes),
'                you may find a division method that will give the correct
'                value every time, but for now this method will suffice.
'
'                With your max filesize being 922,337 Gb, the highest
'                value that LongHigh will hold is 2,147,630, way below the
'                maximum positive value that a Long can hold. This means
'                that you do not need to monitor it.
'
' Parameters:    curFileSize - File size to be evaluated
'                lngLowOrder - The highest value this will hold is 2,147,483,647
'                lngHighOrder - The highest value this will hold is 2,147,630
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Sub Size2Long(ByVal curFileSize As Currency, _
                      ByRef lngLowOrder As Long, _
                      ByRef lngHighOrder As Long)

    ' Called by API_ReadFile()
    '           API_WriteFile()
    '           API_SetEndOfFile()
    '           FillFileWithData()

    Dim curCutoff As Currency

    curCutoff = CCur(MAX_LONG)
    curCutoff = curCutoff + MAX_LONG
    curCutoff = curCutoff + 1       ' now we hold the value of 4294967295 and not -1

    lngHighOrder = 0   ' Init return values
    lngLowOrder = 0

    ' If the filesize is smaller than 4 Gb, the Do loop does not execute.
    ' Between 4 and 8 Gb the Do loop executes once.  Between 8-12 Gb, it
    ' executes twice.
    Do Until curFileSize < curCutoff
        lngHighOrder = lngHighOrder + 1
        curFileSize = curFileSize - curCutoff
    Loop

    ' Take the remainder and decide whether it needs the sign bit of the
    ' long to hold a value (&H80000000). If it's smaller than 2,147,483,648
    ' you can pass the value directly to the low order. If not, you have to
    ' convert the value to the hex equivalent in signed 32bit.
    If curFileSize > MAX_LONG Then
        lngLowOrder = CLng(-(curCutoff - (curFileSize - 1)))   ' Larger than 2gb
    Else
        lngLowOrder = CLng(curFileSize)   ' Less than 2gb
    End If

End Sub

' ***************************************************************************
' Routine:       Long2Size
'
' Description:   This routine will convert two Long values into one
'                Currency value.  The multiplication here does not
'                give any erroneous results, but note that the currency
'                type variable must be listed first; otherwise, VB will
'                try to use a Long type variable to temporally store
'                the result. (This is one of the minor issues in VB6.
'                When doing calculations, VB uses the same variable
'                type of your first variable in the calculation and not
'                the variable type of the destination variable. This
'                has been documented on several sites.)
'
' Parameters:    lngLowOrder  - Highest value = 2,147,483,647
'                lngHighOrder - Highest value = 2,147,630
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 03-Mar-2008  Kenneth Ives  kenaso@tx.rr.com
'              Renamed, modified and documented
' ***************************************************************************
Private Function Long2Size(ByVal lngLowOrder As Long, _
                           ByVal lngHighOrder As Long) As Currency

    ' Called by CalcFileSize()

    Dim curFileSize As Currency
    Dim curCutoff As Currency

    curCutoff = CCur(MAX_LONG)
    curCutoff = curCutoff + MAX_LONG
    curCutoff = curCutoff + 1       ' now we hold the value of 4294967295 and not -1

    curFileSize = (curCutoff * lngHighOrder)

    If lngLowOrder < 0 Then
        Long2Size = curFileSize + (curCutoff + (lngLowOrder + 1))
    Else
        Long2Size = curFileSize + lngLowOrder
    End If

End Function

' ***************************************************************************
' Routine:       NumberToHex
'
' Description:   Convert a numeric string to hexadecimal format.  Input
'                can be be a positive or negative value.  I keep the
'                conversions in two steps for easier trouble shooting.
'
' Parameters:    strNumber - Numeric value to be converted
'
' Returns:       Hex value representing converted numeric data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function NumberToHex(ByRef strNumber As String) As String

    ' Called by CalcFileSize()

    Dim lngIdx    As Long
    Dim strHex    As String
    Dim strBinary As String

    NumberToHex = vbNullString              ' Empty return variable
    strBinary = NumberToBinary(strNumber)   ' First convert to binary

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    strHex = BinaryToHex(strBinary)         ' Convert binary to hex

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Strip leading zeroes
    For lngIdx = 1 To Len(strHex)

        If Mid$(strHex, lngIdx, 1) <> "0" Then
            strHex = Mid$(strHex, lngIdx)   ' Adjust returning hex string
            Exit For                        ' Exit For..Next loop
        End If

    Next lngIdx

    NumberToHex = strHex   ' Return hex string

End Function

' ***************************************************************************
' Routine:       NumberToBinary
'
' Description:   This function implements the sign and magnitude method.
'                Although it would be acceptable to simply add a negative
'                sign to the far left of any binary number to indicate a
'                negative, this is not how information is stored and used
'                in binary systems, so the sign itself must also be
'                transformed into binary. The easiest way to create negative
'                numbers in binary is to use the most significant bit (the
'                leftmost bit) to act as the sign, where 0 indicates a
'                positive value and 1 indicates a negative value, which is
'                just like putting a negative sign ahead of the number.
'
'                However as one of the bits is used for the sign, it effects
'                the range of values that can be held. For example a byte is
'                eight bits 0000 0000 and if it is unsigned it can hold
'                values between 0-255 (1111 1111(binary) = 255(decimal)).
'                However, if the far left bit is used for the sign, it means
'                there are only 7 binary digits to hold the value of number
'                being expressed meaning the range is then -127 to +127
'                where 0111 1111 = + 127 and 1111 1111 = - 127.  The key is
'                remembering that the most significant bit is on for negative
'                numbers and that's all there is to it.
'
' Parameters:    vntNumber       - Number to be converted
'                lngNumberOfBits - Optional - limit the size of the
'                            returned answer by specifying the number
'                            of bits.  Default = 64
'
' Returns:       Binary string representing converted number
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-Jan-2003  Rick Rothstein
'              http://forums.devx.com/showthread.php?69341-decimal-to-binary-conversion-in-VB6
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function NumberToBinary(ByVal vntNumber As Variant) As String

    ' Called by NumberToHex()

    Dim strBinary       As String
    Dim blnNegative     As Boolean
    Dim lngNumberOfBits As Long

    On Error GoTo NumberToBinary_Error

    NumberToBinary = vbNullString
    strBinary = vbNullString

    ' See if number is less than zero
    If CDec(vntNumber) < 0 Then
        blnNegative = True                 ' Set flag for negative numbers
        vntNumber = Abs(CDec(vntNumber))   ' Change to positive number
    Else
        blnNegative = False                ' Set Flag for positive numbers
    End If

    Do While CDec(vntNumber) > 0
        strBinary = Trim$(Str$(CDec(vntNumber) - 2 * Int(CDec(vntNumber) / 2))) & strBinary
        vntNumber = Int(CDec(vntNumber) / 2)
    Loop

    lngNumberOfBits = Len(strBinary)   ' Determine binary string length

    ' Determine length of binary string
    Select Case lngNumberOfBits
           Case Is > 64: lngNumberOfBits = 96   ' Very big number
           Case Is > 32: lngNumberOfBits = 64   ' Big number
           Case Is > 16: lngNumberOfBits = 32   ' Long Integer
           Case Else:    lngNumberOfBits = 16   ' Short Integer
    End Select

    ' Format return string to correct length
    strBinary = Right$(String$(lngNumberOfBits, "0") & strBinary, lngNumberOfBits)

    ' If value is a negative then
    ' verify first position is "1"
    If blnNegative Then
        Mid$(strBinary, 1, 1) = "1"
    End If

NumberToBinary_CleanUp:
    NumberToBinary = strBinary
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

NumberToBinary_Error:
    Err.Clear                       ' Remove error code
    gblnStopProcessing = True       ' Set flag to stop processing
    strBinary = vbNullString        ' Empty return string
    Resume NumberToBinary_CleanUp

End Function

' ***************************************************************************
' Routine:       BinaryToHex
'
' Description:   Convert a binary string to hexadecimal format.
'
' Parameters:    strBinary - Binary value to be converted
'
' Returns:       Hex value representing converted binary data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function BinaryToHex(ByRef strBinary As String) As String

    ' Called by NumberToHex()

    Dim lngPointer As Long
    Dim strHex     As String
    Dim strBits    As String

    On Error GoTo BinaryToHex_Error

    strHex = vbNullString

    ' Parse binary string
    For lngPointer = 1 To Len(strBinary) Step 4

        ' Capture four character chunks
        strBits = Mid$(strBinary, lngPointer, 4)

        Select Case strBits
               Case "0000": strHex = strHex & "0"
               Case "0001": strHex = strHex & "1"
               Case "0010": strHex = strHex & "2"
               Case "0011": strHex = strHex & "3"
               Case "0100": strHex = strHex & "4"
               Case "0101": strHex = strHex & "5"
               Case "0110": strHex = strHex & "6"
               Case "0111": strHex = strHex & "7"
               Case "1000": strHex = strHex & "8"
               Case "1001": strHex = strHex & "9"
               Case "1010": strHex = strHex & "A"
               Case "1011": strHex = strHex & "B"
               Case "1100": strHex = strHex & "C"
               Case "1101": strHex = strHex & "D"
               Case "1110": strHex = strHex & "E"
               Case "1111": strHex = strHex & "F"
        End Select

    Next lngPointer

BinaryToHex_CleanUp:
    BinaryToHex = strHex         ' Return new value
    On Error GoTo 0              ' Nullify this error trap
    Exit Function

BinaryToHex_Error:
    Err.Clear                    ' Remove error code
    gblnStopProcessing = True    ' Set flag to stop processing
    strHex = vbNullString        ' Empty return string
    Resume BinaryToHex_CleanUp

End Function

' ***************************************************************************
' Routine:       Replicate
'
' Description:   Replicates a string of data.  One to many characters
'                in length.
'
' Parameters:    lngRetLength - Length of the return string
'                strPattern - Data to be replicated
'
' Returns:       Data string with replicated pattern
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Dec-2006  Peter Weighill, pweighill@btinternet.com
'              http://www.xbeat.net/vbspeed/c_Replicate.htm
' 02-Jul-2007  Kenneth Ives  kenaso@tx.rr.com
'              Modifed and documented
' ***************************************************************************
Private Function Replicate(ByVal lngRetLength As Long, _
                           ByRef strPattern As String) As String

    ' Called by FillFileWithData()

    Dim lngWorkLength As Long
    Dim lngPatternLen As Long

    Replicate = vbNullString          ' Verify empty return string
    lngPatternLen = Len(strPattern)   ' Capture pattern length

    ' Verify length desired is greater than zero
    If lngRetLength < 1 Or lngPatternLen < 1 Then
        Exit Function
    End If

    lngWorkLength = (lngRetLength * lngPatternLen)   ' Calc length of return string

    ' Preload return string with first character in pattern
    Replicate = String$(lngWorkLength, strPattern)

    ' See if length of pattern is more than one character
    If lngPatternLen > 1 Then

        ' Insert a copy of pattern starting in first position
        Mid$(Replicate, 1, lngPatternLen) = strPattern

        ' See if return length is greater than pattern length
        If lngRetLength > lngPatternLen Then

            ' Replicate pattern to end of allocated string
            Mid$(Replicate, lngPatternLen + 1) = Replicate

            ' Return correct length of data
            Replicate = Left$(Replicate, lngRetLength)
        Else
            ' Return length is less than or
            ' equal to pattern length
            Replicate = Left$(Replicate, lngRetLength)
        End If

    End If

End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'                Shlwapi.dll first shipped with Internet Explorer 4.0
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))

End Function


