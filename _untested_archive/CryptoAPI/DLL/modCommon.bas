Attribute VB_Name = "modCommon"
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME              As String = "modCommon"
  Private Const MAX_BYTE                 As Long = 256
  Private Const MAX_SIZE                 As Long = 260
  Private Const ADD_SPACES               As Long = 50000

' ***************************************************************************
' Global Constants
' ***************************************************************************
  ' password ranges
  Public Const MIN_PWD_LENGTH            As Long = 8
  Public Const MAX_PWD_LENGTH            As Long = 50

  ' miscellaneous
  Public Const DLL_NAME                  As String = "kiCryptoAPI"
  Public Const ENCRYPT_EXT               As String = ".ENC"
  Public Const DECRYPT_EXT               As String = ".DEC"
  Public Const FILE_ATTRIBUTE_NORMAL     As Long = &H80&
  Public Const MOVEFILE_REPLACE_EXISTING As Long = &H1&
  Public Const MOVEFILE_COPY_ALLOWED     As Long = &H2&

' ***************************************************************************
' Module API Declares
' ***************************************************************************
  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid. Returns nonzero if the
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi.dll" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  ' The GetTempPath function retrieves the path of the directory designated
  ' for temporary files.  The GetTempPath function gets the temporary file
  ' path as follows:
  '   1.  The path specified by the TMP environment variable.
  '   2.  The path specified by the TEMP environment variable, if TMP
  '       is not defined.
  '   3.  The current directory, if both TMP and TEMP are not defined.
  Private Declare Function GetTempPath Lib "kernel32.dll" _
          Alias "GetTempPathA" (ByVal nBufferLength As Long, _
          ByVal lpBuffer As String) As Long

  ' The GetTempFileName function creates a name for a temporary file.
  ' The filename is the concatenation of specified path and prefix strings,
  ' a hex string formed from a specified integer, and the .TMP extension.
  Private Declare Function GetTempFileName Lib "kernel32.dll" _
          Alias "GetTempFileNameA" (ByVal lpszPath As String, _
          ByVal lpPrefixString As String, ByVal wUnique As Long, _
          ByVal lpTempFileName As String) As Long

' ***************************************************************************
' Global API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another.  For overlapped blocks, use the RtlMoveMemory function
  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' MoveFileEx Function moves an existing file or directory, including its
  ' children, with various move options.  If successful then return code is
  ' nonzero.
  Public Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" _
         (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
         ByVal dwFlags As Long) As Long

  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
         (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetPath
'
' Description:   Capture complete path up to filename.  Path must end with
'                a backslash.
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Complete path to last backslash
'
' Example:       "C:\Kens Software" <- "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetPath(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetPath = objFSO.GetParentFolderName(strPathFile)
    Set objFSO = Nothing

End Function

' ***************************************************************************
' Routine:       GetFilename
'
' Description:   Capture file name
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Just the file name
'
' Example:       "Gif89.dll" <- "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFilename(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilename = objFSO.GetFilename(strPathFile)
    Set objFSO = Nothing

End Function

' ***************************************************************************
' Routine:       GetFilenameExt
'
' Description:   Capture file name extension
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       File name extension
'
' Example:       "dll" <- "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFilenameExt(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFilenameExt = objFSO.GetExtensionName(strPathFile)
    Set objFSO = Nothing

End Function

' ***************************************************************************
' Routine:       GetVersion
'
' Description:   Capture file version information
'
' Parameters:    strPathFile - Path and file name
'
' Returns:       Version information
'
' Example:       "1.0.0.1" <- "C:\Kens Software\Gif89.dll"
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetFileVerData(ByVal strPathFile As String) As String

    Dim objFSO As New Scripting.FileSystemObject
    GetFileVerData = objFSO.GetFileVersion(strPathFile)
    Set objFSO = Nothing

End Function

' ***************************************************************************
' Routine:       CreateTempFile
'
' Description:   System generated temporary folder and file.  The folder
'                will be located in the Windows default temp directory and
'                is system generated.
'
' Parameters:    strPath - Path to a folder.
'
' Returns:       Unique name of a temporary file
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function CreateTempFile() As String

    Dim strFile As String
    Dim strPath As String

    Const FILE_PREFIX  As String = "~ki"  ' User defined prefix

    strFile = Space$(MAX_SIZE)  ' preload with spaces, not nulls

    ' Locate Windows default temp folder. This
    ' is where Windows creates its temp files.
    strPath = GetTempFolder()

    ' Create a unique temporary file name.
    ' A hex value is returned by the system.
    ' Ex:  "C:\DOCUME~1\Owner\LOCALS~1\Temp\~ki99.tmp"
    GetTempFileName strPath, FILE_PREFIX, 0, strFile

    strFile = TrimStr(strFile)  ' Remove any trialing nulls
    CreateTempFile = strFile    ' Return path\name of temp file

    strFile = vbNullString
    strPath = vbNullString

End Function

' ***************************************************************************
' Routine:       GetTempFolder
'
' Description:   Find system generated temporary folder.
'
' Parameters:    None.
'
' Returns:       Path to the windows default temp folder
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetTempFolder() As String

    Dim strTempFolder As String
    Dim lngRetCode    As Long

    strTempFolder = Space$(MAX_SIZE)        ' preload with spaces, not nulls

    lngRetCode = GetTempPath(MAX_SIZE, strTempFolder)  ' read the path name

    ' Extract data from the variable
    ' Ex:  "C:\DOCUME~1\Owner\LOCALS~1\Temp\"
    If lngRetCode Then
        ' Found Windows default Temp folder.  Remove
        ' any trailing nulls and append backslash
        strTempFolder = TrimStr(strTempFolder)
        strTempFolder = QualifyPath(strTempFolder)
    Else
        ' Did not find Windows default temp folder
        ' therefore, use root level of drive C:
        strTempFolder = "C:\"   ' should never happen
    End If

    ' Return the path and name of the temp file
    GetTempFolder = strTempFolder
    strTempFolder = vbNullString

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
Public Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))

End Function


' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function QualifyPath(ByVal strPath As String, _
                   Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If

End Function

' ***************************************************************************
' Routine:       UnQualifyPath
'
' Description:   Removes a trailing character from the path
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to remove
'                          Default = "\"
'
' Returns:       Fully qualified path without a specific trailing character
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function UnQualifyPath(ByVal strPath As String, _
                     Optional ByVal strChar As String = "\") As String

    strPath = TrimStr(strPath)

    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        UnQualifyPath = Left$(strPath, Len(strPath) - 1)
    Else
        UnQualifyPath = strPath
    End If

End Function

' ***************************************************************************
' Routine:       IsArrayInitialized
'
' Description:   This is an ArrPtr function that determines if the passed
'                array is initialized, and if so will return the pointer
'                to the safearray header. If the array is not initialized,
'                it will return zero. Normally you need to declare a VarPtr
'                alias into msvbvm50.dll or msvbvm60.dll depending on the
'                VB version, but this function will work with vb5 or vb6.
'                It is handy to test if the array is initialized as the
'                return value is non-zero.  Use CBool to convert the return
'                value into a boolean value.
'
'                This function returns a pointer to the SAFEARRAY header of
'                any Visual Basic array, including a Visual Basic string
'                array. Substitutes both ArrPtr and StrArrPtr. This function
'                will work with vb5 or vb6 without modification.
'
'                ex:  If CBool(IsArrayInitialized(array_being_tested)) Then ...
'
' Parameters:    vntData - Data to be evaluated
'
' Returns:       Zero     - Bad data (FALSE)
'                Non-zero - Good data (TRUE)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Mar-2008  RD Edwards
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=69970
' ***************************************************************************
Public Function IsArrayInitialized(ByVal avntData As Variant) As Long

    Dim intDataType As Integer   ' Variable must be a short integer

    On Error GoTo IsArrayInitialized_Exit

    IsArrayInitialized = 0  ' preset to FALSE

    ' Get the real VarType of the argument, this is similar
    ' to VarType(), but returns also the VT_BYREF bit
    CopyMemory intDataType, avntData, 2&

    ' if a valid array was passed
    If (intDataType And vbArray) = vbArray Then

        ' get the address of the SAFEARRAY descriptor
        ' stored in the second half of the Variant
        ' parameter that has received the array.
        ' Thanks to Francesco Balena and Monte Hansen.
        CopyMemory IsArrayInitialized, ByVal VarPtr(avntData) + 8&, 4&

    End If

IsArrayInitialized_Exit:
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       EmptyCollection
'
' Description:   Properly empty and deactivate a collection
'
' Parameters:    colData - Collection to be processed
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub EmptyCollection(ByRef colData As Collection)

    ' Has collection been deactivated?
    If colData Is Nothing Then
        Exit Sub
    End If

    ' Is the collection empty?
    Do While colData.Count > 0

        ' Parse backwards thru collection and delete data.
        ' Backwards parsing prevents a collection from
        ' having to reindex itself after each data removal.
        colData.Remove colData.Count
    Loop

    ' Free collection object from memory
    Set colData = Nothing

End Sub

' **************************************************************************
' Routine:       CalcProgress
'
' Description:   Calculates current amount of completion
'
' Parameters:    curCurrAmt   - Current value
'                curMaxAmount - Maximum value
'
' Returns:       percentage of progression
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function CalcProgress(ByVal curCurrAmt As Currency, _
                             ByVal curMaxAmount As Currency) As Long

    Dim lngPercent As Long

    Const MAX_PERCENT As Long = 100

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

    CalcProgress = lngPercent

End Function

' ***************************************************************************
' Routine:       MixAppendedData
'
' Description:   Performs simple Encryption/Decryption on the information
'                that is to be appended to the original data after normal
'                encryption.  By mixing the appended data you are keeping
'                prying eyes from knowing required information needed to
'                perform decryption easily.  When calling this routine
'                while performing decryption the data will be decrypted.
'
' Parameters:    abytData() - Byte array to be encrypted/decrypted
'                lngMixCount - Optional - Number of passes to mix the data
'                        Default = 5
'
' Returns:       Return data in a byte array.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Simplified mixing process
' 01-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Mix count is now an optional value
'              - Updated documentation
' ***************************************************************************
Public Sub MixAppendedData(ByRef abytData() As Byte, _
                  Optional ByVal lngMixCount As Long = 5)

    Dim lngHigh    As Long
    Dim lngStep    As Long
    Dim lngLoop    As Long
    Dim lngIndex   As Long
    Dim abytTemp() As Byte

    Erase abytTemp()          ' Always start with an empty array
    ReDim abytTemp(MAX_BYTE)  ' Size temp array

    ' Verify number of mixing loops
    ' are within an acceptable range
    Select Case lngMixCount
           Case Is < 2:  lngMixCount = 2   ' Set to minimum
           Case Is > 10: lngMixCount = 10  ' set to maximum
    End Select

    lngHigh = UBound(abytData)
    lngStep = (lngHigh + lngMixCount) Mod MAX_BYTE

    ' Load with ASCII decimal values (0-255)
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytTemp(lngIndex) = CByte(lngIndex)
    Next lngIndex

    ' Extra looping for additional security
    For lngLoop = 1 To lngMixCount

        ' Perform simple encryption/decryption using Xor
        For lngIndex = 0 To lngHigh
            abytData(lngIndex) = abytData(lngIndex) Xor abytTemp((lngStep + lngIndex) Mod MAX_BYTE)
        Next lngIndex

    Next lngLoop

    Erase abytTemp()   ' Always empty array when not needed

End Sub

' ***************************************************************************
' Routine:       ExpandData
'
' Description:   Expand byte array to a designated length.
'
' Parameters:    abytInput() - Incoming byte array
'                lngReturnLen - Output length of return byte array
'
' Returns:       Expanded byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ExpandData(ByRef abytInput() As Byte, _
                           ByVal lngReturnLen As Long) As Byte()

    ' Called by cArcFour.EvaluateKey()
    '           cGost.EvaluateKey()
    '           cTwofish.EvaluateKey()
    '           cSkipjack.EvaluateKey()
    '           cBlowfish.EvaluateKey()
    '           cRijndael.EvaluateKey()
    '           cSerpent.EvaluateKey()

    Dim lngIndex     As Long
    Dim lngStart     As Long
    Dim lngTmpIdx    As Long
    Dim lngInputLen  As Long
    Dim abytTemp()   As Byte
    Dim abytOutput() As Byte

    Const ROUTINE_NAME As String = "ExpandData"

    On Error GoTo ExpandData_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo ExpandData_CleanUp
    End If

    Erase abytOutput()  ' Always start with empty arrays
    Erase abytTemp()

    ReDim abytOutput(lngReturnLen)   ' Resize output array
    lngInputLen = UBound(abytInput)  ' Capture length of input array

    ' Load output array
    For lngIndex = 0 To lngInputLen - 1

        ' Copy data from input array to output array
        abytOutput(lngIndex) = abytInput(lngIndex)

        ' If there is more data than the output
        ' array can hold then exit this loop
        If lngIndex = (lngReturnLen - 1) Then
            Exit For
        End If

    Next lngIndex

    ' Length of incoming data is less than
    ' new output length then insert extra
    ' data into output array
    If lngInputLen < lngReturnLen Then

        lngTmpIdx = 0                            ' Init temp array index
        lngStart = lngIndex                      ' Save last output array position
        abytTemp() = LoadXBoxArray(abytInput())  ' Load temp array with 0-255 mixed

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            GoTo ExpandData_CleanUp
        End If

        ' Load rest of output array
        For lngIndex = lngStart To lngReturnLen - 1
            abytOutput(lngIndex) = abytTemp(lngTmpIdx)  ' Copy temp array to output array
            lngTmpIdx = (lngTmpIdx + 1) Mod MAX_BYTE    ' increment temp array index
        Next lngIndex

    End If

    ExpandData = abytOutput()   ' Return expanded data

ExpandData_CleanUp:
    Erase abytOutput()  ' Always empty arrays when not needed
    Erase abytTemp()
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

ExpandData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ExpandData_CleanUp

End Function

' ***************************************************************************
' Routine:       LoadXBoxArray
'
' Description:   The incoming data array (n bytes) is passed to become part
'                of the mixing process. This routine does not duplicate data
'                in the x-Box array (0-255), just rearranges it.  Duplication
'                allows for missing values in the original data.  Be aware of
'                other mixing routines because they may produce duplicate
'                values during the mixing process.  Note that I do not
'                randomly select any data.  The selection process must be
'                repeatable to be able to encrypt\decrypt data.
'
'                WARNING:  If you make any changes to this routine, verify
'                the end results are repeatable.  Remember, this mixing
'                process deals with both encryption and decryption.
'
' Parameters:    abytInput() - Input byte array
'                lngMixCount - [Optional] - number of iterations used for
'                    mixing the data.  Default = 25
'
' Returns:       Byte array contaning mixed ASCII values 0-255 with no
'                duplicates.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function LoadXBoxArray(ByRef abytInput() As Byte, _
                     Optional ByVal lngMixCount As Long = 25) As Byte()

    Dim lngHigh     As Long   ' Number of array elements
    Dim lngLoop     As Long   ' Loop counter
    Dim lngIndex    As Long   ' Loop counter
    Dim lngNewIdx   As Long   ' Calculated index for swapping
    Dim abytMixed() As Byte   ' Array of mixed values 0-255
    Dim abytTemp()  As Byte   ' Holds input data multiple times

    Const ROUTINE_NAME As String = "LoadXBoxArray"

    On Error GoTo LoadXBoxArray_Error

    Erase abytTemp()   ' Always start with empty arrays
    Erase abytMixed()

    ReDim abytTemp(MAX_BYTE)      ' Size temp array
    ReDim abytMixed(MAX_BYTE)     ' Size output array
    lngHigh = UBound(abytInput)   ' Capture size of incoming array
    lngNewIdx = 7                 ' Starting index (my birth day)

    ' Verify number of mixing loops
    ' are within an acceptable range
    Select Case lngMixCount
           Case Is < 25: lngMixCount = 25   ' Set to minimum
           Case Is > 99: lngMixCount = 99   ' set to maximum
    End Select

    ' Load work arrays
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytMixed(lngIndex) = CByte(lngIndex)                  ' load ASCII decimal array (0-255)
        abytTemp(lngIndex) = abytInput(lngIndex Mod lngHigh)   ' load array based on input data
    Next lngIndex

    ' Outer loop is for obtaining a good mix
    For lngLoop = 1 To lngMixCount

        ' Calculate new index (0-255)
        lngNewIdx = (lngNewIdx + abytTemp(lngNewIdx) + abytMixed(lngNewIdx)) Mod MAX_BYTE

        ' Loop thru array and rearrange data
        For lngIndex = 0 To (MAX_BYTE - 1)

            ' Calculate new index
            lngNewIdx = (lngNewIdx + abytMixed(lngIndex)) Mod MAX_BYTE

            ' If current index and new index are not
            ' the same then swap data with each other
            If lngIndex <> lngNewIdx Then
                SwapBytes abytMixed(lngIndex), abytMixed(lngNewIdx)
            End If

        Next lngIndex
    Next lngLoop

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo LoadXBoxArray_CleanUp
    End If

    LoadXBoxArray = abytMixed()   ' Return mixed data

LoadXBoxArray_CleanUp:
    Erase abytMixed()   ' Always empty arrays when not needed
    Erase abytTemp()

    On Error GoTo 0     ' Nullify error trap in this routine
    Exit Function

LoadXBoxArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadXBoxArray_CleanUp

End Function

' ***************************************************************************
' Routine:       ByteArrayToString
'
' Description:   Converts a byte array to string data
'
' Parameters:    abytData - array of bytes
'
' Returns:       Data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function ByteArrayToString(ByRef abytData() As Byte) As String

    ByteArrayToString = StrConv(abytData(), vbUnicode)

End Function

' ***************************************************************************
' Routine:       StringToByteArray
'
' Description:   Converts string data to a byte array
'
' Parameters:    strData - Data string to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function StringToByteArray(ByVal strData As String) As Byte()

     StringToByteArray = StrConv(strData, vbFromUnicode)

End Function

' ***************************************************************************
' Routine:       ConvertStringToHex
'
' Description:   Take one character at a time and convert first to an
'                integer then to hex.  Prefix with two zeros in case the
'                result is 0x00 to 0x0F (leading zeros tend to disappear).
'                Then capture the last two characters.  This will give a
'                good two character hex display.
'
' Parameters:    blnRetUppercase - (Optional) [Default] - TRUE=Convert data
'                        to uppercase before leaving this routine.
'                        FALSE=Do not convert the data to uppercase.
'
' Returns:       hex data string in uppercase
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************
Public Function ConvertStringToHex(ByVal strInput As String) As String

    Dim strTemp       As String
    Dim strOutput     As String
    Dim lngMax        As Long
    Dim lngIndex      As Long
    Dim lngPointer    As Long
    Dim lngPaddingLen As Long

    ConvertStringToHex = vbNullString
    strTemp = vbNullString

    lngPointer = 1                      ' Index pointer for output string
    lngMax = Len(strInput)              ' Length of input data string
    lngPaddingLen = ADD_SPACES          ' 50,000 blank spaces
    strOutput = Space$(lngPaddingLen)   ' Preload output string

    ' Convert to hex
    For lngIndex = 1 To lngMax

        ' See if some more padding has
        ' to be added to output string
        If (lngPointer + 2) >= lngPaddingLen Then
            strOutput = strOutput & Space$(ADD_SPACES)  ' Append more blank spaces
            lngPaddingLen = Len(strOutput)              ' Capture new output length
        End If

        ' Convert data to hex and
        ' insert into output string
        Mid$(strOutput, lngPointer, 2) = Right$("0" & Hex$(Asc(Mid$(strInput, lngIndex, 1))), 2)
        lngPointer = lngPointer + 2   ' Increment string pointer

    Next lngIndex

    ConvertStringToHex = RTrim$(strOutput)   ' Return hex string

    strOutput = vbNullString

End Function

' ***************************************************************************
' Routine:       ConvertHexToChar
'
' Description:   Convert hex data to ASCII decimal string format.
'
' Parameters:    strHex - Data to be converted
'
' Returns:       data string in ASCII decimal format
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-DEC-2000  Kenneth Ives  kenaso@home.com
'              Wrote routine
' ***************************************************************************
Public Function ConvertHexToChar(ByVal strHex As String) As String

    Dim strTemp       As String
    Dim strOutput     As String
    Dim lngMax        As Long
    Dim lngIndex      As Long
    Dim lngLength     As Long
    Dim lngPointer    As Long
    Dim lngPaddingLen As Long

    ConvertHexToChar = vbNullString
    strTemp = vbNullString

    lngPointer = 1                      ' Index pointer for output string
    lngMax = Len(strHex)                ' Length of input hex string
    lngPaddingLen = ADD_SPACES          ' 50,000 blank spaces
    strOutput = Space$(lngPaddingLen)   ' Preload output string

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        Exit Function
    End If

    ' Convert data
    For lngIndex = 1 To lngMax Step 2

        ' See if some more padding has
        ' to be added to output string
        If (lngPointer + 2) >= lngPaddingLen Then
            strOutput = strOutput & Space$(ADD_SPACES)  ' Append more blank spaces
            lngPaddingLen = Len(strOutput)              ' Capture new output length
        End If

        strTemp = Chr$(Val("&H" & Mid$(strHex, lngIndex, 2)))   ' Convert hex to ASCII character
        lngLength = Len(strTemp)                                ' Save length of converted data
        Mid$(strOutput, lngPointer, lngLength) = strTemp        ' Insert data into output string
        lngPointer = lngPointer + lngLength                     ' Increment output string pointer

    Next lngIndex

    ConvertHexToChar = RTrim$(strOutput)   ' Return data string

    strOutput = vbNullString

End Function

' ***************************************************************************
' Routine:       IsHexData
'
' Description:   Parses a string of data to determine if it is in hex format.
'
' Parameters:    strData - String of data to be evaluated
'
' Returns:       TRUE  - Data string is in hex format
'                FALSE - Not in hex format
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic testing
' ***************************************************************************
Public Function IsHexData(ByRef strData As String) As Boolean

    Dim lngIndex  As Long
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "IsHexData"
    Const HEX_DATA     As String = "0123456789ABCDEF"

    IsHexData = True   ' Preset to TRUE

    ' Prepare data string
    strData = UCase$(strData)             ' Convert to uppercase
    strData = Replace(strData, " ", "")   ' Remove all blank spaces
    strData = Replace(strData, "-", "")   ' Remove all dashes
    strData = Replace(strData, ".", "")   ' Remove all periods
    strData = Replace(strData, "*", "")   ' Remove all asteriks
    strData = Replace(strData, ",", "")   ' Remove all commas
    strData = Replace(strData, "&", "")   ' Remove all ampersand symbols
    strData = Replace(strData, "H", "")   ' Remove all "H" characters

    If StrComp(Left$(strData, 2), "0X", vbBinaryCompare) = 0 Then
        strData = Mid$(strData, 3)   ' Drop first two chars
    End If

    strData = TrimStr(strData)   ' Remove unwanted leading/trailings chars
    lngLength = Len(strData)     ' Capture length of data string

    If lngLength > 0 Then

        ' Parse data string to verify
        ' each character is valid
        For lngIndex = 1 To lngLength

            If InStr(1, HEX_DATA, Mid$(strData, lngIndex, 1)) = 0 Then
                InfoMsg "Invalid character [ " & Mid$(strData, lngIndex, 1) & _
                        " ] found in hex data." & vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                IsHexData = False
                Exit For   ' Found invalid character
            End If

        Next lngIndex
    Else
        InfoMsg "Incoming data string is empty." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        IsHexData = False
    End If

End Function

' ***************************************************************************
' Routine:       SwapBytes
'
' Description:   Swap data with each other without using a temp.  Using Math
'                to do this takes a substantial amount more processing than
'                Logic Gates. The logic gates foundation is in all processors.
'                Which you could argue that so is Math, but running a math
'                process with a temp variable in fact uses a considerable
'                amount more processor cycles than logic gates.
'
'                Using 3 Xor's over 3 Basic Equations greatly speeds up the
'                application. (especially when dealing with millions of
'                equations/Logic Gates)
'
'                This routine works with byte, integer and long values.
'                Change the parameter data type accordingly.
'
' Parameters:    AA - data to be swapped with Value2
'                BB - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Mar-2003  Ulli
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43903&lngWId=1
' ***************************************************************************
Private Sub SwapBytes(ByRef AA As Byte, _
                      ByRef BB As Byte)

    ' Swap byte values

    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub
