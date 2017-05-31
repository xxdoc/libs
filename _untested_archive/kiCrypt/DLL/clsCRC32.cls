VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cCRC32  (clsCRC32.cls)
'
' Description:   The CRC is a very powerful but easily implemented technique
'                to obtain data reliability. The CRC technique is used to
'                protect blocks of data called Frames. Using this technique,
'                the transmitter appends an extra n- bit sequence to every
'                frame called Frame Check Sequence (FCS). The FCS holds
'                redundant information about the frame that helps the
'                transmitter detect errors in the frame. The CRC is one of
'                the most used techniques for error detection in data
'                communications. The technique gained its popularity because
'                it combines three advantages:
'
'                      - Extreme error detection capabilities
'                      - Little overhead
'                      - Ease of implementation
'
' References:    This code was originally written in Visual Basic by David
'                Ireland.  DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Documented, modified, updated variable names
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Moved GetBlockSize() routine to common module.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in table sizing.  Only appeared if last element
'              of table were accessed.
'              Removed obsolete variables.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while calculating the
'                Cyclic Redundancy Check of the data (Calc_CRC32() routine).
'              - Removed obsolete code.
'              - Updated documentation.
' 18-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug.  I was reloading table and resetting total
'                accumulator with every pass.
'              - Renamed LoadTable() routine to ResetVariables().
'              - Created constant MINUS_ONE to replace &HFFFFFFFF.
'              - Moved initial accumulator setting from Calc_CRC32() to
'                ResetVariables() routine.
' 19-Feb-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug.  I was emptying a module level array in the
'              wrong routine.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsCRC32"
  Private Const ZEROES_8    As String = "00000000"
  Private Const MB_5        As Long = &H500000   ' 5242880 bytes
  Private Const MAX_BYTE    As Long = 256
  Private Const MINUS_ONE   As Long = -1
  
' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Long)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a lng Table
' Variable name:     malngTable
'
' ***************************************************************************
  Private mblnArrayLoaded As Boolean
  Private mcurMaxSize     As Currency
  Private mcurAccrued     As Currency
  Private mlngCRC         As Long
  Private malngTable()    As Long


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CRCProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll (tm) v" & App.Major & "." & App.Minor & "." & App.Revision & _
              " Copyright (c) 2004-2017 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Let StopProcessing(ByVal blnData As Boolean)
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
Attribute StopProcessing.VB_Description = "Switch that designates if the user opts to stop processing."
    StopProcessing = gblnStopProcessing
    DoEvents
End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       CRC32_File
'
' Description:   Returns a hex value representing a file.
' 
' Syntax:        ' convert string data to byte array 
'                abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)
' 
' Parameters:    abytFile() - Path/filename in byte array
'                blnCheckSum - [Optional] - TRUE - A checksum must return
'                             a full 8 characters including leading zeros
'                             FALSE - A non-checksum removes leading zeros
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' ***************************************************************************
Public Function CRC32_File(ByRef abytFile() As Byte, _
                  Optional ByVal blnCheckSum As Boolean = False) As String

    Dim hFile        As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strHex       As String
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "CRC32_File"

    On Error GoTo CRC32_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
         
    ' Test incoming data array
    If Not CBool(IsArrayInitialized(abytFile())) Then
        InfoMsg "Incoming array has not been properly initialized." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    Set objBigFiles = New cBigFiles                      ' Instantiate class object
    strSource = ByteArrayToString(abytFile())            ' Convert byte array to string
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL   ' Reset file attributes to normal
    
    mcurAccrued = 0@        ' Init accrued progress
    curFilePos = 0@         ' Set to first position in file
    ResetVariables          ' Load constant table array
    strHex = vbNullString
    
    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        mcurMaxSize = curAmtLeft              ' Save file size for progress calcs
                 
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo CRC32_File_CleanUp
        End If
            
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo CRC32_File_CleanUp
        End If
        
        Do
            Erase abytData()                               ' Empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)        ' Process source file in chunks
            ReDim abytData(lngBlockSize - 1)               ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)   ' Adjust amount of file left to process
            
            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            Calc_CRC32 abytData()   ' Calculate CRC from byte array
        
            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If
    
            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If
            
        Loop
        
        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo CRC32_File_CleanUp
        End If
    
        .API_CloseFile hFile   ' Close all files opened by this routine
    End With
    
    ' convert to hex
    If blnCheckSum Then
        ' A checksum must be a full 8 characters
        strHex = Right$(ZEROES_8 & Hex$(mlngCRC Xor MINUS_ONE), 8)
    Else
        ' A non-checksum drops any leading zeroes
        strHex = Hex$(mlngCRC Xor MINUS_ONE)
    End If
    
    
CRC32_File_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close all files opened by this routine
    Erase abytData()                 ' Always empty arrays when not needed
    Set objBigFiles = Nothing        ' Always free objects when not needed
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strHex = vbNullString
    End If

    CRC32_File = strHex   ' Return hex calculation
    On Error GoTo 0       ' Nullify this error trap
    Exit Function

CRC32_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strHex = vbNullString
    Resume CRC32_File_CleanUp

End Function

' ***************************************************************************
' Routine:       CRC32_String
'
' Description:   Returns a hex value representing a file.
' 
' Parameters:    abytData() - String data in byte array
'                blnCheckSum - [Optional] - TRUE - A checksum must return
'                             a full 8 characters.
'                             FALSE - A non-checksum removes leading zeroes
'
' Example:       CRC32_String("hello world") = 0D4A1185  checksum = TRUE
'                CRC32_String("hello world") =  D4A1185  checksum = FALSE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-JUL-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' ***************************************************************************
Public Function CRC32_String(ByRef abytData() As Byte, _
                    Optional ByVal blnCheckSum As Boolean = False) As String

    Dim strHex As String

    Const ROUTINE_NAME As String = "CRC32_String"

    On Error GoTo CRC32_String_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True   ' Set flag to stop processing
        Exit Function
    End If

    mcurMaxSize = CCur(UBound(abytData))  ' Get string length for progress display
    mcurAccrued = 0@                      ' Init accrued progress
    ResetVariables                        ' Load constant table array
    strHex = vbNullString
    
    ' Calculate CRC from byte array
    If Not Calc_CRC32(abytData()) Then
        InfoMsg "Failed to calculate a CRC value." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo CRC32_String_CleanUp
    End If

    ' convert to hex
    If blnCheckSum Then
        ' A checksum must be a full 8 characters
        strHex = Right$(ZEROES_8 & Hex$(mlngCRC Xor MINUS_ONE), 8)
    Else
        ' A non-checksum drops leading zeroes
        strHex = Hex$(mlngCRC Xor MINUS_ONE)
    End If
    
CRC32_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strHex = vbNullString
    End If

    CRC32_String = strHex   ' Return hex calculation
    Erase abytData()        ' Always empty arrays when not needed
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

CRC32_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strHex = vbNullString
    Resume CRC32_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function Calc_CRC32(ByRef abytData() As Byte) As Boolean

    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngProgress As Long

    On Error GoTo Calc_CRC32_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo Calc_CRC32_CleanUp
    End If

    For lngIndex = 0 To UBound(abytData)
        
        lngIdx = (mlngCRC And &HFF) Xor abytData(lngIndex)          ' Calc table index
        mlngCRC = ((mlngCRC And &HFFFFFF00) \ &H100) And 16777215   ' Shift right 8 bit positions
        mlngCRC = mlngCRC Xor malngTable(lngIdx)                    ' Perform XoR with table selection
        
    Next lngIndex
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo Calc_CRC32_CleanUp
    End If
    
    ' 19-Feb-2012 Update progress bar. Separate
    '             steps for easier debugging
    DoEvents
    mcurAccrued = mcurAccrued + UBound(abytData)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CRCProgress(lngProgress)
    
    Calc_CRC32 = True

Calc_CRC32_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Calc_CRC32 = False
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Calc_CRC32_Error:
    Calc_CRC32 = False
    Resume Calc_CRC32_CleanUp

End Function

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Called by CRC32_File()
    '           CRC32_String()
    
    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select
    
End Function

Private Sub ResetVariables()

    Dim lngCRC   As Long
    Dim lngIdx   As Long
    Dim lngIndex As Long
    
    Const POLYNOMIAL As Long = &HEDB88320   ' -306674912
    
    mlngCRC = MINUS_ONE   ' Initialize accumulator to -1
    
    ' If constant array has already
    ' been loaded then exit this routine
    If mblnArrayLoaded Then
        Exit Sub
    End If
    
    ReDim malngTable(MAX_BYTE)   ' Size constant table array

    ' Calculate contents of table array
    For lngIndex = 0 To MAX_BYTE - 1

        lngCRC = lngIndex

        For lngIdx = 8 To 1 Step -1
            
            If (lngCRC And 1) Then
                lngCRC = ((lngCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                lngCRC = lngCRC Xor POLYNOMIAL
            Else
                lngCRC = ((lngCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If

        Next lngIdx

        malngTable(lngIndex) = lngCRC   ' Load CRC constant table

    Next lngIndex
    
    mblnArrayLoaded = True   ' Flag set to show array has been loaded

End Sub

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    StopProcessing = False
    mblnArrayLoaded = False   ' Flag to show array has NOT been loaded
    ResetVariables            ' Load constant table array

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    Erase malngTable()
    
End Sub
