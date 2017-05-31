VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMath32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cMath32  (clsMath32.cls)
'
' Description:   Multiple bit manipulating routines.  Some are not used.  I
'                left the unused ones so I would not have to search for them
'                if I use this module in another application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
' 08-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added four routines to manipulate bits within a byte.
'              Thanks to Alfred Hellmüller for suggesting that I add
'              these routines.
' 09-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added routines w16Shift and w16Rotate for manipulating
'              integer values (-32678 to 32767).
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote NumberToBinary(), BinaryToNumber().
'              - Modified w16Shift(), w16Rotate() to handle short integers.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified BinaryToNumber() to handle byte values and testing
'              required string length.
' 01-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added ByteArrayToHexStringArray() and HexArrayToByteArray() routines
' 12-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added w8Shift() and w8Rotate() routines to manipulate bytes.
' 01-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Correct logic bug in w8Shift() routine.
'              - Added LoadShiftArrays() routine.
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsHexData() and IsBinaryData() routines.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Replaced Err.Raise with InfoMsg so programs will continue
'                to flow smoothly.
'              - Added GetPrimeNumbers() and IsPrime() routines.
' 05-Apr-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsHexData(), IsBinaryData(), BinaryToNumber(),
'              DisplayNumber(), NumberToBinary(), HexToLong(), and
'              HexToBinary() routines.
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Updated GetPrimeNumbers() routine.
'              - Rewrote IsHexData(), IsBinaryData() routines.
' 29-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote LongToInteger(), UnsignedToLong() routines.
' 31-May-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Removed input parameter lngRetDataType.  lngRetDataType now
'                determined within routines.  See BinaryToNumber(),
'                NumberToBinary(), HexToBinary() routines.
'              - Updated NumberToHex() routine for correct return length.
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated LongArrayToString(), UnsignedToLong(), ByteArrayToByteArray(),
'              LongToInteger() and NumberToBinary() routines.
' 05-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a potential bug in w16Rotate() by changing return value
'              to a long integer.  Also updated documentation.
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Updated logic testing in IsHexData() and IsBinaryData()
'                routines.
'              - Renamed and updated ByteArrayToLong() routine.
'              - Removed optional parameter in HexToByteArray() routine.
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug in w8Rotate() routine.
'              - Updated BinaryToNumber(), NumberToBinary(), HexToBinary(),
'                NumberToHex(), ByteArrayToLongArray() routines.
' 22-Jan-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Added EvaluateNumber() routine
'              - Updated calling routines
' 01-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote DisplayNumber() routine
'              - Updated logic in several routines
' 06-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated DisplayNumber() routine
' 15-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Minor tweaking and updating documentation
' ***************************************************************************
'
' Data Measurements
'    Bit            = Single Binary Digit (1 or 0)
'    Byte           = 8 bits
'    Kilobyte  (kB) = 1024 Bytes
'    Megabyte  (MB) = 1024 Kilobytes
'    Gigabyte  (GB) = 1024 Megabytes
'    Terabyte  (TB) = 1024 Gigabytes
'    Petabyte  (PB) = 1024 Terabytes
'    Exabyte   (EB) = 1024 Petabytes
'    Zettabyte (ZB) = 1024 Exabytes
'    Yottabyte (YB) = 1024 Zettabytes
'
'    kB                       1024 = 1024 * 1
'    MB                    1048576 = 1024 * 1024
'    GB                 1073741824 = 1024 * 1048576
'    TB              1099511627776 = 1024 * 1073741824
'    PB           1125899906842624 = 1024 * 1099511627776
'    EB        1152921504606846976 = 1024 * 1125899906842624
'    ZB     1180591620717411303424 = 1024 * 1152921504606846976
'    YB  1208925819614629174706176 = 1024 * 1180591620717411303424
'
' Actual DIsk Space Calculation
'
'    In 1998 the IEC changed it's measurements so that what you consider
'    a gigabyte (1024 mb or 2^30) was renamed a gibibyte.  A gibibyte is
'    now formally recognised as 1000 mb, even though no operating system
'    (like windows) will use this definition.
'
'    Hard drives use the term "gigabytes" which would be what the 300 GB
'    stands for.  But, Microsoft uses a different way to measure gigabytes
'    (they're actually gibibytes) so that's why you "lost" some 21 GB from
'    your hard drive. It happens to everyone.  It is just the conflict of
'    two different numbering systems.
'
'    The difference between those two numbering systems is seven percent.
'    So, when you buy a 300 GB drive, Windows sees it as 279 GB because:
'
'       300 GB - 7% = 279 GB
'
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "clsMath32"
  Private Const MAX_BYTE    As Long = &H100&             '  256
  Private Const MAX_INT     As Long = &H7FFF             '  32767
  Private Const MIN_INT     As Long = &H8000             ' -32768
  Private Const kB_1        As Long = &H400&             '  1024
  Private Const kB_64       As Long = &H10000            '  65536
  Private Const MAX_LONG    As Long = &H7FFFFFFF         '  2147483647
  Private Const MIN_LONG    As Long = &H80000000         ' -2147483648
  Private Const GB_4        As Double = 4294967296#      ' 4294967296 = (2^32)
  Private Const MAX_DWORD   As Double = 4294967295#      ' 4294967295 = (2^32) - 1
  Private Const DBL_LOW     As Double = 0.000000000001
  Private Const DBL_HIGH    As Double = (1.999999999998 / MAX_DWORD)  ' 0.000000000465661287415694

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Private Enum enumBITS
      eBits4  ' 0
      eBits8  ' 1
  End Enum

  Public Enum enumDataType
      eLong      ' 0
      eShort     ' 1
      eByte      ' 2
      eUnknown   ' 3
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Power2
' Variable name:     mabytPower2
'
' ***************************************************************************
  Private mblnLoaded     As Boolean
  Private mabytValue(8)  As Byte
  Private mabytPower2(8) As Byte


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents
End Property

Public Property Get StopProcessing() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       BinaryToHex
'
' Description:   Convert a binary value to its hex equivalent.
'                Please note that it is faster to load an output string
'                with spaces and then insert the data.  Constantly
'                appending data will cause the system to have to
'                reallocate space, re-evaluate new string size and
'                possibly designate a new internal address to hold this
'                data.  All of which takes extra cycles and time.
'
' Input:         00000010100010100010111110011000
' Output:        028A2F98
'
' Parameters:    strBinary - Binary value to be converted to hex
'
' Returns:       Hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function BinaryToHex(ByVal strBinary As String) As String

    Dim lngPos     As Long
    Dim lngPointer As Long
    Dim strHex     As String
    Dim strBits    As String

    On Error GoTo BinaryToHex_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify this is binary data
    If Not IsBinaryData(strBinary) Then
        GoTo BinaryToHex_CleanUp
    End If

    strHex = Space$(64)   ' Preload with spaces
    lngPointer = 1        ' Set output pointer to first position

    For lngPos = 1 To Len(strBinary) Step 4

        ' Capture four character chunks
        strBits = Mid$(strBinary, lngPos, 4)

        Select Case strBits
               Case "0000": Mid$(strHex, lngPointer, 1) = "0"
               Case "0001": Mid$(strHex, lngPointer, 1) = "1"
               Case "0010": Mid$(strHex, lngPointer, 1) = "2"
               Case "0011": Mid$(strHex, lngPointer, 1) = "3"
               Case "0100": Mid$(strHex, lngPointer, 1) = "4"
               Case "0101": Mid$(strHex, lngPointer, 1) = "5"
               Case "0110": Mid$(strHex, lngPointer, 1) = "6"
               Case "0111": Mid$(strHex, lngPointer, 1) = "7"
               Case "1000": Mid$(strHex, lngPointer, 1) = "8"
               Case "1001": Mid$(strHex, lngPointer, 1) = "9"
               Case "1010": Mid$(strHex, lngPointer, 1) = "A"
               Case "1011": Mid$(strHex, lngPointer, 1) = "B"
               Case "1100": Mid$(strHex, lngPointer, 1) = "C"
               Case "1101": Mid$(strHex, lngPointer, 1) = "D"
               Case "1110": Mid$(strHex, lngPointer, 1) = "E"
               Case "1111": Mid$(strHex, lngPointer, 1) = "F"
        End Select

        lngPointer = lngPointer + 1   ' Increment output pointer

    Next lngPos

    strHex = Trim$(strHex)  ' Remove excess blank spaces

BinaryToHex_CleanUp:
    BinaryToHex = strHex   ' Return new value
    On Error GoTo 0        ' Nullify this error trap
    Exit Function

BinaryToHex_Error:
    Err.Clear                   ' Remove error code
    gblnStopProcessing = True   ' Set flag to stop processing
    strHex = vbNullString       ' Empty return string
    Resume BinaryToHex_CleanUp

End Function

' ***************************************************************************
' Routine:       BinaryToNumber
'
' Description:   Converts a binary string to a long integer.  Handles
'                positive and negative values.  Most conversion routines
'                I have looked at will generate the wrong value if the
'                input designates that it is a negative value.  This may
'                be slow but it seems to be accurate.
'
' Input:         Long   11111111111111111111111111111111  ->   -1
'                Short  1111111111111111                  ->   -1
'                Byte   11111111                          ->  255
'
' Parameters:    strBinary   - Binary data string to be converted
'                lngRetDataType - Optional - Determines data type being passed.
'                              Data types are 0-eLong, 1-eShort, 2-eByte
'                              Default = eLong
'
' Returns:       Numeric value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Dec-2000  Donald Leassu donald@xbeat.net
'              http://www.xbeat.net/vbspeed/c_BitToLong.htm
' 28-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified, documented and added short integer capability
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added byte capability and testing required string length.
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made lngRetDataType an optional parameter.  User must now
'              determine type of number being evaluated.
' 05-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic.
' ***************************************************************************
Public Function BinaryToNumber(ByVal strBinary As String, _
                      Optional ByRef lngRetDataType As enumDataType = eUnknown) As Variant

    Dim strHex    As String
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "BinaryToNumber"

    On Error GoTo BinaryToNumber_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    BinaryToNumber = Empty   ' Preset to empty
    strHex = vbNullString

    ' Verify this is binary data
    If Not IsBinaryData(strBinary) Then
        GoTo BinaryToNumber_CleanUp
    End If

    lngLength = Len(strBinary)   ' Capture string length

    ' If all zeroes then leave
    If StrComp(String$(lngLength, "0"), strBinary, vbBinaryCompare) = 0 Then
        BinaryToNumber = 0
        GoTo BinaryToNumber_CleanUp
    End If

    ' Determine return data type
    Select Case lngLength
           Case 32: lngRetDataType = eLong
           Case 16: lngRetDataType = eShort
           Case 8:  lngRetDataType = eByte
    End Select

    strHex = BinaryToHex(strBinary)                        ' First convert binary to hex
    BinaryToNumber = HexToNumber(strHex, lngRetDataType)   ' Convert hex to a number

BinaryToNumber_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

BinaryToNumber_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume BinaryToNumber_CleanUp

End Function

' ***************************************************************************
' Routine:       NumberToHex
'
' Description:   Convert a Long to a Hex string. Handles both positive and
'                negative values.
'
' Parameters:    lngValue   - Number to be converted to string data
'                lngRetDataType - Optional - Determines data type being returned.
'                              Data types are 0-eLong, 1-eShort, 2-eByte
'                              Default = eLong
'
' Returns:       Converted number
'                Ex:   2147483647 --> 7fffffff  (2gb)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function NumberToHex(ByVal vntValue As Variant, _
                   Optional ByVal lngRetDataType As enumDataType = eUnknown) As String

    Dim strHex    As String
    Dim bytValue  As Byte
    Dim intValue  As Integer
    Dim lngValue  As Long
    Dim lngLength As Long

    Const ZEROES_8 As String = "00000000"

    NumberToHex = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        NumberToHex = vbNullString
        Exit Function
    End If

    ' Verify data type is within parameters
    If lngRetDataType = eUnknown Then
        If Not EvaluateNumber(vntValue, lngRetDataType) Then
            Exit Function
        End If
    End If

    strHex = ZEROES_8   ' Preload with zeroes

    ' Determine correct return length
    Select Case lngRetDataType

           Case eLong
                lngLength = 8
                lngValue = CLng(vntValue)
                Mid$(strHex, 9 - Len(Hex$(lngValue))) = Hex$(lngValue)   ' Convert long to hex string

           Case eShort
                lngLength = 4
                intValue = CInt(vntValue)
                Mid$(strHex, 9 - Len(Hex$(intValue))) = Hex$(intValue)   ' Convert integer to hex string

           Case eByte
                lngLength = 2
                bytValue = CByte(vntValue)
                Mid$(strHex, 9 - Len(Hex$(bytValue))) = Hex$(bytValue)   ' Convert byte to hex string

           Case Else
                lngLength = 8
                lngValue = CLng(vntValue)
                Mid$(strHex, 9 - Len(Hex$(lngValue))) = Hex$(lngValue)   ' Convert number to hex string
    End Select

    strHex = Right$(strHex, lngLength)   ' Capture correct length
    NumberToHex = strHex                 ' Return hex string

End Function

' ***************************************************************************
' Routine:       NumberToBinary
'
' Description:   Converts a number (without decimals) to a binary string.
'                Handles both positive and negative values.
'
' Parameters:    lngValue    - numeric value to be converted
'                lngRetDataType - Optional - Determines data type being passed.
'                              Data types are 0-eLong, 1-eShort, 2-eByte
'                              Default = eLong
'
' Returns:       Binary data string
'
' Example:       Long    -1  ->  11111111111111111111111111111111
'                Short   -1  ->  1111111111111111
'                Byte   255  ->  11111111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Dec-2000  Peter Nierop
'              http://www.xbeat.net/vbspeed/c_LongToBit.htm
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified, documented and added short integer capability
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Verify binary strings are proper length
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made lngRetDataType an optional parameter.  User must now
'              determine type of number being evaluated.
' 05-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function NumberToBinary(ByVal vntValue As Variant, _
                      Optional ByVal lngRetDataType As enumDataType = eUnknown) As String

    Dim bytValue  As Byte
    Dim intValue  As Integer
    Dim lngValue  As Long
    Dim strBinary As String
    Dim astrBin() As String

    Const ROUTINE_NAME As String = "NumberToBinary"

    On Error GoTo NumberToBinary_Error

    Erase astrBin()            ' Always start with empty arrays
    strBinary = vbNullString   ' Empty variable

    ' Verify data type is within parameters
    If lngRetDataType = eUnknown Then
        If Not EvaluateNumber(vntValue, lngRetDataType) Then
            Exit Function
        End If
    End If

    Select Case lngRetDataType

           Case eLong  ' Long integer returns 32 char binary string

                lngValue = CLng(vntValue)
                astrBin() = LoadBinaryArray(eBits8)  ' load binary equivalents

                ' Process a negative value
                If lngValue < 0 Then

                    lngValue = lngValue And &H7FFFFFFF

                    strBinary = astrBin(128 + lngValue \ &H1000000 And &HFF&) & _
                                astrBin(lngValue \ &H10000 And &HFF&) & _
                                astrBin(lngValue \ &H100& And &HFF&) & _
                                astrBin(lngValue And &HFF&)
                Else
                    ' Process a positive value
                    strBinary = astrBin(lngValue \ &H1000000 And &HFF&) & _
                                astrBin(lngValue \ &H10000 And &HFF&) & _
                                astrBin(lngValue \ &H100& And &HFF&) & _
                                astrBin(lngValue And &HFF&)
                End If

                ' Pad with zeroes to maximum valid length
                strBinary = Right$(String$(32, "0") & strBinary, 32)

                If lngValue < 0 Then
                    Mid$(strBinary, 1, 1) = "1"  ' Designates a negative value
                End If

           Case eShort  ' Short integer returns 16 char binary string

                intValue = CInt(vntValue)
                astrBin() = LoadBinaryArray(eBits4)  ' load binary equivalents

                ' Process a negative value
                If intValue < 0 Then

                    intValue = intValue And &H7FFF&

                    strBinary = astrBin(128 + intValue \ &H1000& And &HF&) & _
                                astrBin(intValue \ &H100& And &HF&) & _
                                astrBin(intValue \ &H10& And &HF&) & _
                                astrBin(intValue And &HF&)
                Else
                    ' Process a positive value
                    strBinary = astrBin(intValue \ &H1000& And &HF&) & _
                                astrBin(intValue \ &H100& And &HF&) & _
                                astrBin(intValue \ &H10& And &HF&) & _
                                astrBin(intValue And &HF&)
                End If

                ' Pad with zeroes to maximum valid length
                strBinary = Right$(String$(16, "0") & strBinary, 16)

                If lngValue < 0 Then
                    Mid$(strBinary, 1, 1) = "1"  ' Designates a negative value
                End If

           Case eByte  ' Byte returns 8 char binary string

                bytValue = CByte(vntValue)
                astrBin() = LoadBinaryArray(eBits4)  ' load binary equivalents

                strBinary = astrBin(bytValue \ &H10& And &HF&) & _
                            astrBin(bytValue And &HF&)

                ' Pad with zeroes to maximum valid length
                strBinary = Right$(String$(8, "0") & strBinary, 8)
    End Select

    NumberToBinary = strBinary   ' Return new binary string

NumberToBinary_CleanUp:
    Erase astrBin()   ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

NumberToBinary_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    strBinary = vbNullString
    Resume NumberToBinary_CleanUp

End Function

' ***************************************************************************
' Routine:       ByteArrayToByteArray
'
' Description:   Transfers contents of one byte array to another
'
' Parameters:    abytData() - Incoming array of bytes
'
' Returns:       Byte array that mirrors incoming array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' nn-Aug-2001  DI Management Services Pty Limited
'              www.di-mgt.com.au/crypto.html
' 30-Jun-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' ***************************************************************************
Public Function ByteArrayToByteArray(ByRef abytData() As Byte) As Byte()

    Dim abytTemp() As Byte

    Const ROUTINE_NAME As String = "ByteArrayToByteArray"

    On Error GoTo ByteArrayToByteArray_Error

    Erase abytTemp()  ' Always start with an empty array

    ' See if the incoming array is properly initialized
    If CBool(IsArrayInitialized(abytData())) Then

        ReDim abytTemp(UBound(abytData) + 1)                    ' Resize temp array one byte larger
        CopyMemory abytTemp(0), abytData(0), UBound(abytData)   ' Transfer data to Temp array
        ReDim Preserve abytTemp(UBound(abytData))               ' Resize temp array to match input array

    Else
        InfoMsg "Incoming data array has not been properly prepared." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    End If

ByteArrayToByteArray_CleanUp:
    ByteArrayToByteArray = abytTemp()  ' Return new array of data

    Erase abytTemp()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ByteArrayToByteArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ReDim abytTemp(0)
    Resume ByteArrayToByteArray_CleanUp

End Function

' ***************************************************************************
' Routine:       ByteArrayToLong (formerly named BytesToLong)
'
' Description:   Convert data from a byte array into a long integer. This
'                routine assumes that the byte array will have at least
'                4 elements.
'
' Reference:     Convert 4 Bytes to Long in VB
'                Filipe Lage
'                http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=56804&lngWId=1
'
' Parameters:    abytData() - Array to hold the data
'                lngIdx     - position to start within the array
'
' Returns:       Long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'
'              The problem arose when converting the hex data back to a long
'              integer (4 byte storate area).
'
'                    Ex:  &HA780 -> -22656  (Bad)
'
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'
'                    Ex:  &HA780& -> 42880  (Good)
'
'              Karl Peterson comments:
'              "The trailing ampersand tells VB it's a Long value (4 byte),
'              even for values that could be held in an Integer (2 byte).
'              Basically, if it's &h8000-&hFFFF, use the ampersand (actually,
'              it's never a bad idea), assuming you really do want to store
'              it in a 4-byte variable. Otherwise, VB strives to confuse, by
'              twiddling the sign bit for those higher values."
'              http://bytes.com/topic/visual-basic-net/answers/387151-ampersands-hexadecimal-numbers-whats-rule
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Renamed and updated this routine
' ***************************************************************************
Public Function ByteArrayToLong(ByRef abytData() As Byte, _
                       Optional ByVal lngIdx As Long = 0) As Long

    Dim strHex As String

    Const ROUTINE_NAME As String = "ByteArrayToLong"

    On Error GoTo ByteArrayToLong_Error

    ByteArrayToLong = 0
    strHex = vbNullString

    ' Evaluate incoming array
    If CBool(IsArrayInitialized(abytData())) Then

        ' Evaluate array index pointer
        Select Case lngIdx

               Case Is < 0
                    InfoMsg "The starting position must zero or greater." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3

               Case Is >= UBound(abytData)
                    InfoMsg "Starting position in byte array exceeds size of array." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3

               Case Is > (UBound(abytData) - 3)
                    InfoMsg "Incoming array does not have enough data to convert." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3

               Case Else
                    ' Convert to hex string then to long integer
                    strHex = CStr("&H" & Right$("0" & Hex$(abytData(lngIdx)), 2) & _
                                         Right$("0" & Hex$(abytData(lngIdx + 1)), 2) & _
                                         Right$("0" & Hex$(abytData(lngIdx + 2)), 2) & _
                                         Right$("0" & Hex$(abytData(lngIdx + 3)), 2))

                    ' See flowerbox above for detail explanation
                    If (Val(strHex) >= &H8000) And _
                       (Val(strHex) <= &HFFFF) Then

                        strHex = strHex & "&"
                    End If

                    ByteArrayToLong = Val(strHex)
        End Select
    Else
        InfoMsg "Incoming array has not been initialized." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
    End If

ByteArrayToLong_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ByteArrayToLong_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ByteArrayToLong_CleanUp

End Function

' ***************************************************************************
' Routine:       ByteArrayToLongArray
'
' Description:   Convert data from a byte array into a long integer array.
'                This routine assumes that the byte array will have at least
'                4 elements.
'
' Parameters:    abytData() - Incoming data in byte format
'                lngPointer - position to start within the array
'                lngReturnSize - number of long integers that will be in the
'                     return array
'
' Returns:       Long integer array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 08-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed:  Used wrong pointer to byte array.
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic flow.
' ***************************************************************************
Public Function ByteArrayToLongArray(ByRef abytData() As Byte, _
                            Optional ByVal lngPointer As Long = 0, _
                            Optional ByVal lngReturnSize As Long = 1) As Long()

    Dim lngIdx       As Long
    Dim lngIndex     As Long
    Dim lngValue     As Long
    Dim lngByteCount As Long
    Dim alngOutput() As Long

    Const ROUTINE_NAME As String = "ByteArrayToLongArray"

    On Error GoTo ByteArrayToLongArray_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Evaluate incoming array
    If CBool(IsArrayInitialized(abytData())) Then

        Select Case UBound(abytData)

               Case Is < 1   ' make sure there is at least one byte
                    InfoMsg "There is not enough data in the incoming array to convert." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3

               Case Is < lngByteCount   ' make sure there is enough incoming bytes to convert
                    InfoMsg "There is not enough data in the incoming array to " & _
                            "create the number of long integers requested." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        End Select

        ' Test pointer value
        Select Case lngPointer

               Case Is < 0
                    InfoMsg "The starting position must zero or greater." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                    GoTo ByteArrayToLongArray_CleanUp

               Case Is >= UBound(abytData)
                    InfoMsg "Starting position in byte array exceeds size of array." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                    GoTo ByteArrayToLongArray_CleanUp

               Case Is > (UBound(abytData) - 3)
                    InfoMsg "Incoming array does not have enough data to convert." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                    GoTo ByteArrayToLongArray_CleanUp
        End Select

        lngIdx = 0                         ' Initialize new array index
        lngByteCount = lngReturnSize * 4   ' calc number of bytes needed
        ReDim alngOutput(lngReturnSize)    ' resize output array

        ' Parse the data in 4-byte chunks and convert
        ' each chunk to a long integer. Insert long
        ' integer into the return array.
        For lngIndex = lngPointer To (lngByteCount - 1) Step 4

            lngValue = ByteArrayToLong(abytData(), lngIndex)   ' convert bytes to long integer
            alngOutput(lngIdx) = lngValue                      ' save long integer to long array
            lngIdx = lngIdx + 1                                ' increment output index

            ' Reached the number of requested elements
            If lngIdx = lngReturnSize Then
                Exit For    ' exit For..Next loop
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

        Next lngIndex

    Else
        InfoMsg "Incoming array has not been initialized." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
    End If

ByteArrayToLongArray_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ReDim alngOutput(0)                   ' return a single zero
    Else
        ReDim Preserve alngOutput(lngIdx - 1) ' resize to exact count
    End If

    ByteArrayToLongArray = alngOutput()

    Erase alngOutput()  ' Always empty arrays when not needed
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

ByteArrayToLongArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    ReDim alngOutput(0)
    Resume ByteArrayToLongArray_CleanUp

End Function

' **************************************************************************
' Routine:       DisplayNumber
'
' Description:   Return a string representing the value in string format
'                to requested number of decimal positions.  Can handle
'                both positive and negative values.
'
'                Data Measurements
'                   Bit            = Single Binary Digit (1 or 0)
'                   Byte           = 8 bits
'                   Kilobyte  (kB) = 1024 Bytes
'                   Megabyte  (MB) = 1024 Kilobytes
'                   Gigabyte  (GB) = 1024 Megabytes
'                   Terabyte  (TB) = 1024 Gigabytes
'                   Petabyte  (PB) = 1024 Terabytes
'                   Exabyte   (EB) = 1024 Petabytes
'                   Zettabyte (ZB) = 1024 Exabytes
'                   Yottabyte (YB) = 1024 Zettabytes
'
'                   kB                       1024 = 1024 * 1
'                   MB                    1048576 = 1024 * 1024
'                   GB                 1073741824 = 1024 * 1048576
'                   TB              1099511627776 = 1024 * 1073741824
'                   PB           1125899906842624 = 1024 * 1099511627776
'                   EB        1152921504606846976 = 1024 * 1125899906842624
'                   ZB     1180591620717411303424 = 1024 * 1152921504606846976
'                   YB  1208925819614629174706176 = 1024 * 1180591620717411303424
'
' Parameters:    vntNumber   - Value to be evaluated
'                lngDecimals - Optional - Number of decimal positions.
'                     Valid values are 0-5.  Change to meet special needs.
'                     Default value = 1 decimal position
'
' Returns:       Reformatted string representation
'
'                Ex:  987654321012345      898.3 TB    ' Positive data
'                    -987654321012345     -898.3 TB    ' Negative data
'                    (987654321012345)    (898.3) TB   ' Negative data
'
' ===========================================================================
'    DATE      NAME \ eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 06-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              - Added negative value functionality
'              - Updated documentation
' ***************************************************************************
Public Function DisplayNumber(ByVal vntNumber As Variant, _
                     Optional ByVal lngDecimals As Long = 1) As String

    Dim lngPointer   As Long
    Dim blnNegative  As Boolean
    Dim blnMinusSign As Boolean

    Const kB_1       As Double = 1024#
    Const MAX_POS    As Long = 5

    On Error GoTo DisplayNumber_CleanUp

    '-------------------------------------------------------------
    ' Test data
    ' Uncomment line you want to use
    '-------------------------------------------------------------
    ' vntNumber = "987"                           ' Bytes
    ' vntNumber = "987654"                        ' kB
    ' vntNumber = "987654321"                     ' MB
    ' vntNumber = "987654321012"                  ' GB
    ' vntNumber = "987654321012345"               ' TB
    ' vntNumber = "987654321012345678"            ' PB
    ' vntNumber = "987654321012345678909"         ' EB
    ' vntNumber = "987654321012345678909876"      ' ZB
    ' vntNumber = "987654321012345678909876543"   ' YB
    '
    ' vntNumber = "9.87654321012346E+29"          ' Bad data
    ' vntNumber = "-987654321012"                 ' Negative data
    ' vntNumber = "(987654321012)"                ' Negative data
    '-------------------------------------------------------------

    DisplayNumber = ""    ' Verify return string is empty
    lngPointer = 0        ' Set pointer to zero
    blnNegative = False   ' Preset to FALSE
    blnMinusSign = False

    If IsNumeric(vntNumber) Then

        ' Verify decimal positions
        ' are within acceptable range
        Select Case lngDecimals
               Case Is < 0:       lngDecimals = 0         ' No decimal positions
               Case Is > MAX_POS: lngDecimals = MAX_POS   ' Maximum decimal positions
        End Select

        ' See if input value is zero
        If vntNumber = 0 Then

            ' Determine number of decimal
            ' positions in return string
            If lngDecimals > 0 Then

                ' One or more decimal positions
                DisplayNumber = FormatNumber$(vntNumber, lngDecimals) & " Bytes"
            Else
                ' No decimal positions
                DisplayNumber = Format$(vntNumber, "0") & " Bytes"
            End If

            GoTo DisplayNumber_CleanUp
        End If

        ' See if value is an exponential number
        If InStr(1, CStr(vntNumber), "E", vbTextCompare) > 0 Then
            InfoMsg "An exponential value  (" & CStr(vntNumber) & ")" & _
                    vbNewLine & "cannot be properly processed.", , , 5
            GoTo DisplayNumber_CleanUp
        End If

        ' See if input value is negative
        If vntNumber < 0 Then

            blnNegative = True   ' Set flag to TRUE

            If InStr(1, vntNumber, "-") > 0 Then
                blnMinusSign = True                       ' Set flag to TRUE
                vntNumber = Replace(vntNumber, "-", "")   ' Remove minus sign
            Else
                vntNumber = Replace(vntNumber, "(", "")   ' Remove left parenthesis
                vntNumber = Replace(vntNumber, ")", "")   ' Remove right parenthesis
            End If
        End If

        ' If value is less than 1024
        ' then return bytes only
        If vntNumber < kB_1 Then

            DisplayNumber = CStr(vntNumber)   ' Capture number

            ' Flagged as a negative value?
            If blnNegative Then

                If blnMinusSign Then
                    DisplayNumber = "-" & vntNumber         ' Prefix with minus sign
                Else
                    DisplayNumber = "(" & vntNumber & ")"   ' Wrap in parenthesis
                End If
            End If

            DisplayNumber = DisplayNumber & " Bytes"
            GoTo DisplayNumber_CleanUp
        End If

        ' Keep dividing numeric data until
        ' value is less than or equal to 1024
        Do While CDbl(vntNumber) >= kB_1
            vntNumber = CDec(vntNumber / kB_1)
            lngPointer = lngPointer + 1
        Loop

        ' Determine number of decimal
        ' positions in return string
        If lngDecimals > 0 Then

            ' One or more decimal positions
            DisplayNumber = FormatNumber$(vntNumber, lngDecimals)
        Else
            ' No decimal positions
            DisplayNumber = Format$(vntNumber, "0")
        End If

        ' Flagged as a negative value?
        If blnNegative Then

            If blnMinusSign Then
                DisplayNumber = "-" & DisplayNumber         ' Prefix with minus sign
            Else
                DisplayNumber = "(" & DisplayNumber & ")"   ' Wrap in parenthesis
            End If
        End If

        ' Append designation size
        If lngPointer > 0 Then
            DisplayNumber = DisplayNumber & " " & _
                            Choose(lngPointer, "kB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
        End If

    Else
        InfoMsg "Input value (" & CStr(vntNumber) & ") is not a valid number."
    End If

    ' Debug.Print DisplayNumber
    ' Stop

DisplayNumber_CleanUp:
    Err.Clear           ' Reset any error code to zero
    vntNumber = Empty
    On Error GoTo 0     ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       DoubleToLong
'
' Description:   Convert a double precision number to a long integer.
'
' Parameters:    dblValue - Number to be converted
'
' Returns:       long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function DoubleToLong(ByVal dblValue As Double) As Long

    DoubleToLong = UnsignedToLong(Fix(dblValue / MAX_DWORD))

End Function

' ***************************************************************************
' Routine:       ExtractFromLongArray
'
' Description:   This function creates a new array from a selected section
'                of an array.  The original array is unaffected.  This
'                function only works with an array of Longs.
'
' Parameters:    alngSource() - One-dimensional array.
'                alngTarget() - One-dimensional array that will hold the
'                    returned slice.
'                lngStart - [Optional] - Start index of slice to be
'                    returned.  Default = 0
'                lngCount - [Optional] - Number of elements in returned
'                    slice.    Default = -1
'                blnExclude - [Optional] - If true the result is inverted;
'                    returned is the source array minus the specified
'                    sub-array.  Default = FALSE
'
' Example:       The source is an array of 100 long integers.  We want to
'                start at element 50 and extract 3 elements.  Do not
'                exclude any.
'
'                If ExtractFromLongArray(alngSource(), alngTarget(), 50, 3, False) then
'                    ...
'                End If
'
' Returns:       TRUE - Successful finish
'                FALSE - Target array is empty
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Nov-2001  Ratmoler HAMSTAK  hamstak@hotmail.com
'              http://www.xbeat.net/vbspeed/c_SliceLng.htm
' 01-Apr-2007  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables and documented
' ***************************************************************************
Public Function ExtractFromLongArray(ByRef alngSource() As Long, _
                                     ByRef alngTarget() As Long, _
                            Optional ByVal lngStart As Long = 0, _
                            Optional ByVal lngCount As Long = -1, _
                            Optional ByVal blnExclude As Boolean = False) As Boolean

    Dim lngSrcLow   As Long
    Dim lngSrcHigh  As Long
    Dim lngDestLow  As Long
    Dim lngDestHigh As Long
    Dim lngOffSet   As Long
    Dim lngIndex    As Long

    On Error GoTo ExtractFromLongArray_Error

    ' Capture array parameters
    lngSrcLow = LBound(alngSource)
    lngSrcHigh = UBound(alngSource)
    lngDestLow = lngSrcLow

    ' test for invalid start position
    If lngStart < lngSrcLow Then
        lngStart = lngSrcLow

    ElseIf lngStart > lngSrcHigh Then

        If blnExclude Then
            blnExclude = False
            lngStart = lngSrcLow
            lngCount = lngSrcHigh - lngSrcLow + 1
        Else
            Erase alngTarget()
            Exit Function
        End If
    End If

    ' Test for invalid number of return elements
    If lngCount < 0 Then
        lngCount = lngSrcHigh - lngStart + 1

    ElseIf lngCount = 0 Then

        If blnExclude Then
            blnExclude = False
            lngStart = lngSrcLow
            lngCount = lngSrcHigh - lngStart + 1
        Else
            Erase alngTarget()
            Exit Function
        End If
    End If

    ' Test to verify the start position and number
    ' of returns falls within the range of this array.
    If lngStart + lngCount > lngSrcHigh Then

        If lngStart = lngSrcLow Then

           If blnExclude Then
               Erase alngTarget()
               Exit Function
           Else
               lngCount = lngSrcHigh - lngSrcLow + 1
           End If
        Else
            lngCount = lngSrcHigh - lngStart + 1
        End If
    End If

    If blnExclude Then

        lngDestHigh = lngSrcHigh - lngCount

        ReDim alngTarget(lngDestLow To lngDestHigh)

        If lngStart > lngSrcLow Then

            For lngIndex = lngSrcLow To lngStart - 1
                alngTarget(lngIndex) = alngSource(lngIndex)
            Next lngIndex

        End If

        If (lngStart + lngCount - 1) < lngSrcHigh Then

            For lngIndex = lngStart To lngDestHigh
                alngTarget(lngIndex) = alngSource(lngIndex + lngCount)
            Next lngIndex

        End If

    Else
        ' Specified segment (lngStart, lngCount) to be returned
        lngOffSet = lngStart - lngSrcLow
        lngDestHigh = lngSrcLow + lngCount - 1

        If lngDestHigh > lngSrcHigh Or lngDestHigh < lngDestLow Then
            lngDestHigh = lngSrcHigh - lngOffSet
        End If

        ReDim alngTarget(lngDestLow To lngDestHigh)

        For lngIndex = lngDestLow To lngDestHigh
            alngTarget(lngIndex) = alngSource(lngIndex + lngOffSet)
        Next lngIndex

    End If

    ExtractFromLongArray = True

ExtractFromLongArray_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ExtractFromLongArray_Error:
    ExtractFromLongArray = False
    Resume ExtractFromLongArray_CleanUp

End Function

' ***************************************************************************
' Routine:       GetComplement
'
' Description:   This routine will determine the inverse value of a byte
'                value (0-255).  Using bitwise NOT sometimes generates
'                the direct opposite (15 to -15) because of bit flipping.
'                What we need is a positive representation of the byte.
'
'                      Input           Complement
'                Ex:   15         -->  240       (decimal value)
'                      00001111   -->  11110000  (binary format)
'
' Parameters:    bytData - value to be evaluated
'
' Returns:       Inverse value of an ASCII decimal value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function GetComplement(ByVal bytData As Byte) As Byte

    GetComplement = CByte((Not bytData) And &HFF&)

End Function

' ***************************************************************************
' Routine:       GetHiLoWord
'
' Description:   Get the low and high word of a long integer.
'
' Parameters:    lngValue - number to be evaluated
'                LOWORD - Low order of a word to be returned
'                HIWORD - High order of a word to be returned
'
' Returns:       Low and high order of a word (long integer)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2001  Donald Lessau donald@xbeat.net
'              http://www.xbeat.net/vbspeed/c_HiWord.html
'              http://www.xbeat.net/vbspeed/c_LoWord.html
' ***************************************************************************
Public Function GetHiLoWord(ByVal lngValue As Long, _
                            ByRef LOWORD As Long, _
                            ByRef HIWORD As Long) As Boolean

    On Error GoTo GetHiLoWord_Error

    ' This is the LOWORD of the input value.
    If lngValue And &H8000& Then
        LOWORD = lngValue Or &HFFFF0000
    Else
        LOWORD = lngValue And &HFFFF&
    End If

    ' This is the HIWORD of the input value.
    HIWORD = (lngValue And &HFFFF0000) \ &H10000

    ' Successful finish
    GetHiLoWord = True
    Exit Function

GetHiLoWord_Error:
    Err.Clear
    LOWORD = 0
    HIWORD = 0
    GetHiLoWord = False

End Function

' ***************************************************************************
' Routine:       GetHiLOBYTE
'
' Description:   Obtain the low and high byte of an integer.
'
' Parameters:    intValue - number to be evaluated
'                LOBYTE - Low order of an integer to be returned
'                HIBYTE - High order of an integer to be returned
'
' Returns:       Low and high order of an integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Dec-2001  Donald Leassu donald@xbeat.net
'              http://VBSpeed/xbeat.net/vbspeed/c_HiByte.html
'              http://VBSpeed/xbeat.net/vbspeed/c_LoByte.html
' ***************************************************************************
Public Function GetHiLoByte(ByVal intValue As Integer, _
                            ByRef LoByte As Integer, _
                            ByRef HiByte As Integer) As Boolean

    On Error GoTo GetHiLOBYTE_Error

    ' This is the LOBYTE of the input value.
    LoByte = intValue And &HFF

    ' This is the HIBYTE of the input value.
    HiByte = (intValue \ &HFF00&) \ &H100

    ' Successful finish
    GetHiLoByte = True
    Exit Function

GetHiLOBYTE_Error:
    Err.Clear
    LoByte = 0
    HiByte = 0
    GetHiLoByte = False

End Function

' ***************************************************************************
' Routine:       HexToBinary
'
' Description:   Convert a hexadecimal value to a binary string.
'                Please note that it is faster to load an output string
'                with spaces and then insert the data.  Constantly
'                appending data will cause the system to have to
'                reallocate space, re-evaluate new string size and
'                possibly designate a new internal address to hold this
'                data.  All of which takes extra cycles and time.
'
' Input:         28a2f98
' Output:        00000010100010100010111110011000
'
' Parameters:    strHex - Hex value to be converted to binary
'
' Returns:       Binary string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function HexToBinary(ByVal strHex As String) As String

    Dim lngPos     As Long
    Dim lngPointer As Long
    Dim strBinary  As String

    On Error GoTo HexToBinary_Error

    HexToBinary = vbNullString   ' Empty return variables

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        GoTo HexToBinary_CleanUp
    End If

    strBinary = Space$(160)   ' Preload with spaces
    lngPointer = 1            ' Set output pointer to first position

    ' Parse hex string and convert
    ' one character at a time
    For lngPos = 1 To Len(strHex)

        Select Case UCase$(Mid$(strHex, lngPos, 1))
               Case "0": Mid$(strBinary, lngPointer, 4) = "0000"
               Case "1": Mid$(strBinary, lngPointer, 4) = "0001"
               Case "2": Mid$(strBinary, lngPointer, 4) = "0010"
               Case "3": Mid$(strBinary, lngPointer, 4) = "0011"
               Case "4": Mid$(strBinary, lngPointer, 4) = "0100"
               Case "5": Mid$(strBinary, lngPointer, 4) = "0101"
               Case "6": Mid$(strBinary, lngPointer, 4) = "0110"
               Case "7": Mid$(strBinary, lngPointer, 4) = "0111"
               Case "8": Mid$(strBinary, lngPointer, 4) = "1000"
               Case "9": Mid$(strBinary, lngPointer, 4) = "1001"
               Case "A": Mid$(strBinary, lngPointer, 4) = "1010"
               Case "B": Mid$(strBinary, lngPointer, 4) = "1011"
               Case "C": Mid$(strBinary, lngPointer, 4) = "1100"
               Case "D": Mid$(strBinary, lngPointer, 4) = "1101"
               Case "E": Mid$(strBinary, lngPointer, 4) = "1110"
               Case "F": Mid$(strBinary, lngPointer, 4) = "1111"
        End Select

        lngPointer = lngPointer + 4   ' Increment output pointer

    Next lngPos

    strBinary = Trim$(strBinary)  ' Remove excess blank spaces

HexToBinary_CleanUp:
    HexToBinary = strBinary   ' Return new value
    On Error GoTo 0           ' Nullify this error trap
    Exit Function

HexToBinary_Error:
    Err.Clear                   ' Remove error code
    gblnStopProcessing = True   ' Set flag to stop processing
    strBinary = vbNullString    ' Return new value
    Resume HexToBinary_CleanUp

End Function

' ***************************************************************************
' Routine:       HexToByteArray
'
' Description:   Convert a Hex string to a byte array
'
' Parameters:    strHex - Hex data to be converted
'
' Returns:       Byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Removed optional parameter
' ***************************************************************************
Public Function HexToByteArray(ByVal strHex As String) As Byte()

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim abytData() As Byte

    Const ROUTINE_NAME As String = "HexToByteArray"

    On Error GoTo HexToByteArray_Error

    Erase abytData()          ' Always starat with empty arrays
    lngLength = Len(strHex)   ' Capture length of hex string

    If lngLength = 0 Then
        InfoMsg "Invalid hex string length for conversion." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo HexToByteArray_CleanUp
    End If

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        InfoMsg "Invalid hex string for conversion." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo HexToByteArray_CleanUp
    End If

    lngPointer = 1  ' string pointer

    ' String must be divisable by 2
    If lngLength Mod 2 = 0 Then

        ReDim abytData(lngLength \ 2)  ' resize output array

        ' start converting data string two
        ' characters at a time to an ASCII
        ' decimal value
        For lngIndex = 0 To UBound(abytData) - 1

            abytData(lngIndex) = CByte("&H" & Mid$(strHex, lngPointer, 2))
            lngPointer = lngPointer + 2

        Next lngIndex

        ReDim Preserve abytData(lngIndex - 1)  ' resize to actual size

    Else
        ReDim abytData(0)
    End If

HexToByteArray_CleanUp:
    HexToByteArray = abytData()

    Erase abytData()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

HexToByteArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ReDim abytData(0)
    gblnStopProcessing = True
    Resume HexToByteArray_CleanUp

End Function

' ***************************************************************************
' Routine:       HexToLong
'
' Description:   Convert a Hex string to a long integer.  Assumes the
'                output value will be a long integer.
'
' Parameters:    strHex - Hex data to be converted
'
' Returns:       Long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'
'              The problem arose when converting the hex data back to a long
'              integer (4 byte storate area).
'
'                    Ex:  &HA780 -> -22656  (Bad)
'
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'
'                    Ex:  &HA780& -> 42880  (Good)
'
'              Karl Peterson comments:
'              "The trailing ampersand tells VB it's a Long value (4 byte),
'              even for values that could be held in an Integer (2 byte).
'              Basically, if it's &h8000-&hFFFF, use the ampersand (actually,
'              it's never a bad idea), assuming you really do want to store
'              it in a 4-byte variable. Otherwise, VB strives to confuse, by
'              twiddling the sign bit for those higher values."
'              http://bytes.com/topic/visual-basic-net/answers/387151-ampersands-hexadecimal-numbers-whats-rule
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine.
' ***************************************************************************
Public Function HexToLong(ByVal strHex As String) As Long

    Dim lngValue As Long

    Const ROUTINE_NAME As String = "HexToLong"

    On Error GoTo HexToLong_Error

    HexToLong = 0&

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        GoTo HexToLong_CleanUp
    End If

    strHex = Right$(String$(8, "0") & strHex, 8)
    lngValue = Val("&H" & strHex)

    ' See flowerbox above for detailed explanation
    If (lngValue >= CLng(&H8000)) And _
       (lngValue <= CLng(&HFFFF)) Then

        lngValue = Val("&H" & strHex & "&")
    End If

    HexToLong = lngValue   ' Return long integer

HexToLong_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Function

HexToLong_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume HexToLong_CleanUp

End Function

' ***************************************************************************
' Routine:       HexToNumber
'
' Description:   Convert a Hex string to a long integer
'
' Parameters:    strHex - Hex data to be converted
'                lngRetDataType - Optional - Determines data type being passed.
'                              Data types are 0-eLong, 1-eShort, 2-eByte
'                              Default = eLong
'                blnReturnAsLong - Optional - Store return number as four
'                              bytes rather than their converted data
'                              type (short integer or byte)
'
' Returns:       Correct value and type
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2006  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 28-DEC-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for notifying me that he had a
'              problem performing a hash with a specific file.  The hashed
'              results were not the same as with other hashing utilities.
'              I found that when a hex value is calculated from a number to
'              hex, a series of letters are created (hex).
'
'                    Ex:  File size = 5360 bytes -> 42880 bits -> A780 hex
'
'              The problem arose when converting the hex data back to a long
'              integer (4 byte storate area).
'
'                    Ex:  &HA780 -> -22656  (Bad)
'
'              By appending a trailing ampersand to the hex value the
'              conversion is corrected.
'
'                    Ex:  &HA780& -> 42880  (Good)
'
'              Karl Peterson comments:
'              "The trailing ampersand tells VB it's a Long value (4 byte),
'              even for values that could be held in an Integer (2 byte).
'              Basically, if it's &h8000-&hFFFF, use the ampersand (actually,
'              it's never a bad idea), assuming you really do want to store
'              it in a 4-byte variable. Otherwise, VB strives to confuse, by
'              twiddling the sign bit for those higher values."
'              http://bytes.com/topic/visual-basic-net/answers/387151-ampersands-hexadecimal-numbers-whats-rule
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine.
' ***************************************************************************
Public Function HexToNumber(ByVal strHex As String, _
                   Optional ByRef lngRetDataType As enumDataType = eUnknown, _
                   Optional ByVal blnReturnAsLong As Boolean = False) As Variant

    Dim bytValue As Byte
    Dim intValue As Integer
    Dim lngValue As Long

    Const ROUTINE_NAME As String = "HexToNumber"

    On Error GoTo HexToNumber_Error

    HexToNumber = Empty

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        GoTo HexToNumber_CleanUp
    End If

    ' Verify data type is within parameters
    If lngRetDataType = eUnknown Then
        If Not EvaluateNumber(Val("&H" & strHex), lngRetDataType) Then
            GoTo HexToNumber_CleanUp
        End If
    End If

    ' Determine correct return length
    Select Case lngRetDataType

           Case eLong
                strHex = Right$(String$(8, "0") & strHex, 8)
                lngValue = Val("&H" & strHex)

                ' See flowerbox above for detailed explanation
                If (lngValue >= CLng(&H8000)) And _
                   (lngValue <= CLng(&HFFFF)) Then

                    lngValue = Val("&H" & strHex & "&")
                End If

                HexToNumber = lngValue

           Case eShort
                strHex = Right$(String$(4, "0") & strHex, 4)
                intValue = CInt("&H" & strHex)

                ' See flowerbox above for detailed explanation
                If blnReturnAsLong Then
                    If (intValue >= CInt(&H8000)) And _
                       (intValue <= CInt(&HFFFF)) Then

                        HexToNumber = CLng(Val("&H" & strHex & "&"))   ' Store value as 4 bytes
                    Else
                        HexToNumber = CLng(intValue)   ' Store value as 4 bytes
                    End If
                Else
                    HexToNumber = intValue             ' Store value as 2 bytes
                End If

           Case eByte
                strHex = Right$(String$(2, "0") & strHex, 2)
                bytValue = Val("&H" & strHex)

                If blnReturnAsLong Then
                    HexToNumber = CLng(bytValue)   ' Store value as 4 bytes
                Else
                    HexToNumber = bytValue         ' Store as 1 byte (0-255)
                End If
    End Select

HexToNumber_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

HexToNumber_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume HexToNumber_CleanUp

End Function

' ***************************************************************************
' Routine:       HexArrayToByteArray
'
' Description:   Convert a hex string, stored in a byte array, into a normal
'                string of data, also stored in a byte array.  One character
'                per byte.
'
' Parameters:    abytData() - Array of hex data in byte format
'
' Returns:       abytData() array filled with normal string data in byte format
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub HexArrayToByteArray(ByRef abytData() As Byte)

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngSize    As Long
    Dim lngPointer As Long
    Dim abytTemp() As Byte

    lngPointer = 0                  ' index pointer for output string
    lngLength = UBound(abytData)    ' capture length of incoming data
    lngSize = lngLength \ 2         ' calc output length

    Erase abytTemp()          ' Always start with empty arrays
    ReDim abytTemp(lngSize)   ' Resize the output data array

    ' Convert hex data to string data
    ' and store in a byte array
    For lngIndex = 0 To lngSize - 1
        abytTemp(lngIndex) = CByte("&H" & Chr$(abytData(lngPointer)) & _
                                          Chr$(abytData(lngPointer + 1)))
        lngPointer = lngPointer + 2
    Next lngIndex

    ReDim Preserve abytTemp(lngSize)   ' resize to amount actually used
    lngLength = UBound(abytTemp)       ' capture size of output array

    Erase abytData()                   ' empty original array
    ReDim abytData(lngLength)          ' resize to hold temp array contents

    ' Transfer data to original array
    CopyMemory abytData(0), abytTemp(0), lngLength

    Erase abytTemp()   ' empty arrays when not in use

End Sub

' ***************************************************************************
' Routine:       ByteArrayToHex1CharArray
'
' Description:   Convert a normal string, stored in a byte array, into
'                a hex array. Separate hex characters prior to making
'                a byte value.
'
' Parameters:    abytData() - Array of byte data
'
' Returns:       Hex array in byte format. One character per element.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ByteArrayToHex1CharArray(ByRef abytData() As Byte) As Byte()

    Dim strHex    As String
    Dim lngIdx    As Long
    Dim lngIndex  As Long
    Dim lngLength As Long
    Dim abytHex() As Byte

    lngIdx = 0                         ' Starting output index
    lngLength = UBound(abytData)       ' capture length of incoming data
    ReDim abytHex(lngLength * 2 + 2)   ' Temp size hex array with buffer

    ' Convert byte array to hex array
    For lngIndex = 0 To lngLength

        strHex = Right$("0" & Hex$(abytData(lngIndex)), 2)
        abytHex(lngIdx) = CByte(Asc(Left$(strHex, 1)))
        abytHex(lngIdx + 1) = CByte(Asc(Right$(strHex, 1)))
        lngIdx = lngIdx + 2   ' Increment output index

    Next lngIndex

    ReDim Preserve abytHex(lngIdx - 1)     ' Adjust to what was used
    ByteArrayToHex1CharArray = abytHex()   ' Return hex byte array

End Function

' ***************************************************************************
' Routine:       ByteArrayToHex2CharArray
'
' Description:   Convert a normal string, stored in a byte array, into
'                a hex array.
'
' Parameters:    abytData() - Array of byte data
'
' Returns:       Hex array in string format. Two characters per element.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function ByteArrayToHex2CharArray(ByRef abytData() As Byte) As String()

    Dim lngIndex  As Long
    Dim lngLength As Long
    Dim astrHex() As String

    lngLength = UBound(abytData)   ' capture length of incoming data
    ReDim astrHex(lngLength)       ' Size hex array

    ' Convert byte array to hex array
    For lngIndex = 0 To lngLength
        astrHex(lngIndex) = Right$("0" & Hex$(abytData(lngIndex)), 2)
    Next lngIndex

    ByteArrayToHex2CharArray = astrHex()   ' Return hex array

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
' Routine:       IsBinaryData
'
' Description:   Parses a string of data to determine if it is in
'                binary format.  Must be zero or one.
'
' Parameters:    strData - String of data to be evaluated
'
' Returns:       TRUE  - Data string is in binary format
'                FALSE - Not in binary format
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
Public Function IsBinaryData(ByRef strData As String) As Boolean

    Dim lngIndex  As Long
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "IsBinaryData"
    Const BINARY_DATA  As String = "01"

    IsBinaryData = True   ' Preset to TRUE

    strData = Replace(strData, " ", "")   ' Remove all blank spaces
    strData = Replace(strData, "+", "")   ' Remove all plus signs
    strData = Replace(strData, "-", "")   ' Remove all dashes
    strData = Replace(strData, ".", "")   ' Remove all periods
    strData = Replace(strData, "*", "")   ' Remove all asteriks
    strData = TrimStr(strData)            ' Remove unwanted leading/trailings chars
    lngLength = Len(strData)              ' Capture length of data string

    ' Must be in multiples of 4
    If lngLength Mod 4 <> 0 Then
        strData = String$(4 - (Len(strData) Mod 4), "0") & strData
    End If

    strData = TrimStr(strData)   ' Remove unwanted leading/trailings chars
    lngLength = Len(strData)     ' Capture length of data string

    If lngLength > 0 Then

        ' Parse data string to verify
        ' each character is valid
        For lngIndex = 1 To lngLength

            If InStr(1, BINARY_DATA, Mid$(strData, lngIndex, 1)) = 0 Then
                InfoMsg "Invalid character [ " & Mid$(strData, lngIndex, 1) & _
                        " ] found in binary data." & vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                IsBinaryData = False   ' Found invalid character
                Exit For
            End If

        Next lngIndex
    Else
        InfoMsg "Incoming data string is empty." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        IsBinaryData = False
    End If

End Function

' ***************************************************************************
' Routine:       IsNumericData
'
' Description:   Parses a string of data to determine if it is in
'                numeric format.
'
' Parameters:    strData - String of data to be evaluated
'
' Returns:       TRUE  - Data string is in numeric format
'                FALSE - Not in numeric format
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
Public Function IsNumericData(ByRef strData As String) As Boolean

    Dim lngIndex  As Long
    Dim lngLength As Long

    Const ROUTINE_NAME As String = "IsNumericData"
    Const NUMERIC_DATA As String = "0123456789-+"

    IsNumericData = True   ' Preset to TRUE

    strData = Replace(strData, " ", "")   ' Remove all blank spaces
    strData = Replace(strData, "*", "")   ' Remove all asteriks
    strData = Replace(strData, ",", "")   ' Remove all commas
    strData = TrimStr(strData)            ' Remove unwanted leading/trailings chars
    lngLength = Len(strData)              ' Capture length of data string

    If lngLength > 0 Then

        ' No decimals allowed
        If InStr(1, strData, ".") > 0 Then
            InfoMsg "No decimals allowed.  Whole numbers only." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
            IsNumericData = False
        End If

        ' Parse data string to verify
        ' each character is valid
        For lngIndex = 1 To lngLength

            If InStr(1, NUMERIC_DATA, Mid$(strData, lngIndex, 1)) = 0 Then
                InfoMsg "Invalid character [ " & Mid$(strData, lngIndex, 1) & _
                        " ] found in numeric data." & vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
                IsNumericData = False   ' Found invalid character
                Exit For
            End If

        Next lngIndex

    Else
        InfoMsg "Incoming data string is empty." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        IsNumericData = False
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
' Routine:       LongArrayToByteArray
'
' Description:   Convert a Long array to a byte array.
'
' Parameters:    alngData() - Array to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function LongArrayToByteArray(ByRef alngData() As Long) As Byte()

    Dim lngMax     As Long
    Dim lngIndex   As Long
    Dim lngIdx     As Long
    Dim abytData() As Byte

    Const ROUTINE_NAME As String = "LongArrayToByteArray"

    On Error GoTo LongArrayToByteArray_Error

    Erase abytData()    ' Start with an empty array

    lngIdx = 0                   ' Byte array index
    lngMax = UBound(alngData)    ' capture incoming array size
    ReDim abytData(lngMax * 4)   ' Resize byte array accordingly

    For lngIndex = 0 To lngMax - 1

        CopyMemory abytData(lngIdx), alngData(lngIndex), 4&  ' Convert long to bytes
        lngIdx = lngIdx + 4                                  ' increment byte array index

    Next lngIndex

    ReDim Preserve abytData(lngIdx - 1)   ' Resize byte array to exact size
    LongArrayToByteArray = abytData()     ' Return byte array

LongArrayToByteArray_CleanUp:
    Erase abytData()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LongArrayToByteArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LongArrayToByteArray_CleanUp

End Function

' ***************************************************************************
' Routine:       LongArrayToString
'
' Description:   Convert a Long array to a string.
'
' Parameters:    alngData() - Long array to be converted
'
' Returns:       string data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a logic bug in formatting string length
' ***************************************************************************
Public Function LongArrayToString(ByRef alngData() As Long) As String

    Dim lngIndex         As Long
    Dim lngPointer       As Long
    Dim strOutput        As String
    Dim strTemp          As String
    Dim abytData(0 To 3) As Byte

    Const ROUTINE_NAME As String = "LongArrayToString"

    On Error GoTo LongArrayToString_Error

    lngPointer = 1
    strOutput = Space$(UBound(alngData) * 4)

    For lngIndex = 0 To UBound(alngData) - 1

        Erase abytData()  ' Empty receiving array
        strTemp = vbNullString      ' Empty temp hold area

        CopyMemory abytData(0), alngData(lngIndex), 4&       ' copy 1 long to byte array
        strTemp = ByteArrayToString(abytData())              ' Convert byte array to string
        Mid$(strOutput, lngPointer, Len(strTemp)) = strTemp  ' Insert string into output string
        lngPointer = lngPointer + Len(strTemp)               ' Update pointer in output string

    Next lngIndex

    LongArrayToString = TrimStr(strOutput)    ' Remove trailing blanks

LongArrayToString_CleanUp:
    Erase abytData()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LongArrayToString_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LongArrayToString_CleanUp

End Function

' ***************************************************************************
' Routine:       LongToByteArray
'
' Description:   Convert a long integer to a byte array.  Personally, I
'                prefer this version.
'
' Parameters:    lngValue - Number to be converted
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function LongToByteArray(ByVal lngValue As Long) As Byte()

    Dim abytData(0 To 3) As Byte

    Const ROUTINE_NAME As String = "LongToByteArray"

    On Error GoTo LongToByteArray_Error

    Erase abytData()                       ' Start with an empty array
    CopyMemory abytData(0), lngValue, 4&   ' Convert long to bytes
    LongToByteArray = abytData()           ' Return the byte array

LongToByteArray_CleanUp:
    Erase abytData()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LongToByteArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LongToByteArray_CleanUp

End Function

' ***************************************************************************
' Routine:       LongToDouble
'
' Description:   Convert a long integer to a double precision number.
'                Returns a decimal position of 14 places. This was
'                chosen so that exponentials would not be returned.
'                If you are working with large numbers then your return
'                values will probably be in string format so as to get
'                the full value of your calculations.
'
'                    Ex: 2147483647  ->  -2.32830732662872E-10
'
'                Desired return value is within this range:
'
'                    -0.99999999999999 to 0.99999999999999
'
'                    Ex: 2147483647  ->  -0.00000000023283
'
'                By changing the return value from Double to String and
'                Uncommenting out the line of code indicated below, you
'                will get an extended return value.
'
'                    Ex: 2147483647  ->  -0.000000000232830732662872
'
' Parameters:    lngValue - Long integer to be converted
'
' Returns:       Double precision value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function LongToDouble(ByVal lngValue As Long) As Double

    Dim dblTemp As Double

    If lngValue < 0 Then
        dblTemp = lngValue + GB_4
    Else
        dblTemp = lngValue
    End If

    LongToDouble = FormatNumber(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#), 14)

    ' Uncomment to create a string return value
    'LongToDouble = CStr(CDec(DBL_LOW + ((dblTemp * DBL_HIGH) - 1#)))

End Function

' ***************************************************************************
' Routine:       LongToHex
'
' Description:   Convert a Long to a Hex string. Handles both positive and
'                negative values.
'
' Parameters:    lngValue - Number to be converted to string data
'
' Returns:       Converted number
'                Ex:   2147483647 --> 7fffffff  (2gb)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine.  LongToHex() now calls NumberToHex().
' ***************************************************************************
Public Function LongToHex(ByVal lngValue As Long) As String

    LongToHex = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    LongToHex = NumberToHex(lngValue, eLong)   ' Return hex string

End Function

' ***************************************************************************
' Routine:       LongToHexRev
'
' Description:   Returns the reversed hexadecimal representation of a
'                specified Long (4 bytes = 8 hex chars, most significant
'                byte right).  Zeroes are right-padded (that's the
'                difference to VB's Hex$() function).
'
'                Ex:    LongToHexRev(2001)       --> "D1070000"
'                       LongToHexRev(&H87654321) --> "21436587"
'
' Parameters:    lngValue - long integer to be converted
'
' Returns:       New hex value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Sep-2001  Donald Leassu donald@xbeat.net
'              http://www.xbeat.net/vbspeed/c_LongToHexRev.htm
' 17-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables and documented
' ***************************************************************************
Public Function LongToHexRev(ByVal lngValue As Long) As String

    Dim lngTemp As Long
    Dim HIWORD  As Integer
    Dim LOWORD  As Integer

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        LongToHexRev = vbNullString
        Exit Function
    End If

    ' extract hiword and loword inline
    HIWORD = CInt((lngValue And &HFFFF0000) \ &H10000)

    If lngValue And &H8000& Then
        LOWORD = CInt(lngValue Or &HFFFF0000)
    Else
        LOWORD = CInt(lngValue And &HFFFF&)
    End If

    ' swap bytes
    HIWORD = ByteSwap(HIWORD)
    LOWORD = ByteSwap(LOWORD)

    ' swap words
    lngTemp = CLng((LOWORD * &H10000) Or (HIWORD And &HFFFF&))

    ' Convert to hex and uppercase
    LongToHexRev = UCase$(Right$("0000000" & Hex$(lngTemp), 8))

End Function

' ***************************************************************************
' Routine:       IntegerToLong
'
' Description:   Convert an integer to an unsigned long.
'
' Parameters:    intValue - Value to be converted to Double
'
' Returns:       New double value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-JUL-2004  Microsoft Knowledge Base Article - 189323
'              HOWTO: Convert Between Signed and Unsigned Numbers
'              http://support.microsoft.com/default.aspx?scid=kb;en-us;189323
' ***************************************************************************
Public Function IntegerToLong(intValue As Integer) As Long

    ' Convert a long integer to a double.  If the long is less
    ' than zero then add to it to make it a positive value.
    If intValue < 0 Then
        IntegerToLong = CLng(intValue + kB_64)
    Else
        IntegerToLong = CLng(intValue)
    End If

End Function

' ***************************************************************************
' Routine:       IntegerToByte
'
' Description:   The function takes an Integer containing a value in the range
'                of an unsigned Integer and returns a byte (0-255).
'
' Parameters:    lngValue - Number to be converted
'
' Returns:       Byte value (0 to 255)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function IntegerToByte(ByVal intValue As Integer) As Byte

    IntegerToByte = CByte(intValue And &HFF&)

End Function

' ***************************************************************************
' Routine:       LongToInteger
'
' Description:   The function takes a Long containing a value in the
'                range of an unsigned Integer and returns an Integer
'                (-32768 to 32767).
'
' Parameters:    lngValue - Number to be converted
'
' Returns:       Positive integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function LongToInteger(ByVal lngValue As Long) As Integer

    LongToInteger = CInt(lngValue And &H7FFF&)

End Function

' ***************************************************************************
' Routine:       CurrencyToLong
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
' Parameters:    curValue     - Value to be evaluated
'                lngLowOrder  - Highest value = 2,147,483,647
'                lngHighOrder - Highest value = 2,147,630
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
Public Sub CurrencyToLong(ByVal curValue As Currency, _
                          ByRef lngLowOrder As Long, _
                          ByRef lngHighOrder As Long)

    lngHighOrder = 0  ' Initialize return values
    lngLowOrder = 0

    ' If curValue is smaller than 4 Gb, the
    ' Do..Loop does not execute.  Between 4gb
    ' and 8gb, the Do..Loop executes once.
    ' Between 8gb and 12gb,  the Do..Loop
    ' executes twice.
    Do Until curValue < MAX_DWORD
        lngHighOrder = lngHighOrder + 1
        curValue = curValue - MAX_DWORD
    Loop

    ' Take the remainder and decide whether
    ' it needs the sign bit of the long to
    ' hold a value (&H80000000).  If it is
    ' smaller than 2,147,483,648 you can
    ' pass the value directly to the low
    ' order.  If not, you have to convert
    ' the value to the hex equivalent in
    ' signed 32bit.
    If curValue > MAX_LONG Then
        lngLowOrder = CLng((MAX_DWORD - (curValue - 1)) * -1)  ' Larger than 2gb
    Else
        lngLowOrder = CLng(curValue)  ' Less than 2gb
    End If

End Sub

' ***************************************************************************
' Routine:       LongToCurrency
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
' Parameters:    curValue     - Value to be calculated
'                lngLowOrder  - Highest value = 2,147,483,647
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
Public Function LongToCurrency(ByVal lngLowOrder As Long, _
                               ByVal lngHighOrder As Long) As Currency

    Dim curValue As Currency

    curValue = MAX_DWORD * CCur(lngHighOrder)

    If lngLowOrder < 0 Then
        LongToCurrency = curValue + CCur(MAX_DWORD + CCur(lngLowOrder + 1))
    Else
        LongToCurrency = curValue + CCur(lngLowOrder)
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedToLong
'
' Description:   This function takes a Double containing a value in the
'                range of an unsigned Long and returns a Long Integer.
'
' Parameters:    dblValue - Number to be converted
'
' Returns:       Positive long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewote routine
' ***************************************************************************
Public Function UnsignedToLong(ByVal dblValue As Double) As Long

    Do
        Do While dblValue > MAX_LONG
            dblValue = dblValue - GB_4
        Loop

        Do While dblValue < 0
            dblValue = dblValue + MAX_LONG
        Loop

    Loop Until (dblValue > 0) And (dblValue <= MAX_LONG)

    UnsignedToLong = CLng(dblValue)

End Function

' ***************************************************************************
' Routine:       MakeWord
'
' Description:   Combines two one-byte values to one 2-byte Word (aka Integer).
'
' Parameters:    LoByte - Low order of a byte
'                HiByte - High order of a byte
'
' Returns:       short integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 07-Dec-2000  Donald Leassu donald@xbeat.net
'              http://www.xbeat.net/vbspeed/c_MakeWord.htm
' ***************************************************************************
Public Function MakeWord(ByVal LoByte As Byte, _
                         ByVal HiByte As Byte) As Integer

    If HiByte And &H80& Then
        MakeWord = ((HiByte * &H100&) Or LoByte) Or &HFFFF0000
    Else
        MakeWord = (HiByte * &H100&) Or LoByte
    End If

End Function

' ***************************************************************************
' Routine:       MakeDWord
'
' Description:   Combines two 2-byte Words (aka Integers) to one 4-byte
'                DoubleWord (aka Long).  High word is coerced to Long to
'                allow it to overflow limits of multiplication which
'                shifts it left.
'
'                The trick to packing values is bit shifting. Because VB does
'                not provide bit shift operators to use, you need to do things
'                the old fashioned way; through multiplication. To make an
'                Integer the high word for a Long value, you need to multiply
'                it by &H10000. This has the effect of shifting the bit values
'                16-bits (2-bytes) to the left, making room for the low word
'                value you want to add.
'
'                Before you can add the low word value, however, you need to
'                make an adjustment. Remember that Visual Basic Integer types
'                are signed values, but the low word value needs to be
'                unsigned if you plan to add it to your high word value. To
'                make sure Visual Basic treats the low word as an unsigned
'                integer, you need to perform a bitwise "And" on the value
'                using &HFFFF& as a mask. In effect, this saves the value as
'                a Long integer with the high (signed) bit cleared but keeps
'                the original Integer's bit value preserved.
'
' Parameters:    LOWORD - Low order of a word
'                HIWORD - High order of a word
'
' Returns:       long integer
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 07-Dec-2000  Karl E. Peterson   http://www.mvps.org/vb
'              http://www.xbeat.net/vbspeed/c_MakeDWord.htm#MakeDWord06
' ***************************************************************************
Public Function MakeDWord(ByVal LOWORD As Integer, _
                          ByVal HIWORD As Integer) As Long

    MakeDWord = (CLng(HIWORD) * &H10000) Or (LOWORD And &HFFFF&)

End Function

' ***************************************************************************
' Routine:       PutWord
'
' Description:   Copy long integer from its memory location, by pointer,
'                and place into a byte array.
'
' Parameters:    lngValue - Number to be captured
'                abytData() - Return byte array
'                lngPointer - Starting position within byte array
'
' Returns:       byte array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub PutWord(ByVal lngValue As Long, _
                   ByRef abytData() As Byte, _
          Optional ByVal lngPointer As Long = 0)

    CopyMemory abytData(lngPointer), ByVal VarPtr(lngValue), Len(lngValue)

End Sub

' ***************************************************************************
' Routine:       StringToBinary
'
' Description:   Converts a string of data to a binary string.
'
' Parameters:    strData - String data to be converted
'
' Returns:       Binary data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-Oct-2001  Donald Leassu donald@xbeat.net
'              http://www.xbeat.net/vbspeed/c_StringToBit.htm
' 02-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modifed and added error trapping
' ***************************************************************************
Public Function StringToBinary(ByVal strData As String) As String

    Dim abytData()   As Byte
    Dim lngIndex     As Long
    Dim strBinary    As String
    Dim astrBinary() As String

    Const ROUTINE_NAME As String = "StringToBinary"

    On Error GoTo StringToBinary_Error

    Erase astrBinary()  ' Always start with empty arrays
    Erase abytData()

    astrBinary() = LoadBinaryArray(eBits8)   ' Load an array of 8-bit data.
    abytData() = StringToByteArray(strData)  ' Convert string data to a byte array
    strBinary = Space$(Len(strData) * 8)     ' Preload output string with spaces

    For lngIndex = 0 To Len(strData) - 1
        Mid$(strBinary, 1 + lngIndex * 8) = astrBinary(abytData(lngIndex))
    Next lngIndex

StringToBinary_CleanUp:
    StringToBinary = strBinary   ' Return binary data string

    Erase astrBinary()   ' Always empty arrays when not needed
    Erase abytData()
    On Error GoTo 0      ' Nullify this error trap
    Exit Function

StringToBinary_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    strBinary = vbNullString
    Resume StringToBinary_CleanUp

End Function

' ***************************************************************************
' Routine:       StringToHex
'
' Description:   Convert data into its 2-char hex format.  First determines
'                if incoming data is in string format or an array.
'
' Parameters:    lngReturnLength - Length of returned data
'                strIncomingData - incoming string data
'                blnReturnString - OPTIONAL - Return data in string or
'                       array format.
'                       TRUE - Return the data as a string (Default)
'                       FALSE - Return data in an array
'
' Returns:       Either string or array of data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-DEC-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 15-Aug-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function StringToHex(ByVal lngReturnLength As Long, _
                            ByVal strIncomingData As String, _
                   Optional ByVal blnReturnString As Boolean = True) As Variant

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strOutput  As String
    Dim astrData() As String

    Const ROUTINE_NAME As String = "StringToHex"

    On Error GoTo StringToHex_Error

    Erase astrData()                  ' Always start with empty arrays
    lngLength = Len(strIncomingData)  ' capture length of incoming data

    If blnReturnString Then

        lngPointer = 1                           ' initialize output pointer
        strOutput = Space$(lngLength * 2 + 4)    ' preload output string

        ' parse the data and convert each character to hex
        For lngIndex = 1 To lngLength

            Mid$(strOutput, lngPointer, 2) = Right$("0" & Hex$(Asc(Mid$(strIncomingData, lngIndex, 1))), 2)
            lngPointer = lngPointer + 2

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

        Next lngIndex

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo StringToHex_CleanUp
        End If

        strOutput = TrimStr(strOutput)                  ' Remove any trailing blanks
        strOutput = Left$(strOutput, lngReturnLength)   ' Capture data length
        StringToHex = strOutput                         ' return hex string

    Else  ' Return byte array

        ReDim astrData(lngReturnLength)   ' Size the output array

        ' parse the data and convert each character to hex
        For lngIndex = 0 To lngLength

            astrData(lngIndex) = Right$("0" & Hex$(Asc(Mid$(strIncomingData, lngIndex + 1, 1))), 2)

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

        Next lngIndex

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo StringToHex_CleanUp
        End If

        ReDim Preserve astrData(lngReturnLength)   ' Adjust array size
        StringToHex = astrData()                   ' return hex array

    End If

StringToHex_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        StringToHex = Empty
    End If

    Erase astrData()   ' Always empty arrays when not needed
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

StringToHex_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume StringToHex_CleanUp

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
' Routine:       HexStringToByteArray
'
' Description:   The message is converted from a hex string to a byte array.
'
' Parameters:    strData - hex string to be converted
'
' Returns:       Message data formatted in an array
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function HexStringToByteArray(ByVal strData As String) As Byte()

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim abytTemp() As Byte

    Erase abytTemp()   ' Always start with an empty array

    ' Test validity of incoming data
    If Not IsHexData(strData) Then

        InfoMsg "Invalid hex string." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & "HexStringToByteArray"

        ReDim abytTemp(0)
        GoTo HexStringToByteArray_CleanUp

    End If

    lngPointer = 1                 ' strData pointer position
    lngLength = Len(strData) \ 2   ' Calc input data length
    ReDim abytTemp(lngLength)      ' Size return array

    ' convert hex data to byte array
    For lngIndex = 0 To lngLength - 1

        abytTemp(lngIndex) = CByte("&H" & Mid$(strData, lngPointer, 2))
        lngPointer = lngPointer + 2

    Next lngIndex

HexStringToByteArray_CleanUp:
    HexStringToByteArray = abytTemp()   ' Return new byte array
    Erase abytTemp()                    ' Always empty arrays when not needed

End Function

' ***************************************************************************
' Routine:       SwapEndian
'
' Description:   Converts hex data from BIG_ENDIAN to LITTLE_ENDIAN or
'                LITTLE_ENDIAN to BIG_ENDIAN format.
'
' Parameters:    vntData - Hex data to be manipulated
'                lngRetLength - Optional - Desired length of the return
'                         string.  Default = 8
'
' Returns:       Hex string in new format
'
' Example:         LITTLE_ENDIAN  to   BIG_ENDIAN
'                   01234567            67452301
'
'                  BIG_ENDIAN     to   LITTLE_ENDIAN
'                   67452301            01234567
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 15-Jun-2015  Kenneth Ives  kenaso@tx.rr.com
'              Updated evaluation logic and documentation
' ***************************************************************************
Public Function SwapEndian(ByVal vntData As Variant, _
                  Optional ByVal lngRetLength As Long = 8) As String

    Dim strHex      As String
    Dim strData     As String
    Dim strOutput   As String
    Dim lngPosition As Long
    Dim lngPointer  As Long

    Const ROUTINE_NAME As String = "SwapEndian"

    ' No negative or zero lengths allowed
    If lngRetLength < 1 Then
        lngRetLength = 2
    End If

    ' Return length must be in multiples of 2
    If lngRetLength Mod 2 <> 0 Then
        lngRetLength = lngRetLength + 1
    End If

    SwapEndian = vbNullString
    strData = CStr(vntData)

    ' Is this hex data ?
    If IsHexData(strData) Then
        strHex = strData   ' Save data for swapping
    Else
        ' Is this numeric data ?
        If IsNumericData(strData) Then
            ' Convert numeric data to hex
            strHex = Right$(String$(lngRetLength, "0") & Hex$(Val(strData)), lngRetLength)
        Else
            ' Bad data entered
            InfoMsg "Cannot identify input data type to perform swap endian." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
            gblnStopProcessing = True
            Exit Function
        End If
    End If

    ' Verify the incoming hex string is the
    ' same as the return length including
    ' any leading zeroes.
    strOutput = Space$(lngRetLength * 2)   ' Preload output string
    strHex = Right$(String$(lngRetLength, "0") & strHex, lngRetLength)
    lngPointer = lngRetLength - 1          ' Initialize pointer

    ' Parse hex string capturing two characters
    ' at a time and inserting them into the
    ' output string
    For lngPosition = 1 To lngRetLength Step 2

        Mid$(strOutput, lngPosition, 2) = Mid$(strHex, lngPointer, 2)
        lngPointer = lngPointer - 2

    Next lngPosition

    ' Remove leading and trailing spaces
    ' before returning hex string
    SwapEndian = Trim$(strOutput)

End Function

' ***************************************************************************
' Routine:       SwapEndianLong
'
' Description:   Returns a Long with reversed byte order.
'
'                Example:        Input                  Output
'                          Numeric     Hex        Numeric       Hex
'                          19158648  &H1245678    2018915346  &H78563412
'
' Parameters:    lngValue - long integer to be converted
'
' Returns:       New long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Sep-2004  Mike Sutton   Mike.Sutton@btclick.com
'              http://www.xbeat.net/vbspeed/c_SwapEndian.htm
' 17-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables
' ***************************************************************************
Public Function SwapEndianLong(ByVal lngValue As Long) As Long

    If lngValue = 0 Then
        SwapEndianLong = 0
        Exit Function
    End If

    SwapEndianLong = (((lngValue And &HFF000000) \ &H1000000) And &HFF&) Or _
                     ((lngValue And &HFF0000) \ &H100&) Or _
                     ((lngValue And &HFF00&) * &H100&) Or _
                     ((lngValue And &H7F&) * &H1000000)

    If (lngValue And &H80&) Then
        SwapEndianLong = SwapEndianLong Or &H80000000
    End If

End Function

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.  I use this
'                for swapping strings, type structures, numbers with
'                decimal values, etc.
'
' Parameters:    vntValue1 - Incoming data to be swapped with Value2
'                vntValue2 - Incoming data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Sub SwapData(ByRef vntData1 As Variant, _
                    ByRef vntData2 As Variant)

    Dim vntHold As Variant

    vntHold = Empty   ' Start with empty variants (prevents crashing)

    vntHold = vntData1
    vntData1 = vntData2
    vntData2 = vntHold

    vntHold = Empty   ' Always empty variants (prevents crashing)

End Sub

' ***************************************************************************
' Routine:       SwapStrings
'
' Description:   Swap string data with each other.  This function was
'                written since BASIC stopped having its own SWAP function.
'                This is used for swapping strings using API CopyMemory()
'                function.  This is very fast and efficient.
'
' Parameters:    strData1 - Incoming data to be swapped with Data2
'                strData2 - Incoming data to be swapped with Data1
'
' Returns:       Swapped string data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 29-Dec-2002  John Korejwa  korejwa@tiac.net
'              Wrote routine
' 09-Apr-2015  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Sub SwapStrings(ByRef strData1 As String, _
                       ByRef strData2 As String)

    Dim lngPointer As Long

    CopyMemory lngPointer, ByVal VarPtr(strData1), 4&               ' Capture 1st string address pointer
    CopyMemory ByVal VarPtr(strData1), ByVal VarPtr(strData2), 4&   ' Copy 2nd string address pointer to 1st pointer
    CopyMemory ByVal VarPtr(strData2), lngPointer, 4&               ' Replace 2nd string address pointer with captured pointer

    CopyMemory lngPointer, 0&, 4&                                   ' Verify hold area is empty

End Sub

' ***************************************************************************
' Routine:       SwapLong
'
' Description:   I wrote this function since BASIC stopped having its own
'                SWAP function.  I use this to Swap data (byte, integer,
'                or long) with each other using a temp hold.
'
'                This routine works with byte, lnteger and long values.
'                Change the parameter data type accordingly.
'
' Note:          I went back to this process of performing a swap after
'                being reminded of "What happens if two values hold the
'                same memory space?".  The answer is undesired results.
'
' Parameters:    lngValue1 - data to be swapped with Value2
'                lngValue2 - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Jul-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routines
' ***************************************************************************
Public Sub SwapLong(ByRef lngValue1 As Long, _
                    ByRef lngValue2 As Long)

    ' Swap long lnteger values (-2,147,483,648 to 2,147,483,647)

    Dim lngHold As Long

    lngHold = lngValue1
    lngValue1 = lngValue2
    lngValue2 = lngHold

End Sub

Public Sub SwapInt(ByRef intValue1 As Integer, _
                   ByRef intValue2 As Integer)

    ' Swap short integer values (-32,768 to 32,767)

    Dim intHold As Integer

    intHold = intValue1
    intValue1 = intValue2
    intValue2 = intHold

End Sub

Public Sub SwapBytes(ByRef bytValue1 As Byte, _
                     ByRef bytValue2 As Byte)

    ' Swap byte values (0 to 255)

    Dim bytHold As Byte

    bytHold = bytValue1
    bytValue1 = bytValue2
    bytValue2 = bytHold

End Sub

Public Function ByteSwap(ByVal intValue As Integer) As Integer

    ' 10-Sep-2001  Donald Leassu donald@xbeat.net
    '              Function ByteSwap01
    '              http://www.xbeat.net/vbspeed/c_LongToHexRev.htm
    '
    ' Swap first byte of a short integer with the second byte
    ' (Short integer = 2 bytes)

    Dim LoByte As Byte
    Dim HiByte As Byte

    LoByte = CByte(intValue And &HFF)
    HiByte = CByte((intValue And &HFF00&) \ &H100)

    If LoByte And &H80 Then
        ByteSwap = CInt(((LoByte * &H100&) Or HiByte) Or &HFFFF0000)
    Else
        ByteSwap = CInt((LoByte * &H100) Or HiByte)
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedAddByte
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    intValue1 - Number 1 (Byte 0 to 255)
'                intValue2 - Number 2 (Byte 0 to 255)
'
' Returns:       Calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function UnsignedAddByte(ByVal intValue1 As Integer, _
                                ByVal intValue2 As Integer) As Byte

    UnsignedAddByte = CByte((intValue1 + intValue2) And &HFF&)

End Function

' ***************************************************************************
' Routine:       UnsignedAddInteger
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    lngValue1 - Number 1 (Short integer -32768 to 32767)
'                lngValue2 - Number 2 (Short integer -32768 to 32767)
'
' Returns:       Calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function UnsignedAddInteger(ByVal lngValue1 As Long, _
                                   ByVal lngValue2 As Long) As Integer

    UnsignedAddInteger = CInt((lngValue1 + lngValue2) And &H7FFF&)

End Function

' ***************************************************************************
' Routine:       UnsignedAdd
'
' Description:   Function to add two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    dblValue1 - Value of A
'                dblValue2 - Value of B
'
' Returns:       Calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 19-Dec-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names
' ***************************************************************************
Public Function UnsignedAdd(ByVal dblValue1 As Double, _
                            ByVal dblValue2 As Double) As Long

    Dim dblTemp As Double

    dblTemp = dblValue1 + dblValue2

    If dblTemp < MIN_LONG Then
        UnsignedAdd = CLng(GB_4 + dblTemp)
    Else
        If dblTemp > MAX_LONG Then
            UnsignedAdd = CLng(dblTemp - GB_4)
        Else
            UnsignedAdd = CLng(dblTemp)
        End If
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedDivide
'
' Description:   Divides the two (signed) Long parameters, treated as
'                unsigned long, and returns the result as a (signed)
'                long integer result.  Overflows are ignored.
'
' Parameters:    lngDividend - Number to be divided (Dividend)
'                lngDivisor - Number performing division (Divisor)
'
' Returns:       New calculated value (Quotient)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 01-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function UnsignedDivide(ByVal lngDividend As Long, _
                               ByVal lngDivisor As Long) As Long

    If lngDividend < 0 Then
        UnsignedDivide = CLng(Fix((GB_4 + lngDividend) / lngDivisor))
    Else
        UnsignedDivide = CLng(Fix(lngDividend / lngDivisor))
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedDivideDbl
'
' Description:   Divides a double value by a (signed) Long divisor, treated
'                as unsigned long, and returns the result as a Double of
'                long integer value.
'
' Parameters:    dblDividend - Number to be divided (Dividend)
'                lngDivisor - Number performing division (Divisor)
'
' Returns:       New calculated value (Quotient)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 01-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables
' ***************************************************************************
Public Function UnsignedDivideDbl(ByVal dblDividend As Double, _
                                  ByVal lngDivisor As Long) As Double

    If dblDividend < 0 Then
        UnsignedDivideDbl = Fix((GB_4 + dblDividend) / lngDivisor)
    Else
        UnsignedDivideDbl = Fix(dblDividend / lngDivisor)
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedMultiply
'
' Description:   Multiplies the two (signed) Long parameters, treated as
'                unsigned long, and returns the lowest 4 bytes of the 8 bytes
'                result as a (signed) Long result.  Overflows are ignored.
'
'                This function emulates the multiplication of two (unsigned
'                long) numbers, and is needed because the type Double has
'                only a 53 bits mantissa and the result of multiplying 2 Long
'                variables might need 64 bits to accurately represent the
'                result in some cases.
'
'                lngValue1   == ABCD  == A000 + B00 + C0 + D
'                lngValue2   == EFGH  == E000 + F00 + G0 + H
'
'         Note:  In the following, "ae" means the 2 bytes result of A*E,
'                "bg" of B*G, etc:
'
'                lngValue1 * lngValue2 == ae 000 000 +  ' discard, result is too high
'                       af 000  00 +     ' discard, result is too high
'                       ag 000   0 +     ' discard, result is too high
'                       ah 000     +     ' take lowest byte
'
'                       be  00 000 +     ' discard, result is too high
'                       bf  00  00 +     ' discard, result is too high
'                       bg  00   0 +     ' take lowest byte
'                       bh  00     +     ' take both bytes
'
'                       ce   0 000 +     ' discard, result is too high
'                       cf   0  00 +     ' take lowest byte
'                       cg   0   0 +     ' take both bytes
'                       ch   0     +     ' take both bytes
'
'                       de     000 +     ' take lowest byte
'                       df      00 +     ' take both bytes
'                       dg       0 +     ' take both bytes
'                       dh               ' take both bytes
'
' Parameters:   lngValue1 - Number that gets multiplied (Multiplicand)
'               lngValue2 - Number doing the multiplying (Multiplier)
'
' Returns:      New calculated value (Product)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 18-Apr-2005  Pablo Mariano Ronchi  pmronchi@yahoo.com.ar
'              Routine created
' 01-May-2005  Kenneth Ives  kenaso@tx.rr.com
'              Renamed variables
' ***************************************************************************
Public Function UnsignedMultiply(ByVal lngValue1 As Long, _
                                 ByVal lngValue2 As Long) As Long

    Dim AA      As Long
    Dim BB      As Long
    Dim CC      As Long
    Dim XX      As Long
    Dim YY      As Long
    Dim ZZ      As Long
    Dim R0      As Long
    Dim R1      As Long
    Dim R2      As Long
    Dim R3      As Long
    Dim dblTemp As Double

    Const K2_8  As Long = 256        ' 256 = (2^8)
    Const K2_16 As Long = 65536      ' 65536 = (2^16)
    Const K2_24 As Long = 16777216   ' 16777216 = (2^24)

    AA = (lngValue1 \ K2_16) Mod K2_8
    BB = (lngValue1 \ K2_8) Mod K2_8
    CC = lngValue1 Mod K2_8
    
    XX = (lngValue2 \ K2_16) Mod K2_8
    YY = (lngValue2 \ K2_8) Mod K2_8
    ZZ = lngValue2 Mod K2_8

    ' Get the 1st (lowest) byte of the result, R0:
    '       dh             'take both bytes
    R0 = CC * ZZ

    ' Get the 2nd byte of the result, R1, and add carry from R0:
    '       ch   0      +  'take both bytes
    '       dg        0    'take both bytes
    R1 = BB * ZZ + CC * YY + R0 \ K2_8

    ' Get the 3rd byte of the result, R2, and add carry from R1:
    '       bh  00      +  'take both bytes
    '       cg   0    0 +  'take both bytes
    '       df       00    'take both bytes
    R2 = AA * ZZ + BB * YY + CC * XX + R1 \ K2_8

    ' Get the 4th (highest) byte of the result, R3, and add carry from R2:
    '       ah 000      +  'take lowest byte
    '       bg  00    0 +  'take lowest byte
    '       cf   0   00 +  'take lowest byte
    '       de      000    'take lowest byte
    R3 = (((lngValue1 \ K2_24) * ZZ + AA * YY + BB * XX + CC * (lngValue2 \ K2_24)) Mod K2_8) + R2 \ K2_8
    dblTemp = CDbl(R3 Mod K2_8) * K2_24 + (R2 Mod K2_8) * K2_16 + (R1 Mod K2_8) * K2_8 + (R0 Mod K2_8)

    ' Now we have a 32 bit number (dblTemp) that can be processed
    ' without losing precision using the 53 bits mantissa of the
    ' Double Precision data type.
    If dblTemp < MIN_LONG Then
        UnsignedMultiply = CLng(GB_4 + dblTemp)
    Else
        If dblTemp > MAX_LONG Then
            UnsignedMultiply = CLng(dblTemp - GB_4)
        Else
            UnsignedMultiply = CLng(dblTemp)
        End If
    End If

End Function

' ***************************************************************************
' Routine:       UnsignedSubtract
'
' Description:   Function to subtract two unsigned numbers together as in C.
'                Overflows are ignored!
'
' Parameters:    lngValue1 - First value
'                lngValue2 - Second value is subtracted from the first
'
' Returns:       calculated value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function UnsignedSubtract(ByVal lngValue1 As Long, _
                                 ByVal lngValue2 As Long) As Long

    Dim lngCarry    As Long
    Dim lngTemp     As Long
    Dim lngIndex    As Long
    Dim abytTemp()  As Byte
    Dim abytData1() As Byte
    Dim abytData2() As Byte

    Const ROUTINE_NAME As String = "UnsignedSubtract"

    On Error GoTo UnsignedSubtract_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    lngTemp = 0&

    abytTemp() = LongToByteArray(lngTemp)
    abytData1() = LongToByteArray(lngValue1)
    abytData2() = LongToByteArray(lngValue2)

    For lngIndex = 0 To 3

        lngTemp = CLng(abytData1(lngIndex)) - CLng(abytData2(lngIndex)) - lngCarry

        If (lngTemp < 0) Then
            lngTemp = lngTemp + MAX_BYTE
            lngCarry = 1
        Else
            lngCarry = 0
        End If

        abytTemp(lngIndex) = lngTemp

    Next lngIndex

    lngTemp = ByteArrayToLong(abytTemp())

UnsignedSubtract_CleanUp:
    UnsignedSubtract = lngTemp

    Erase abytData1()  ' Always empty variants and arrays
    Erase abytData2()
    Erase abytTemp()

    On Error GoTo 0    ' Nullify this error trap
    Exit Function

UnsignedSubtract_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    lngTemp = 0
    gblnStopProcessing = True
    Resume UnsignedSubtract_CleanUp

End Function

' ***************************************************************************
' Routine:       w8Shift (Byte shifting)
'
' Description:   Shifts the bits to the right or left the specified number
'                of positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around on the opposite side.  Some
'                common languages like C/C++ or Java have an operator for
'                this job:  ">>" and "<<".
'
' Parameters:    bytValue    - numeric value to be evaluated
'                lngBitShift - number of bit positions
'                              Positive value = left shift
'                              Negative value = right shift
'
' Returns:       Reformatted value
'
'                Byte    Binary
' Original:        60   00111100
'   Left 3:       224   11100000
'  Right 3:         7   00000111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 01-Feb-2009  Kenneth Ives  kenaso@tx.rr.com
'              Corrected logic in testing constant arrays
' ***************************************************************************
Public Function w8Shift(ByVal bytValue As Byte, _
                        ByVal lngBitShift As Long) As Byte

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                       ' Nothing to do, return original value
           Case Is > 7:  bytValue = 0   ' return zero too many positive bit shift positions
           Case Is < -7: bytValue = 0   ' return zero too many negative bit shift positions

           ' Positive bit shift means shift left
           Case Is > 0
                If bytValue <> 0 Then
                    If lngBitShift = 7 Then

                        ' Test for non-zero
                        If bytValue And 1 Then
                            bytValue = &H80
                        Else
                            bytValue = 0
                        End If
                    Else
                        ' See if arrays are already loaded
                        If Not mblnLoaded Then
                            LoadShiftArrays
                        End If

                        bytValue = ((bytValue And mabytValue(7 - lngBitShift)) * mabytPower2(lngBitShift))
                    End If
                End If

           ' Negative bit shift means shift right
           Case Is < 0
                lngBitShift = Abs(lngBitShift)   ' Make bit shift value positive

                If bytValue <> 0 Then
                    If lngBitShift = 7 Then

                        ' Test for non-zero
                        If bytValue And &H80 Then
                            bytValue = 1
                        Else
                            bytValue = 0
                        End If
                    Else
                        If Not mblnLoaded Then
                            LoadShiftArrays   ' See if arrays are already loaded
                        End If

                        bytValue = bytValue \ mabytPower2(lngBitShift)
                    End If
                End If
    End Select

    w8Shift = bytValue   ' Return new value

End Function

' ***************************************************************************
' Routine:       w8Rotate (Byte rotation)
'
' Description:   Shifts the bits to the right or left the specified number
'                of positions and returns the new value.  Bits "falling off"
'                the edge will wrap around on the opposite side.  Some
'                common languages like C/C++ or Java have an operator for
'                this job:  ">>>" or "<<<".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
' Parameters:    bytValue    - numeric value to be evaluated
'                lngBitShift - number of bit positions
'                              Positive value = left rotation
'                              Negative value = right rotation
'
' Returns:       Reformatted value
'
'                Byte    Binary
' Original:        60   00111100
'   Left 3:       225   11100001
'  Right 3:       135   10000111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Apr-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 17-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed a bug in rotating right.
'              - Corrected erroneous documentation.
' ***************************************************************************
Public Function w8Rotate(ByVal bytValue As Byte, _
                         ByVal lngBitShift As Long) As Byte

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                       ' Nothing to do, return original value
           Case Is > 7:  bytValue = 0   ' return zero too many positive bit shift positions
           Case Is < -7: bytValue = 0   ' return zero too many negative bit shift positions

           ' Positive bit shift means shift left
           Case Is > 0
                bytValue = w8Shift(bytValue, lngBitShift) Or _
                           w8Shift(bytValue, -(8 - lngBitShift))

           ' Negative bit shift means shift right
           Case Is < 0
                bytValue = w8Shift(bytValue, lngBitShift) Or _
                           w8Shift(bytValue, (8 - Abs(lngBitShift)))
    End Select

    w8Rotate = bytValue   ' Return new value

End Function

' ***************************************************************************
' Routine:       w16Shift (Short Integer)
'
' Description:   Shifts the bits to the right or left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around.  The fill bits are zeroes on
'                the opposite side.  Some common languages like C/C++ or
'                this job: ">>" and "<<".
'
'                When viewing the binary output, remember the first position
'                is the sign bit.  Zero is positive and one is negative.
'
'                Verify your results with MS Calculator (CALC.EXE) that comes
'                with Windows.  Use the scientific view.
'
' Parameters:    lngValue     - Number to be manipulated
'                lngBitShift  - number of bit positions
'                               Positive value = left shift
'                               Negative value = right shift
'
' Returns:       Reformatted short integer.
'
'                Short Integer       Binary
' Original:           32767     0111111111111111
'   Left 5:             -32     1111111111100000
'  Right 5:            1023     0000001111111111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle short integers
' ***************************************************************************
Public Function w16Shift(ByVal lngValue As Long, _
                         ByVal lngBitShift As Long) As Integer

    Dim lngLoop    As Long
    Dim lngSignBit As Long

    ' Return original value if data is out of range
    If (lngValue < -32768) Or _
       (lngValue > 32767) Then

        w16Shift = lngValue
        Exit Function
    End If

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                        ' Nothing to do, return original value
           Case Is > 15:  lngValue = 0   ' return zero too many positive bit positions
           Case Is < -15: lngValue = 0   ' return zero too many negative bit positions

           ' Positive bit shift means shift left
           Case Is > 0
                ' Shift left by one bit but take in account of
                ' an overflow error from VB. Mask accordingly.
                If lngValue <> 0 Then
                    For lngLoop = 1 To lngBitShift

                        ' Calculate sign bit of result
                        lngSignBit = lngValue And &HC000

                        ' Clear Most Significant Bit (MSB), that
                        ' would be lost anyway, also clear sign bit
                        lngValue = (lngValue And &H3FFF) * 2

                        ' set or clear MSB
                        If lngSignBit And &H4000 Then
                            lngValue = lngValue Or &H8000
                        Else
                            lngValue = lngValue And &H7FFF
                        End If

                    Next lngLoop
                End If

           ' Negative bit shift means shift right
           Case Is < 0
                ' Shift left by one bit but take in account of
                ' an overflow error from VB. Mask accordingly.
                If lngValue <> 0 Then
                    For lngLoop = lngBitShift To -1

                        ' Calculate sign bit of result
                        lngSignBit = lngValue And &H8001

                        ' Clear Most Significant Bit (MSB), that
                        ' would be lost anyway, also clear sign bit
                        lngValue = (lngValue And &H7FFF) \ 2

                        ' set or clear the old sign bit
                        If lngSignBit And &H8000 Then
                            lngValue = lngValue Or &H4000
                        Else
                            lngValue = lngValue And &HBFFF
                        End If

                    Next lngLoop
                End If
    End Select

    w16Shift = CInt(lngValue)  ' Return new value

End Function

' ***************************************************************************
' Routine:       w16Rotate (Short Integer)
'
' Description:   Shifts the bits to the right or left the specified number
'                of positions and returns the new value.  Bits "falling off"
'                the edge will wrap around to the opposite side.  Some
'                common languages like C/C++ or Java have an operator for
'                this job: ">>>" or "<<<".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
'                When viewing the binary output, remember the first position
'                is the sign bit.  Zero is positive and one is negative.
'
'                Verify your results with MS Calculator (CALC.EXE) that comes
'                with Windows.  Use the scientific view.
'
' Parameters:    lngValue     - Number to be manipulated
'                lngBitShift - number of bit positions
'                              Positive value = left rotation
'                              Negative value = right rotation
'
' Returns:       Reformatted short integer.
'
'                Short Integer       Binary
' Original:           32767     0111111111111111
'   Left 5:             -17     1111111111101111
'  Right 5:           -1025     1111101111111111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 12-Oct-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle short integers
' ***************************************************************************
Public Function w16Rotate(ByVal lngValue As Long, _
                          ByVal lngBitShift As Long) As Long

    Dim lngLoop    As Long
    Dim lngSignBit As Long

    ' Return original value if data is out of range
    If (lngValue < -32768) Or _
       (lngValue > 32767) Then

        w16Rotate = lngValue
        Exit Function
    End If

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                        ' Nothing to do, return original value
           Case Is > 15:  lngValue = 0   ' return zero too many positive bit positions
           Case Is < -15: lngValue = 0   ' return zero too many negative bit positions

           ' Positive bit shift means rotate left
           Case Is > 0
                If lngValue <> 0 Then
                    For lngLoop = 1 To lngBitShift

                        ' remember the two most significant bits
                        lngSignBit = lngValue And &HC000

                        ' clear the bit and shift left by one position
                        lngValue = (lngValue And &H3FFF) * 2

                        ' if number was negative, then add 1
                        ' if bit 30 was set, then set the sign bit
                        lngValue = lngValue Or _
                                   ((lngSignBit < 0) And &H1&) Or _
                                   (CBool(lngSignBit And &H4000) And &H8000)

                    Next lngLoop
                End If

           ' Negative bit shift means rotate right
           Case Is < 0
                If lngValue <> 0 Then
                    For lngLoop = lngBitShift To -1

                        ' remember the sign bit and bit 0
                        lngSignBit = lngValue And &H8001

                        ' clear the bit and shift right by one position
                        lngValue = (lngValue And &H7FFE) \ 2

                        ' if number was negative, then reinsert the bit
                        ' if bit 0 was set, then set the sign bit
                        lngValue = lngValue Or _
                                   ((lngSignBit < 0) And &H4000) Or _
                                   (CBool(lngSignBit And 1) And &H8000)

                    Next lngLoop
                End If
    End Select

    w16Rotate = CInt(lngValue)   ' Return new value

End Function

' ***************************************************************************
' Routine:       w32Shift (Long Integer)
'
' Description:   Shifts bits to right or left a specified number of
'                positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around.  Fill bits are zeroes on
'                the opposite side.  Some common languages like C/C++ or
'                this job:  ">>" and "<<"
'
'                When viewing the binary output, remember the first position
'                is the sign bit.  Zero is positive and one is negative.
'
'                Verify your results with MS Calculator (CALC.EXE) that comes
'                with Windows.  Use the scientific view.
'
' Parameters:    lngValue    - Number to be manipulated
'                lngBitShift - number of shift positions
'                              Positive value = left shift
'                              Negative value = right shift
'
' Returns:       Reformatted value
'
'                   Number                Binary
' Original:       123456789   00000111010110111100110100010101
'   Left 5:      -344350048   11101011011110011010001010100000
'  Right 5:       3858024     00000000001110101101111001101000
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and documented
' ***************************************************************************
Public Function w32Shift(ByVal lngValue As Long, _
                         ByVal lngBitShift As Long) As Long

    Dim lngLoop    As Long
    Dim lngSignBit As Long

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                        ' Nothing to do, return original value
           Case Is > 31:  lngValue = 0   ' return zero too many positive bit positions
           Case Is < -31: lngValue = 0   ' return zero too many negative bit positions

           ' Positive bit shift means shift left
           Case Is > 0
                ' Shift left by one bit but take in account of
                ' an overflow error from VB. Mask accordingly.
                If lngValue <> 0 Then
                    For lngLoop = 1 To lngBitShift

                        ' Calculate sign bit of result
                        lngSignBit = lngValue And &HC0000000

                        ' Clear Most Significant Bit (MSB), that
                        ' would be lost anyway, also clear sign bit
                        lngValue = (lngValue And &H3FFFFFFF) * 2

                        ' set or clear MSB
                        If lngSignBit And &H40000000 Then
                            lngValue = lngValue Or &H80000000
                        Else
                            lngValue = lngValue And &H7FFFFFFF
                        End If

                    Next lngLoop
                End If

           ' Negative bit shift means shift right
           Case Is < 0
                ' Shift right by one bit but take in account of
                ' an overflow error from VB. Mask accordingly.
                If lngValue <> 0 Then
                    For lngLoop = lngBitShift To -1

                        ' Calculate sign bit of result
                        lngSignBit = lngValue And &H80000001

                        ' Clear Most Significant Bit (MSB), that
                        ' would be lost anyway, also clear sign bit
                        lngValue = (lngValue And &H7FFFFFFF) \ 2

                        ' set or clear the old sign bit
                        If lngSignBit And &H80000000 Then
                            lngValue = lngValue Or &H40000000
                        Else
                            lngValue = lngValue And &HBFFFFFFF
                        End If

                    Next lngLoop
                End If
    End Select

    w32Shift = lngValue   ' return new value

End Function

' ***************************************************************************
' Routine:       w32Rotate
'
' Description:   Rotate (sometimes called a circular shift) a Long Integer
'                to the left or right a specified number of bits.  Bits
'                "falling off" the edge wrap around to the opposite end.
'                Some common languages like C/C++ or Java have an operator
'                for this job:  "<<<" and ">>>".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
' Parameters:    lngValue - Number value being manipulated
'                lngBitShift - Number of bits to be manipulated.
'                    If bit value is positive then rotate left
'                    else rotate right.
'
' Returns:       Numeric value after bit manipulation
'
'                   Number                Binary
' Original:        123456789    00000111010110111100110100010101
'   Left 5:       -344350048    11101011011110011010001010100000
'  Right 5:      -1472536984    10101000001110101101111001101000
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified variable names and documented
' ***************************************************************************
Public Function w32Rotate(ByVal lngValue As Long, _
                          ByVal lngBitShift As Long) As Long

    Dim lngLoop    As Long
    Dim lngSignBit As Long

    ' Test amount of bit movement
    Select Case lngBitShift

           Case 0                        ' Nothing to do, return original value
           Case Is > 31:  lngValue = 0   ' return zero too many positive bit positions
           Case Is < -31: lngValue = 0   ' return zero too many negative bit positions

           ' Positive bit shift means rotate left
           Case Is > 0
                If lngValue <> 0 Then
                    For lngLoop = 1 To lngBitShift

                        ' remember the two most significant bits
                        lngSignBit = lngValue And &HC0000000

                        ' clear the bit and shift left by one position
                        lngValue = (lngValue And &H3FFFFFFF) * 2

                        ' if number was negative, then add 1
                        ' if bit 30 was set, then set the sign bit
                        lngValue = lngValue Or _
                                   ((lngSignBit < 0) And &H1&) Or _
                                   (CBool(lngSignBit And &H40000000) And &H80000000)

                    Next lngLoop
                End If

           ' Negative bit shift means rotate right
           Case Is < 0
                If lngValue <> 0 Then
                    For lngLoop = lngBitShift To -1

                        ' remember the sign bit and bit 0
                        lngSignBit = lngValue And &H80000001

                        ' clear the bit and shift right by one position
                        lngValue = (lngValue And &H7FFFFFFF) \ 2

                        ' if number was negative, then reinsert the bit
                        ' if bit 0 was set, then set the sign bit
                        lngValue = lngValue Or _
                                   ((lngSignBit < 0) And &H40000000) Or _
                                   (CBool(lngSignBit And 1) And &H80000000)

                    Next lngLoop
                End If
    End Select

    w32Rotate = lngValue

End Function

' ***************************************************************************
' Routine:       CubeRoot
'
' Description:   Compute cube root of (vntInput), by iteration, to up to 29
'                digits.  The iteration continues until the limit of
'                precision is reached.
'
' Parameters:    vntInput - Value to be evaluated
'
' Returns:       Result of computation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2002  Jay Tanner
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function CubeRoot(ByVal vntInput As Variant) As Variant

    Dim vntTemp    As Variant    ' Argument - May be positive or negative value
    Dim vntWork1   As Variant    ' Any general approximation to cube root
    Dim vntWork2   As Variant    ' Next successive approximation to cube root
    Dim intCounter As Integer    ' Cycle loop control counter
    Dim strSign    As String     ' strSign of argument - Attached to result

    Const ROUTINE_NAME As String = "CubeRoot"

    On Error GoTo CubeRoot_Error

    vntWork1 = CDec(0)   ' Initialize (vntWork1) as decimal data type
    vntWork2 = CDec(0)   ' Initialize (vntWork2) as decimal data type
    vntTemp = vntInput   ' Check for invalid numeric argument

    If Not IsNumeric(vntTemp) Then
        GoTo CubeRoot_Error
    End If

    vntTemp = CDec(vntTemp)   ' Convert argument into decimal data type
    strSign = vbNullString    ' Consider the sign of the argument

    If vntTemp < 0 Then
        vntTemp = -vntTemp
        strSign = "-"
    End If

    If vntTemp = 0 Then
        vntTemp = 1
    Else
        vntWork1 = vntTemp ^ (1 / 3)  ' Use VB cube root as 1st approximation
        intCounter = 50               ' Set limit of iterations to 50 max

        ' A loop to grind out the cube root using a series of
        ' successive approximations, starting with (vntWork1).
        Do
            ' Compute next approx (vntWork2) from (vntWork1)
            vntWork2 = ((2 * vntWork1) + vntTemp / (vntWork1 * vntWork1)) / 3

            ' Check if finished
            If (vntWork2 = vntWork1) Or intCounter <= 0 Then
                Exit Do   ' exit Do..Loop
            End If

            vntWork1 = vntWork2           ' Update approx to current value
            intCounter = intCounter - 1   ' Update limit counter
        Loop
    End If

CubeRoot_CleanUp:
    CubeRoot = TrimStr(strSign & vntWork2)

    vntWork1 = Empty  ' Always empty variants when not needed
    vntWork2 = Empty
    vntTemp = Empty

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CubeRoot_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    GoTo CubeRoot_CleanUp

End Function

' ***************************************************************************
' Routine:       SquareRoot
'
' Description:   Compute square root of (vntInput), by iteration, to up to 29
'                digits.  The iteration continues until the limit of
'                precision is reached.
'
'                Generally, a square root refers to only positive arguments,
'                but this function will accept negative arguments and produce
'                an (imaginary) square root by returning a value with " i"
'                attached to the end.
'
' Parameters:    vntInput - Value to be evaluated
'
' Returns:       Result of computation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2002  Jay Tanner
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function SquareRoot(ByVal vntInput As Variant) As Variant

    Dim vntTemp       As Variant   ' Argument - May be positive or negative value
    Dim vntWork1      As Variant   ' Any general approximation to cube root
    Dim vntWork2      As Variant   ' Next successive approximation to cube root
    Dim intCounter    As Integer   ' Cycle loop control counter
    Dim strSqrMinus1  As String    ' Represents the square root of minus 1

    Const ROUTINE_NAME As String = "SquareRoot"

    On Error GoTo SquareRoot_Error

    strSqrMinus1 = vbNullString
    vntWork1 = CDec(0)           ' Initialize (vntWork1) as decimal data type
    vntWork2 = CDec(0)           ' Initialize (vntWork2) as decimal data type
    vntTemp = vntInput           ' Check for invalid numeric argument

    If Not IsNumeric(vntTemp) Then
        GoTo SquareRoot_Error
    End If

    vntTemp = CDec(vntTemp)    ' Convert argument into decimal data type

    ' Account for a negative argument
    If vntTemp < 0 Then
        vntTemp = -vntTemp
        strSqrMinus1 = " i"
    End If

    ' Check for zero argument
    If vntTemp = 0 Then
        SquareRoot = 0
        Exit Function
    End If

    vntWork1 = Sqr(vntTemp)   ' Use VB square root as 1st approximation
    intCounter = 50           ' Set limit of iterations to 50 max

    ' A loop to grind out the square root using a series of
    ' successive approximations, starting with (vntWork1).
    Do
        ' Compute next approx (vntWork2) from (vntWork1)
        vntWork2 = (vntWork1 + vntTemp / vntWork1) / 2

        ' Check if finished
        If (vntWork2 = vntWork1) Or intCounter <= 0 Then
            Exit Do   ' exit Do..Loop
        End If

        vntWork1 = vntWork2           ' Update approx to current value
        intCounter = intCounter - 1   ' Update limit counter
    Loop

SquareRoot_CleanUp:
    SquareRoot = TrimStr(vntWork2 & strSqrMinus1)

    vntWork1 = Empty  ' Always empty variants when not needed
    vntWork2 = Empty
    vntTemp = Empty

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

SquareRoot_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    GoTo SquareRoot_CleanUp

End Function

' ***************************************************************************
' Routine:       ExamineBit
'
' Description:   The ExamineBit function will return True or False depending
'                on the value of the nth bit (lngBitPosition ) of a long
'                integer (lngValue). The sign bit is not used because if no
'                other bits have been set to "1" then the sign bit would be
'                ignored and remain "0".
'
' Parameters:    lngValue - Data to be evaluated
'                lngBitPosition - Bit position number
'
' Returns:       True or False depending on the value of the nth bit of a
'                long integer.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Rodney Willis
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6222&lngWId=1
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.  Thanks to Alfred Hellmüller for
'              suggesting that I add this routine.
' ***************************************************************************
Public Function ExamineBit(ByVal lngValue As Long, _
                           ByVal lngBitPosition As Long) As Boolean

    Const ROUTINE_NAME As String = "ExamineBit"

    ExamineBit = False  ' preset flag to FALSE

    ' Evaluate the requested bit position
    Select Case lngBitPosition

           Case Is < 0   ' Invalid bit position entered (negative value)
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Is > 30  ' Bit position exceeds string length
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Else
                ExamineBit = CBool(lngValue And (2 ^ lngBitPosition))
    End Select

End Function

' ***************************************************************************
' Routine:       ClearBit
'
' Description:   The ClearBit Sub will change the state of the nth bit
'                (lngBitPosition) of a long integer (lngValue).  The
'                sign bit is not used because if no other bits have
'                been set to "1" then the sign bit would be ignored and
'                remain "0"
'
' Parameters:    lngValue - byte to be manipulated
'                lngBitPosition - Bit position number
'
' Returns:       A new value (lngValue)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Rodney Willis
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6222&lngWId=1
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.  Thanks to Alfred Hellmüller for
'              suggesting that I add this routine.
' 19-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug in my logic.
' ***************************************************************************
Public Sub ClearBit(ByRef lngValue As Long, _
                    ByVal lngBitPosition As Long)

    Const ROUTINE_NAME As String = "ClearBit"

    ' Evaluate the requested bit position
    Select Case lngBitPosition

           Case Is < 0   ' Invalid bit position entered (negative value)
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Is > 30  ' Bit position exceeds string length
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Else
                If ExamineBit(lngValue, lngBitPosition) Then
                    lngValue = lngValue - (2 ^ lngBitPosition)
                End If
    End Select

End Sub

' ***************************************************************************
' Routine:       SetBit
'
' Description:   The SetBit Sub will set the nth bit (lngBitPosition) of a
'                long integer (lngValue).  The sign bit is not used because
'                if no other bits have been set to "1" then the sign bit
'                would be ignored and remain "0".
'
' Parameters:    lngValue - byte to be manipulated
'                lngBitPosition - Bit position number
'
' Returns:       A new value (lngValue)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Rodney Willis
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6222&lngWId=1
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.  Thanks to Alfred Hellmüller for
'              suggesting that I add this routine.
' ***************************************************************************
Public Sub SetBit(ByRef lngValue As Long, _
                  ByVal lngBitPosition As Long)

    Const ROUTINE_NAME As String = "SetBit"

    ' Evaluate the requested bit position
    Select Case lngBitPosition

           Case Is < 0   ' Invalid bit position entered (negative value)
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Is > 30  ' Bit position exceeds string length
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Else
                If ExamineBit(lngValue, lngBitPosition) Then
                    Exit Sub
                Else
                    lngValue = lngValue + (2 ^ lngBitPosition)
                End If
    End Select

End Sub

' ***************************************************************************
' Routine:       ToggleBit
'
' Description:   The ToggleBit Sub will change the state of the nth bit
'                (lngBitPosition) of a long integer (lngValue).  The sign
'                bit is not used because if no other bits have been set to
'                "1" then the sign bit would be ignored and remain "0".
'
' Parameters:    lngValue - byte to be manipulated
'                lngBitPosition - Bit position number
'
' Returns:       A new value (lngValue)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 23-Feb-2000  Rodney Willis
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6222&lngWId=1
' 25-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.  Thanks to Alfred Hellmüller for
'              suggesting that I add this routine.
' ***************************************************************************
Public Sub ToggleBit(ByRef lngValue As Long, _
                     ByVal lngBitPosition As Long)

    Const ROUTINE_NAME As String = "ToggleBit"

    ' Evaluate the requested bit position
    Select Case lngBitPosition

           Case Is < 0   ' Invalid bit position entered (negative value)
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Is > 30  ' Bit position exceeds string length
                lngValue = 0
                InfoMsg "Valid bit positions are 0 to 30" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

           Case Else
                ' See if the bit position is a "1"
                If ExamineBit(lngValue, lngBitPosition) Then
                    lngValue = lngValue - (2 ^ lngBitPosition)   ' set to "0"
                Else
                    lngValue = lngValue + (2 ^ lngBitPosition)   ' set to "1"
                End If
    End Select

End Sub

' ***************************************************************************
' Routine:       GetPrimeNumbers
'
' Description:   This is where an array is filled with LONG INTEGERS only.
'
'                Prime Number - A prime number is a positive integer that
'                has exactly two positive integer factors, 1 and itself.
'                Two is the only even prime number, since any bigger even
'                number is divided by two.  For example, if we list the
'                factors of 28, we have 1, 2, 4, 7, 14, and 28.  That's
'                six factors.  If we list the factors of 29, we only have
'                1 and 29.  Prime numbers have exactly two factors.  So we
'                say that 29 is a prime number and 28 is a composite number
'                and not a prime.
'
'                Eratosthenes worked on prime numbers.  He is remembered for
'                his prime number sieve, the "Sieve of Eratosthenes" which,
'                in modified form, is still an important tool in number
'                theory research.  He was also one of the first ancient Greek
'                scientists to calculate the circumference of the earth, and
'                with an astonishing accuracy.  He lived 276 BC to 194 BC.
'
' Reference:     Ask Dr. Math
'                http://mathforum.org/dr.math/faq/faq.prime.num.html
'
'                The Prime Pages
'                http://primes.utm.edu/
'
'                Primality Testing Is Easy
'                http://mathworld.wolfram.com/news/2002-08-07/primetest/
'
'                10,000 Smallest primes by Peter Alfeld
'                Department of Mathematics, University of Utah
'                http://www.math.utah.edu/~pa/math/p10000.html
'
' Parameters:    lngQtyNeeded - [Optional] - Number of prime numbers to return
'                                            Default = 1
'
' Returns:       An array of prime numbers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic bug in determining starting value
' ***************************************************************************
Public Function GetPrimeNumbers(ByVal lngQtyNeeded As Long, _
                       Optional ByVal lngStartValue As Long = 2, _
                       Optional ByVal lngMaxValue As Long = MAX_LONG) As Long()

    ' Called by HexRepresentation()

    Dim lngNumber    As Long
    Dim lngPrimeCnt  As Long
    Dim alngPrimes() As Long

    Const ROUTINE_NAME As String = "GetPrimeNumbers"

    On Error GoTo GetPrimeNumbers_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GetPrimeNumbers_CleanUp
    End If

    Erase alngPrimes()  ' Always start with empty arrays
    lngPrimeCnt = 0     ' Init array index

    '-------------------------------------------------
    ' Evaluate input parameters
    If lngQtyNeeded < 1 Then
        InfoMsg "Return quantity must be a positive number greater than zero." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    If (lngStartValue < 2) Or (lngMaxValue < 2) Then
        InfoMsg "Starting and maximum values must be a positive number greater than one." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    If lngStartValue > lngMaxValue Then
        InfoMsg "Starting value must be less than or equal maximum value." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If
    '-------------------------------------------------

    ' Size return array to number of
    ' primes needed with a buffer.
    ReDim alngPrimes(lngQtyNeeded + 2)

    ' Determine starting value
    If lngStartValue = 2 Then
        alngPrimes(0) = 2  ' Insert first two primes into array
        alngPrimes(1) = 3
        lngPrimeCnt = 2    ' Number of primes
        lngStartValue = 5  ' Next odd value

    ElseIf lngStartValue = 3 Then
        lngStartValue = 5  ' Next odd value

    Else
        ' Start value must be odd
        If lngStartValue Mod 2 = 0 Then
            lngStartValue = lngStartValue + 1
        End If
    End If

    If lngPrimeCnt >= lngQtyNeeded Then
        ReDim Preserve alngPrimes(lngQtyNeeded)  ' Resize prime number array
        GetPrimeNumbers = alngPrimes()           ' Return prime number array
        GoTo GetPrimeNumbers_CleanUp             ' Time to leave
    End If

    ' Second test to verify values
    ' are still within range
    If lngStartValue > lngMaxValue Then
        InfoMsg "Starting value must be less than or equal maximum value." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo GetPrimeNumbers_CleanUp
    End If

    ' Load rest of requested prime numbers
    ' into an array.  Use step 2 because
    ' only odd value numbers are processed.
    For lngNumber = lngStartValue To lngMaxValue Step 2

        ' Test to see if this is a prime number
        If IsPrime(lngNumber) Then
            alngPrimes(lngPrimeCnt) = lngNumber   ' Add number to array
            lngPrimeCnt = lngPrimeCnt + 1         ' Increment array index
        End If

        ' If enough data has been collected
        ' then exit this loop
        DoEvents
        If lngPrimeCnt >= lngQtyNeeded Then
            Exit For   ' exit For..Next loop
        End If

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit For   ' exit For..Next loop
        End If

    Next lngNumber                                        '

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo GetPrimeNumbers_CleanUp
    End If

    ReDim Preserve alngPrimes(lngQtyNeeded)  ' Resize prime number array
    GetPrimeNumbers = alngPrimes()           ' Return prime number array

GetPrimeNumbers_CleanUp:
    Erase alngPrimes()  ' Always empty arrays when not needed
    On Error GoTo 0     ' Nullify this error trap
    Exit Function

GetPrimeNumbers_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume GetPrimeNumbers_CleanUp

End Function

' ***************************************************************************
' Routine:       IsPrime
'
' Description:   Determines whether a positive whole number is a prime.
'
'                Prime Number - A prime number is a positive integer that
'                has exactly two positive integer factors, 1 and itself.
'                Two is the only even prime number, since any bigger even
'                number is divided by two.  For example, if we list the
'                factors of 28, we have 1, 2, 4, 7, 14, and 28.  That's
'                six factors.  If we list the factors of 29, we only have
'                1 and 29.  Prime numbers have exactly two factors.  So we
'                say that 29 is a prime number and 28 is a composite number.
'
'                Eratosthenes worked on prime numbers.  He is remembered for
'                his prime number sieve, the "Sieve of Eratosthenes" which,
'                in modified form, is still an important tool in number
'                theory research.  He was also one of the first ancient Greek
'                scientists to calculate the circumference of the earth, and
'                with an astonishing accuracy.  He lived 276 BC to 194 BC.
'
' Reference:     Ask Dr. Math
'                http://mathforum.org/dr.math/faq/faq.prime.num.html
'
'                The Prime Pages
'                http://primes.utm.edu/
'
'                Primality Testing Is Easy
'                http://mathworld.wolfram.com/news/2002-08-07/primetest/
'
'                10,000 Smallest primes by Peter Alfeld
'                Department of Mathematics, University of Utah
'                http://www.math.utah.edu/~pa/math/p10000.html
'
' Parameters:    dblNumber - Number to be evaluated
'
' Returns:       TRUE - Number is a prime value
'                FALSE - Number is a composite (not a prime)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-May-2000  Francesco Balena
'              http://www.devx.com/vb2themax/Tip/19051
' 12-Jul-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function IsPrime(ByVal dblNumber As Double) As Boolean

    ' Called by GetPrimeNumbers()

    Dim lngStep    As Long
    Dim dblDivisor As Double
    Dim dblSqrRoot As Double

    Const ROUTINE_NAME As String = "IsPrime"

    IsPrime = False             ' Preset to FALSE for input testing
    dblNumber = Fix(dblNumber)  ' Remove any decimal positions

    If dblNumber < 1# Then
        InfoMsg "Number to be evaluated must be greater than zero." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    ' Pretest with two smallest primes
    If dblNumber Mod 2# = 0 Or _
       dblNumber Mod 3# = 0 Then

        Exit Function
    End If

    IsPrime = True    ' Assume input value is a prime
    dblDivisor = 5#   ' Starting divisor
    lngStep = 2       ' Incremental steps

    ' Calc square root of incoming number
    dblSqrRoot = Int(Sqr(dblNumber)) + 1

    ' No need to go higher than the
    ' square root of the number
    Do While dblSqrRoot > dblDivisor

        If (dblNumber Mod dblDivisor) = 0 Then
            IsPrime = False   ' Not a prime number
            Exit Do           ' exit Do..Loop
        End If

        dblDivisor = dblDivisor + lngStep  ' Update divisor
        lngStep = IIf(lngStep = 2, 4, 2)   ' Toggle between 2 and 4

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            IsPrime = False   ' Set flag to FALSE
            Exit Do           ' exit Do..Loop
        End If

    Loop

End Function


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

Private Function EvaluateNumber(ByVal vntData As Variant, _
                                ByRef lngRetDataType As enumDataType) As Boolean

    ' Called by NumberToBinary()
    '           NumberToHex()
    '           HexToNumber()

    On Error Resume Next

    lngRetDataType = eUnknown   ' Preset values
    EvaluateNumber = False

    ' Verify this is numeric data
    If IsNumericData(CStr(vntData)) Then

        If InStr(1, CStr(vntData), "E", vbTextCompare) > 0 Then
            InfoMsg "No exponential values allowed.", , , 3
            Exit Function
        End If

        If InStr(1, CStr(vntData), ".", vbBinaryCompare) > 0 Then
            InfoMsg "No decimals or non-numeric characters allowed.", , , 3
            Exit Function
        End If

        If (CDec(vntData) < MIN_LONG) Or (CDec(vntData) > MAX_LONG) Then
            InfoMsg "Input value exceeds range of a long integer.", , , 3
            Exit Function
        End If

        ' If value is greater than or equal zero
        ' and less than or equal 255
        If (CLng(vntData) >= 0) And (CLng(vntData) <= 255) Then
            lngRetDataType = eByte
            EvaluateNumber = True

            ' If value is greater than or equal to -32768
            ' and less than or equal to 32767
        ElseIf (CLng(vntData) >= MIN_INT) And (CLng(vntData) <= MAX_INT) Then
            lngRetDataType = eShort
            EvaluateNumber = True

            ' If value is greater than or equal to -2147483648
            ' and less than or equal to 2147483647
        ElseIf (CLng(vntData) >= MIN_LONG) And (CLng(vntData) <= MAX_LONG) Then
            lngRetDataType = eLong
            EvaluateNumber = True

        Else
            ' Unknown data (Should never get here)
        End If

    Else
        InfoMsg "Non-numeric data is not allowed.", , , 3
    End If

    On Error GoTo 0

End Function

Private Function LoadBinaryArray(ByVal lngBits As enumBITS) As String()

    ' Called by NumberToBinary()
    '           StringToBinary()
    '           HexToBinary()

    Dim lngIndex   As Long
    Dim astrBin4() As String
    Dim astrBin8() As String

    Erase astrBin4()    ' Always start with empty arrays
    Erase astrBin8()

    ReDim astrBin4(16)  ' Size base array

                           ' Hex  Decimal
    astrBin4(0) = "0000"   '  0      0
    astrBin4(1) = "0001"   '  1      1
    astrBin4(2) = "0010"   '  2      2
    astrBin4(3) = "0011"   '  3      3
    astrBin4(4) = "0100"   '  4      4
    astrBin4(5) = "0101"   '  5      5
    astrBin4(6) = "0110"   '  6      6
    astrBin4(7) = "0111"   '  7      7
    astrBin4(8) = "1000"   '  8      8
    astrBin4(9) = "1001"   '  9      9
    astrBin4(10) = "1010"  '  A     10
    astrBin4(11) = "1011"  '  B     11
    astrBin4(12) = "1100"  '  C     12
    astrBin4(13) = "1101"  '  D     13
    astrBin4(14) = "1110"  '  E     14
    astrBin4(15) = "1111"  '  F     15

    ' Create an array of binary data.
    Select Case lngBits

           Case eBits4   ' astrBin4(0) = "0000" ... astrBin4(15) = "1111"
                LoadBinaryArray = astrBin4()

           Case eBits8   ' astrBin8(0) = "00000000" ... astrBin8(255) = "11111111"
                ReDim astrBin8(MAX_BYTE)

                For lngIndex = 0 To (MAX_BYTE - 1)
                    astrBin8(lngIndex) = astrBin4(lngIndex \ &H10&) & _
                                         astrBin4(lngIndex And &HF&)
                Next lngIndex

                LoadBinaryArray = astrBin8()
    End Select

    Erase astrBin4()  ' Always empty arrays when not needed
    Erase astrBin8()

End Function

Private Sub LoadShiftArrays()

    ' Called by w8Shift()

    On Error Resume Next

    If IsArrayInitialized(mabytPower2()) Then

        ' See if arrays are already loaded
        If mabytPower2(0) = 1 And mabytValue(0) = 1 Then
            mblnLoaded = True              ' Set flag to TRUE
            GoTo LoadShiftArrays_CleanUp   ' Leave this routine
        End If
    End If

    ' Array has not been loaded
    mblnLoaded = False     ' Verify flag is FALSE
    Erase mabytPower2()    ' Always start with empty arrays
    Erase mabytValue()

    ' Load power 2 array      Binary
    mabytPower2(0) = 1     ' 00000001
    mabytPower2(1) = 2     ' 00000010
    mabytPower2(2) = 4     ' 00000100
    mabytPower2(3) = 8     ' 00001000
    mabytPower2(4) = 16    ' 00010000
    mabytPower2(5) = 32    ' 00100000
    mabytPower2(6) = 64    ' 01000000
    mabytPower2(7) = 128   ' 10000000

    ' Load byte value array
    mabytValue(0) = 1      ' 00000001
    mabytValue(1) = 3      ' 00000011
    mabytValue(2) = 7      ' 00000111
    mabytValue(3) = 15     ' 00001111
    mabytValue(4) = 31     ' 00011111
    mabytValue(5) = 63     ' 00111111
    mabytValue(6) = 127    ' 01111111
    mabytValue(7) = 255    ' 11111111

    mblnLoaded = True      ' Set flag to TRUE

LoadShiftArrays_CleanUp:
    On Error GoTo 0

End Sub

Private Sub ResetVariables()

    Erase mabytPower2()  ' Empty 8-bit shift arrays
    Erase mabytValue()
    mblnLoaded = False   ' Set flag to FALSE

End Sub

' ***************************************************************************
' Routine:       ForceEnumCase
'
' Description:   This routine exists only to ensure the CASE of these
'                constants are not altered while editing code, as can
'                happen with Enums.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Apr-2010  Karl E. Peterson
'              Customizing the Ride Part 2
'              https://visualstudiomagazine.com/articles/2010/04/12/customizing-the-ride-part-2.aspx
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Modified to support this module
' ***************************************************************************
#If False Then
Private Sub ForceEnumCase()
    ' Enum enumBITS
    Const eBits4 As Long = 0&
    Const eBits8 As Long = 1&

    ' Enum enumDataType
    Const eLong    As Long = 0&
    Const eShort   As Long = 1&
    Const eByte    As Long = 2&
    Const eUnknown As Long = 3&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables
    StopProcessing = False

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables

End Sub
