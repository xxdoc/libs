VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMath64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cMath64  (clsMath64.cls)
'
' Description:   Multiple 64-bit manipulating routines.  Some are not used.
'                I left the unused ones so I would not have to search for
'                them if I use this module in another application.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-JAN-2000  Kenneth Ives  kenaso@tx.rr.com
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated IsHexData() and IsBinaryData() routines.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
'              Added parameter testing in HexRepresentation() routine.
' 10-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Updated GetPrimeNumbers(), IsHexData() routines.
' 11-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated ConvertBase() and its supporting routines
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed LoadCubeRoots() and LoadSquareRoots() routines to
'              display output details as shown in their flowerboxes if user
'              hard codes the exact prime number indicated in the routine.
' 02-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed ConvertBase() routine.
' 10-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated logic testing in IsHexData() and IsBinaryData()
'              routines.
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote ConvertBase() routine.
'              - Added Base2_to_Base16(), Base2_to_Base10(), Base16_to_Base2(),
'                Base16_to_Base10(), Base10_to_Base2(), Base10_to_Base16() routines.
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Added routines Add_64(), Subtract_64(), Shift_64(),
'              Rotate_64(), Xor_64() to process two long integers as
'              a single 64-bit hex string.
' 29-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Updated BigAdd(), BigSubtract(), BigMultiply(), BigDivide()
'              routines.
' 15-Jan-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Established minimum and maximum numeric ranges
'              - Updated several routines
' 06-Dec-2016  Kenneth Ives  kenaso@tx.rr.com
'              Updated several routines
' 17-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              - Updated Base2_to_Base16(), Base16_to_Base2() routines
'              - Updated HexRepresentation(), LoadSquareRoots(), and
'                LoadCubeRoots() routines to allow print detailed
'                structure of a specific Prime Number to the VB6 IDE
'                Immediate (Debug) Window.
' 24-Feb-2017  Kenneth Ives  kenaso@tx.rr.com
'              Minor tweaking to bit comparison routines
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
' In 1998 the IEC changed it's measurements so that what you consider
' a gigabyte (1024 mb or 2^30) was renamed a gibibyte.  A gibibyte is
' now formally recognised as 1000 mb, even though no operating system
' (like windows) will use this definition.
'
' Hard drives use the term "gigabytes" which would be what the 300 GB
' stands for.  But, Microsoft uses a different way to measure gigabytes
' (they're actually gibibytes) so that's why you "lost" some 21 GB from
' your hard drive. It happens to everyone. It is just the conflict of
' two different numbering systems.
'
' The difference between those two numbering systems is seven percent.
' So, when you buy a 300 GB drive, Windows sees it as 279 GB because:
'
'    300 GB - 7% = 279 GB
'
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME  As String = "clsMath64"
  Private Const ZEROES_8     As String = "00000000"
  Private Const ZEROES_16    As String = ZEROES_8 & ZEROES_8
  Private Const MIN_RANGE    As String = "-18446744073709551615"
  Private Const MAX_RANGE    As String = "9223372036854775807"
  Private Const HEX_BASE     As Double = 16#   ' Used for computing cube & square roots
  Private Const HEX_LENGTH   As Long = 16
  Private Const BIN_LENGTH   As Long = 64
  Private Const MAX_LONG     As Long = &H7FFFFFFF   ' 2147483647

' ***************************************************************************
' Public Enumerations
' ***************************************************************************
  Public Enum enumHEX_REPRESENTATION
      eCubeRoots    ' 0
      eSquareRoots  ' 1
  End Enum

  Public Enum enumHEX_RETURN_FORMAT
      e16Chars    ' 0
      eLeft8      ' 1
      eRight8     ' 2
  End Enum

' ***************************************************************************
' Private Enumerations
' ***************************************************************************
  Private Enum enum64_BaseValues
      e64_Binary = 2
      e64_Numeric = 10
      e64_Hex = 16
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)



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
' Routine:       w64HexToNumber
'
' Description:   Convert 16 character hex string (64-bit) to a big number in
'                string format.
'
' Parameters:    strHex - hex string
'
' Returns:       Big number in string format
'
' Example:       Input:   428a2f98d728ae22
'                Output:  4794697086780616226
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexToNumber(ByVal strHex As String) As String

    ' Called by w64HexMultiply()
    '           w64HexDivide()

    If gblnStopProcessing Then
        Exit Function
    End If

    w64HexToNumber = ConvertBase(strHex, e64_Hex, e64_Numeric)

End Function

' ***************************************************************************
' Routine:       w64HexToBinary
'
' Description:   Convert 16 character hex string (64-bit) to binary string.
'
' Parameters:    strHex - hex string
'
' Returns:       Binary string
'
' Example:       Input:   428a2f98d728ae22
'                Output:  0100001010001010001011111001100011010111001010001010111000100010
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexToBinary(ByVal strHex As String) As String

    ' Called by w64Shift()
    '           w64Rotate()

    If gblnStopProcessing Then
        Exit Function
    End If

    w64HexToBinary = ConvertBase(strHex, e64_Hex, e64_Binary)

End Function

' ***************************************************************************
' Routine:       w64BinaryToNumber
'
' Description:   Convert a binary string to a big number in string format.
'
' Parameters:    strBinary - Binary string
'
' Returns:       Big number in string format
'
' Example:       Input:   0100001010001010001011111001100011010111001010001010111000100010
'                Output:  4794697086780616226
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64BinaryToNumber(ByVal strBinary As String) As String

    If gblnStopProcessing Then
        Exit Function
    End If

    w64BinaryToNumber = ConvertBase(strBinary, e64_Binary, e64_Numeric)

End Function

' ***************************************************************************
' Routine:       w64BinaryToHex
'
' Description:   Convert a binary string to a hex string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       16 character hex string (64-bit)
'
' Example:       Input:   0100001010001010001011111001100011010111001010001010111000100010
'                Output:  428a2f98d728ae22
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64BinaryToHex(ByVal strBinary As String) As String

    If gblnStopProcessing Then
        Exit Function
    End If

    w64BinaryToHex = ConvertBase(strBinary, e64_Binary, e64_Hex)

End Function

' ***************************************************************************
' Routine:       w64NumberToHex
'
' Description:   Convert a big (whole) number in string format to a sixteen
'                character hex string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       16 character hex string (64-bit)
'
' Example:       Input:   4794697086780616226
'                Output:  428a2f98d728ae22
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64NumberToHex(ByVal strNumber As String) As String

    ' Called by w64HexMultiply()
    '           w64HexDivide()

    If gblnStopProcessing Then
        Exit Function
    End If

    w64NumberToHex = ConvertBase(strNumber, e64_Numeric, e64_Hex)

End Function

' ***************************************************************************
' Routine:       w64NumberToBinary
'
' Description:   Convert a big (whole) number in string format to a binary
'                string.
'
' Parameters:    strBinary - Binary string
'
' Returns:       Binary string
'
' Example:       Input:   4794697086780616226
'                Output:  0100001010001010001011111001100011010111001010001010111000100010
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64NumberToBinary(ByVal strNumber As String) As String

    If gblnStopProcessing Then
        Exit Function
    End If

    w64NumberToBinary = ConvertBase(strNumber, e64_Numeric, e64_Binary)

End Function

' ***************************************************************************
' Routine:       w64HexAdd
'
' Description:   Perform Hex addition on two 16 char hex values (64-bit).
'
'                Input:     b5c0fbcfec4d3b2f
'                           428a2f98d728ae22
'                           ----------------
'                Output:    f84b2b68c375e951
'
' Reference:     Ask Dr. Math - Adding Hexadecimals
'                http://mathforum.org/library/drmath/view/55725.html
'
' Parameters:    strHex1 - First hex string (64-bit)
'                strHex2 - Second hex string (64-bit)
'
' Returns:       16 char hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine.  I kept this original routine because I needed
'              the speed.
' ***************************************************************************
Public Function w64HexAdd(ByVal strHex1 As String, _
                          ByVal strHex2 As String) As String

    Dim lngIndex  As Long
    Dim strOutput As String
    Dim vntResult As Variant

    Const ROUTINE_NAME As String = "w64HexAdd"

    On Error GoTo w64HexAdd_Error

    vntResult = Empty
    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo w64HexAdd_CleanUp
    End If

    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Make sure we have some valid data
    If StrComp(ZEROES_16, strHex1, vbBinaryCompare) = 0 Then
        If StrComp(ZEROES_16, strHex2, vbBinaryCompare) = 0 Then
            w64HexAdd = ZEROES_16
        Else
            w64HexAdd = strHex2
        End If

        Exit Function
    Else
        If StrComp(ZEROES_16, strHex2, vbBinaryCompare) = 0 Then
            w64HexAdd = strHex1
            Exit Function
        End If
    End If

    ' Step thru one value at a time and perform the hex addition
    For lngIndex = Len(strHex1) To 1 Step -1

        vntResult = CDec(vntResult + _
                    Val("&H" & Mid$(strHex1, lngIndex, 1)) + _
                    Val("&H" & Mid$(strHex2, lngIndex, 1)))

        strOutput = Hex$(vntResult Mod HEX_LENGTH) & strOutput
        vntResult = CDec(Int(vntResult / HEX_LENGTH))

    Next lngIndex

w64HexAdd_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        w64HexAdd = vbNullString
    Else
        ' If the final calculated value is greater than
        ' zero then prefix it to the output string
        If Fix(vntResult) > 0 Then
            strOutput = Hex$(vntResult) & strOutput
        End If

        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)   ' Verify 16 character length
        w64HexAdd = UCase$(strOutput)                           ' Convert hex string to uppercase
    End If

    vntResult = Empty  ' Always empty variants when not needed
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

w64HexAdd_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexAdd_CleanUp

End Function

' ***************************************************************************
' Routine:       w64HexSubtract
'
' Description:   Perform Hex subtraction on two 16 char hex values (64-bit).
'
'                Input:     b5c0fbcfec4d3b2f
'                           428a2f98d728ae22
'                           ----------------
'                Output:    7336cc3715248d0d
'
' Reference:     Ask Dr. Math - Adding Hexadecimals
'                http://mathforum.org/library/drmath/view/55725.html
'
' Parameters:    strHex1 - First hex string (64-bit)
'                strHex2 - Second hex string (64-bit)
'
' Returns:       16 char hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine.  I kept this original routine because I needed
'              the speed.
' ***************************************************************************
Public Function w64HexSubtract(ByVal strHex1 As String, _
                               ByVal strHex2 As String) As String

    Dim lngCarry  As Long
    Dim lngIndex  As Long
    Dim lngResult As Long
    Dim lngValue1 As Long
    Dim lngValue2 As Long
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64HexSubtract"

    On Error GoTo w64HexSubtract_Error

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo w64HexSubtract_CleanUp
    End If

    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Make sure we have some valid data
    If StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Then
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexSubtract = ZEROES_16
        Else
            w64HexSubtract = strHex2
        End If

        Exit Function
    Else
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexSubtract = strHex1
            Exit Function
        End If
    End If

    ' Step thru one value at a time and perform the hex addition
    For lngIndex = Len(strHex1) To 1 Step -1

        lngValue1 = CLng("&H" & Mid$(strHex1, lngIndex, 1))
        lngValue2 = CLng("&H" & Mid$(strHex2, lngIndex, 1))

        If lngCarry > 0 Then
            lngValue1 = lngValue1 - 1
        End If

        If lngValue1 < lngValue2 Then
            lngCarry = HEX_LENGTH
        Else
            lngCarry = 0
        End If

        lngResult = (lngValue1 + lngCarry) - lngValue2
        strOutput = Hex$(lngResult Mod HEX_LENGTH) & strOutput

    Next lngIndex

w64HexSubtract_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        w64HexSubtract = vbNullString
    Else
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)   ' Verify 16 character length
        w64HexSubtract = UCase$(strOutput)                      ' Convert hex string to uppercase
    End If

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexSubtract_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexSubtract_CleanUp

End Function

' ***************************************************************************
' Routine:       w64HexMultiply
'
' Description:   Perform Hex multiplication on two 16 char hex values (64-bit).
'
'                Input:     428a2f98d728ae22
'                           b5c0fbcfec4d3b2f
'                           ----------------
'                Output:    2f3ddcbe4a93effec7c951e798d3ce3e  Complete output
'                Return:    2f3ddcbe4a93effe                  Return first 16 chars only
'
' Parameters:    strHex1 - First hex string (64-bit) (Multiplicand)
'                strHex2 - Second hex string (64-bit)(Multiplier)
'
' Returns:       Left-most 16 chars of the newly calculated hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexMultiply(ByVal strHex1 As String, _
                               ByVal strHex2 As String) As String

    Dim strOutput  As String
    Dim strNumber1 As String
    Dim strNumber2 As String

    Const ROUTINE_NAME As String = "w64HexMultiply"

    On Error GoTo w64HexMultiply_Error

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo w64HexMultiply_CleanUp
    End If

    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Make sure we have some valid data
    If StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Then
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexMultiply = ZEROES_16
        Else
            w64HexMultiply = strHex2
        End If

        Exit Function
    Else
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexMultiply = strHex1
            Exit Function
        End If
    End If

    strNumber1 = w64HexToNumber(strHex1)              ' convert hex values to large number
    strNumber2 = w64HexToNumber(strHex2)
    strOutput = BigMultiply(strNumber1, strNumber2)   ' perform multiplication

w64HexMultiply_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        w64HexMultiply = vbNullString
    Else
        strOutput = w64NumberToHex(strOutput)                   ' convert computation to hex
        strOutput = Left$(strOutput, HEX_LENGTH)                ' Save first 16 characters
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)   ' Verify 16 character length
        w64HexMultiply = UCase$(strOutput)                      ' Convert hex string to uppercase
    End If

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexMultiply_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexMultiply_CleanUp

End Function

' ***************************************************************************
' Routine:       w64HexDivide
'
' Description:   Perform Hex division on two 16 char hex values (64-bit).
'
'                           Quotient           Dividend          Divisor
'                    Ex:  00028AF4A2855646 = 428a2f98d728ae22  /  1a2b
'
' Parameters:    strHex1 - First hex string
'                strHex2 - Second hex string
'
' Returns:       16 char hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function w64HexDivide(ByVal strHex1 As String, _
                             ByVal strHex2 As String) As String

    Dim strOutput   As String
    Dim strDivisor  As String
    Dim strDividend As String

    Const ROUTINE_NAME As String = "w64HexDivide"

    On Error GoTo w64HexDivide_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo w64HexDivide_CleanUp
    End If

    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Make sure we have some valid data
    If StrComp(ZEROES_16, strHex1, vbTextCompare) = 0 Then
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then

            w64HexDivide = ZEROES_16
            Exit Function
        End If
    Else
        If StrComp(ZEROES_16, strHex2, vbTextCompare) = 0 Then
            w64HexDivide = strHex1
            Exit Function
        End If
    End If

    strDividend = w64HexToNumber(strHex1)   ' convert hex values to large number
    strDivisor = w64HexToNumber(strHex2)

    ' Test the numbers
    If Val(strDividend) < Val(strDivisor) Then

        InfoMsg "The Divisor cannot exceed the Dividend." & vbNewLine & vbNewLine & _
                "Dividend:   " & strHex1 & " (" & strDividend & ")" & vbNewLine & _
                "Divisor:    " & strHex2 & " (" & strDivisor & ")" & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        GoTo w64HexDivide_CleanUp
    End If

    strOutput = BigDivide(strDividend, strDivisor)   ' perform the calculation

w64HexDivide_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        w64HexDivide = vbNullString
    Else
        strOutput = w64NumberToHex(strOutput)                   ' convert computation to hex
        strOutput = Left$(strOutput, HEX_LENGTH)                ' Save first 16 characters
        strOutput = Right$(ZEROES_16 & strOutput, HEX_LENGTH)   ' Verify 16 character length
        w64HexDivide = UCase$(strOutput)                        ' Convert hex string to uppercase
    End If

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64HexDivide_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64HexDivide_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Shift
'
'' Description:  Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift do not wrap around.  Fill bits
'                coming in from the opposite side are zeros.  Some common
'                languages like C/C++ or Java have an operator for this
'                job: ">>" and "<<".
'
' Parameters:    strHex - hex value to be manipulated
'                lngBitShift - Number of shift positions
'                              Positive value = left shift
'                              Negative value = right shift
'                blnReturnAsHex - [Optional] how the return data is formatted
'                                TRUE - Data returned in hex format (Default)
'                                FALSE - Data returned in binary format
'
' Returns:       New manipulated hex string
'
'                  Hex                          Binary
' Original:  1234567890abcdef  0001001000110100010101100111100010010000101010111100110111101111
'   Left 5:  468acf121579bde0  0100011010001010110011110001001000010101011110011011110111100000
'  Right 5:  0091a2b3c4855e6f  0000000010010001101000101011001111000100100001010101111001101111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan  dermot@bitwisemag.com
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function w64Shift(ByVal strHex As String, _
                         ByVal lngBitShift As Long, _
                Optional ByVal blnReturnAsHex As Boolean = True) As String

    Dim lngIndex    As Long
    Dim lngTmpLeft  As Long
    Dim lngHexLeft  As Long
    Dim lngHexRight As Long
    Dim lngTmpRight As Long

    Const ROUTINE_NAME As String = "w64Shift"

    On Error GoTo w64Shift_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test the amount of bit shifting
    Select Case lngBitShift

           Case 0:        w64Shift = strHex      ' Return original value if no shift positions
           Case Is < -63: w64Shift = ZEROES_16   ' Return zeroes if too many negative shift positions
           Case Is > 63:  w64Shift = ZEROES_16   ' Return zeroes if too many positive shift positions

           Case Is > 0
                ' A positive bit shift value means shift left
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = ("&H" & Mid$(strHex, 1, 8))
                lngHexRight = ("&H" & Mid$(strHex, 9, 8))

                ' We need to shift left by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = 1 To lngBitShift

                    ' Remember the 2 most significant bits
                    lngTmpLeft = lngHexRight And &HC0000000
                    lngTmpRight = lngHexLeft And &HC0000000

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before multiplying (the code shifts a single
                    ' bit to the left).
                    lngHexRight = lngHexRight And &H3FFFFFFF
                    lngHexRight = lngHexRight * 2

                    lngHexLeft = lngHexLeft And &H3FFFFFFF
                    lngHexLeft = lngHexLeft * 2

                    ' Set the Most Significant Bit (MSB)
                    If lngTmpRight And &H40000000 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    Else
                        lngHexLeft = lngHexLeft And &H7FFFFFFF
                    End If

                    ' If the number was negative, then add 1
                    ' if bit 30 was set, then set the sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexLeft = lngHexLeft Or 1
                    End If

                    ' Reset the Most Significant Bit (MSB)
                    If lngTmpLeft And &H40000000 Then
                        lngHexRight = lngHexRight Or &H80000000
                    Else
                        lngHexRight = lngHexRight And &H7FFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Shift = UCase$(strHex)           ' return hex string in uppercase
                Else
                    w64Shift = w64HexToBinary(strHex)   ' return as binary string
                End If

           Case Is < 0
                ' A negative bit shift value means shift right
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = ("&H" & Mid$(strHex, 1, 8))
                lngHexRight = ("&H" & Mid$(strHex, 9, 8))

                ' We need to shift right by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = lngBitShift To -1

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &H80000001
                    lngTmpRight = lngHexLeft And &H80000001

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before dividing (the code shifts a single
                    ' bit to the right).
                    lngHexRight = lngHexRight And &H7FFFFFFF
                    lngHexRight = lngHexRight \ 2

                    lngHexLeft = lngHexLeft And &H7FFFFFFF
                    lngHexLeft = lngHexLeft \ 2

                    If lngTmpRight And 1 Then
                        lngHexRight = lngHexRight Or &H80000000
                    End If

                    ' Set or clear the old sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexRight = lngHexRight Or &H40000000
                    Else
                        lngHexRight = lngHexRight And &HBFFFFFFF
                    End If

                    ' if the number was negative, then re-insert
                    ' the bit if bit 0 was set, then set the sign bit
                    If lngTmpRight And &H80000000 Then
                        lngHexLeft = lngHexLeft Or &H40000000
                    Else
                        lngHexLeft = lngHexLeft And &HBFFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Shift = UCase$(strHex)           ' return hex string in uppercase
                Else
                    w64Shift = w64HexToBinary(strHex)   ' return as binary string
                End If
    End Select

w64Shift_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Shift_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    w64Shift = vbNullString
    gblnStopProcessing = True
    Resume w64Shift_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Rotate
'
' Description:   Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift wrap around to the opposite end.
'                Some common languages like C/C++ or Java have an operator
'                for this job: ">>>" or "<<<".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
' Parameters:    strHex - hex value to be manipulated
'                lngBitShift - Number of shift positions
'                              Positive value = left rotation
'                              Negative value = right rotation
'                blnReturnAsHex - [Optional] how the return data is formatted
'                                TRUE - Data returned in hex format (Default)
'                                FALSE - Data returned in binary format
'
' Returns:       New manipulated hex string
'
'                  Hex                          Binary
' Original:  1234567890abcdef  0001001000110100010101100111100010010000101010111100110111101111
'   Left 5:  468acf121579bde2  0100011010001010110011110001001000010101011110011011110111100010
'  Right 5:  7891a2b3c4855e6f  0111100010010001101000101011001111000100100001010101111001101111
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-May-2006  Dermot Hogan  dermot@bitwisemag.com
'              http://www.bitwisemag.com/2/Bit-Shifting-in-Visual-Basic-6
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function w64Rotate(ByVal strHex As String, _
                          ByVal lngBitShift As Long, _
                 Optional ByVal blnReturnAsHex As Boolean = True) As String

    Dim lngIndex    As Long
    Dim lngTmpLeft  As Long
    Dim lngHexLeft  As Long
    Dim lngHexRight As Long
    Dim lngTmpRight As Long

    Const ROUTINE_NAME As String = "w64Rotate"

    On Error GoTo w64Rotate_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Test the amount of bit shifting
    Select Case lngBitShift

           Case 0:        w64Rotate = strHex      ' Return original value if no shift positions
           Case Is < -63: w64Rotate = ZEROES_16   ' Return zeroes if too many negative shift positions
           Case Is > 63:  w64Rotate = ZEROES_16   ' Return zeroes if too many positive shift positions

           Case Is > 0
                ' A positive bit shift value means rotate left
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = ("&H" & Mid$(strHex, 1, 8))
                lngHexRight = ("&H" & Mid$(strHex, 9, 8))

                ' We need to shift left by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = 1 To lngBitShift

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &HC0000000
                    lngTmpRight = lngHexLeft And &HC0000000

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before multiplying (the code shifts a single
                    ' bit to the left).
                    lngHexRight = lngHexRight And &H3FFFFFFF
                    lngHexRight = lngHexRight * 2

                    lngHexLeft = lngHexLeft And &H3FFFFFFF
                    lngHexLeft = lngHexLeft * 2

                    ' Set the Most Significant Bit (MSB)
                    If lngTmpRight And &H40000000 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    Else
                        lngHexLeft = lngHexLeft And &H7FFFFFFF
                    End If

                    ' Rotate Most Significant Bit (MSB) into
                    ' Least Significant Bit (LSB)
                    If lngTmpRight And &H80000000 Then
                        lngHexRight = lngHexRight Or 1
                    End If

                    If lngTmpLeft And &H80000000 Then
                        lngHexLeft = lngHexLeft Or 1
                    End If

                    ' Reset the Most Significant Bit (MSB)
                    If lngTmpLeft And &H40000000 Then
                        lngHexRight = lngHexRight Or &H80000000
                    Else
                        lngHexRight = lngHexRight And &H7FFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Rotate = UCase$(strHex)           ' return hex string in uppercase
                Else
                    w64Rotate = w64HexToBinary(strHex)   ' return as binary string
                End If

           Case Is < 0
                ' A negative bit shift value means rotate right
                '
                ' Prepare data prior to bit movement.
                ' Split hex value into two long integers.
                strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)
                lngHexLeft = ("&H" & Mid$(strHex, 1, 8))
                lngHexRight = ("&H" & Mid$(strHex, 9, 8))

                ' We need to shift right by one bit but
                ' take account of overflow error from
                ' VB.  So mask accordingly.
                For lngIndex = lngBitShift To -1

                    ' Save the sign bit
                    lngTmpLeft = lngHexRight And &H80000001
                    lngTmpRight = lngHexLeft And &H80000001

                    ' To prevent an overflow from the transition from
                    ' a positive to a negative quantity when you shift
                    ' to the left.  Mask the top two bits of a 32 bit
                    ' word before dividing (the code shifts a single
                    ' bit to the right).
                    lngHexRight = lngHexRight And &H7FFFFFFF
                    lngHexRight = lngHexRight \ 2

                    lngHexLeft = lngHexLeft And &H7FFFFFFF
                    lngHexLeft = lngHexLeft \ 2

                    ' Rotate Most Significant Bit (MSB) into
                    ' Least Significant Bit (LSB)
                    If lngTmpLeft And 1 Then
                        lngHexLeft = lngHexLeft Or &H80000000
                    End If

                    If lngTmpRight And 1 Then
                        lngHexRight = lngHexRight Or &H80000000
                    End If

                    ' Set or clear the old sign bit
                    If lngTmpLeft And &H80000000 Then
                        lngHexRight = lngHexRight Or &H40000000
                    Else
                        lngHexRight = lngHexRight And &HBFFFFFFF
                    End If

                    ' if the number was negative, then re-insert
                    ' the bit if bit 0 was set, then set the sign bit
                    If lngTmpRight And &H80000000 Then
                        lngHexLeft = lngHexLeft Or &H40000000
                    Else
                        lngHexLeft = lngHexLeft And &HBFFFFFFF
                    End If

                Next lngIndex

                ' rejoin the long integers and convert back to hex
                strHex = Right$(ZEROES_8 & Hex$(lngHexLeft), 8) & _
                         Right$(ZEROES_8 & Hex$(lngHexRight), 8)

                ' Return manipulated value
                If blnReturnAsHex Then
                    w64Rotate = UCase$(strHex)           ' return hex string in uppercase
                Else
                    w64Rotate = w64HexToBinary(strHex)   ' return as binary string
                End If
    End Select

w64Rotate_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Function

w64Rotate_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    w64Rotate = vbNullString
    gblnStopProcessing = True
    Resume w64Rotate_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_NOT
'
' Description:   NOT expression used for bit comparison does not exist
'                in Visual Basic.
'
'                This function looks at the binary representation of
'                the values of the expression and does a bitwise negation
'                operation on it. The result of this operation behaves
'                as shown in this example:
'
'                     X = (NOT 1011)
'
'                         1011
'                        ------
'                     X = 0100    The bits are reversed
'
' Special Note:  This is the same as the Visual Basic bit function NOT except
'                that it works with 64-bit hex strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
' Output:   bd75d06728d751dd  1011110101110101110100000110011100101000110101110101000111011101
'
' Parameters:    strHex - Hex string to be evaluated (input/output)
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_NOT(ByVal strHex As String) As String

    Dim lngPos    As Long
    Dim lngValue  As Long
    Dim lngIndex  As Long
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64Hex_NOT"

    On Error GoTo w64Hex_NOT_Error

    DoEvents
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Initialize variable to proper length
    strHex = Right$(ZEROES_16 & strHex, HEX_LENGTH)

    ' Please note that it is faster to load an output string
    ' with spaces and then insert the data.  Constantly
    ' appending data will cause the system to have to
    ' reallocate space, re-evaluate new string size and
    ' possibly designate a new internal address to hold this
    ' data.  All of which takes extra cycles and time.
    strOutput = Space$(24)   ' Extra spaces
    lngPos = 1               ' Starting position in output string

    For lngIndex = 1 To HEX_LENGTH Step 2

        lngValue = (Not ("&H" & Mid$(strHex, lngIndex, 2)))

        Mid$(strOutput, lngPos, 2) = Right$("0" & Hex$(lngValue), 2)
        lngPos = lngPos + 2   ' Increment position pointer

    Next lngIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = UCase$(TrimStr(strOutput))   ' Return hex string in uppercase
    End If

w64Hex_NOT_CleanUp:
    w64Hex_NOT = strOutput  ' Return new value
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

w64Hex_NOT_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_NOT_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_AND
'
' Description:   The And operator also performs a bitwise comparison of
'                identically positioned bits in two numeric expressions
'                and sets the corresponding bit in result according to
'                the following example:
'
'                    X = 1011 AND 1001
'
'                         1011
'                         1001
'                        ------
'                     X = 1001    Different bits become "0"
'
' Special Note:  This is the same as the Visual Basic bit function AND except
'                that it works with 64-bit hex strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   02820a8046002e20  0000001010000010000010101000000001000110000000000010111000100000
'
' Parameters:    strHex1 - Hex string to be evaluated (input/output)
'                strHex2 - Hex string to be evaluated (input/output)
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_AND(ByVal strHex1 As String, _
                           ByVal strHex2 As String) As String

    Dim lngPos    As Long
    Dim lngValue  As Long
    Dim lngIndex  As Long
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64Hex_AND"

    On Error GoTo w64Hex_AND_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify input data is same length
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Please note that it is faster to load an output string
    ' with spaces and then insert the data.  Constantly
    ' appending data will cause the system to have to
    ' reallocate space, re-evaluate new string size and
    ' possibly designate a new internal address to hold this
    ' data.  All of which takes extra cycles and time.
    strOutput = Space$(24)   ' Extra spaces
    lngPos = 1               ' Starting position in output string

    ' Perform "AND" binary comparisons
    For lngIndex = 1 To HEX_LENGTH Step 2

        lngValue = ("&H" & Mid$(strHex1, lngIndex, 2)) And _
                   ("&H" & Mid$(strHex2, lngIndex, 2))

        Mid$(strOutput, lngPos, 2) = Right$("0" & Hex$(lngValue), 2)
        lngPos = lngPos + 2   ' Increment position pointer

    Next lngIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = UCase$(TrimStr(strOutput))   ' Return hex string in uppercase
    End If

w64Hex_AND_CleanUp:
    w64Hex_AND = strOutput  ' Return new value
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

w64Hex_AND_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_AND_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_OR
'
' Description:   The OR operator also performs a bitwise comparison of
'                identically positioned bits in two numeric expressions
'                and sets the corresponding bit in result according to
'                the following example:
'
'                     X = 1011 OR 1001
'
'                         1011
'                         1001
'                        ------
'                     X = 1011    The higher bit value wins
'
' Special Note:  This is the same as the Visual Basic bit function OR except
'                that it works with 64-bit hex strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   67bf2f9dd7faaffe  0110011110111111001011111001110111010111111110101010111111111110
'
' Parameters:    strHex1 - Hex string to be evaluated (input/output)
'                strHex2 - Hex string to be evaluated (input/output)
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_OR(ByVal strHex1 As String, _
                          ByVal strHex2 As String) As String

    Dim lngPos    As Long
    Dim lngValue  As Long
    Dim lngIndex  As Long
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64Hex_OR"

    On Error GoTo w64Hex_OR_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify input data is same length
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Please note that it is faster to load an output string
    ' with spaces and then insert the data.  Constantly
    ' appending data will cause the system to have to
    ' reallocate space, re-evaluate new string size and
    ' possibly designate a new internal address to hold this
    ' data.  All of which takes extra cycles and time.
    strOutput = Space$(24)   ' Extra spaces
    lngPos = 1               ' Starting position in output string

    ' Perform "Or" binary comparisons
    For lngIndex = 1 To HEX_LENGTH Step 2

        lngValue = ("&H" & Mid$(strHex1, lngIndex, 2)) Or _
                   ("&H" & Mid$(strHex2, lngIndex, 2))

        Mid$(strOutput, lngPos, 2) = Right$("0" & Hex$(lngValue), 2)
        lngPos = lngPos + 2   ' Increment position pointer

    Next lngIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = UCase$(TrimStr(strOutput))   ' Return hex string in uppercase
    End If

w64Hex_OR_CleanUp:
    w64Hex_OR = strOutput  ' Return new value
    On Error GoTo 0        ' Nullify this error trap
    Exit Function

w64Hex_OR_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_OR_CleanUp

End Function

' ***************************************************************************
' Routine:       w64Hex_XoR
'
' Description:   The Xor operator performs as both a logical and bitwise
'                operator. A bit-wise comparison of two expressions using
'                exclusive-or logic to form the result, as shown in the
'                following example:
'
'                         X = 1001 XoR 1101
'
'                         1001
'                         1101
'                        ------
'                     X = 0100   Same values become "0"
'
' Special Note:  This is the same as the Visual Basic bit function XoR except
'                that it works with 64-bit hex strings.
'
' Input:    428a2f98d728ae22  0100001010001010001011111001100011010111001010001010111000100010
'           27b70a8546d22ffc  0010011110110111000010101000010101000110110100100010111111111100
'          -----------------  ----------------------------------------------------------------
' Output:   653d251d91fa81de  0110010100111101001001010001110110010001111110101000000111011110
'
' Parameters:    strHex1 - Hex string to be evaluated (input/output)
'                strHex2 - Hex string to be evaluated (input/output)
'
' Returns:       New hex data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function w64Hex_XoR(ByVal strHex1 As String, _
                           ByVal strHex2 As String) As String

    Dim lngPos    As Long
    Dim lngValue  As Long
    Dim lngIndex  As Long
    Dim strOutput As String

    Const ROUTINE_NAME As String = "w64Hex_XoR"

    On Error GoTo w64Hex_XoR_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify input data is same length
    strHex1 = Right$(ZEROES_16 & strHex1, HEX_LENGTH)
    strHex2 = Right$(ZEROES_16 & strHex2, HEX_LENGTH)

    ' Please note that it is faster to load an output string
    ' with spaces and then insert the data.  Constantly
    ' appending data will cause the system to have to
    ' reallocate space, re-evaluate new string size and
    ' possibly designate a new internal address to hold this
    ' data.  All of which takes extra cycles and time.
    strOutput = Space$(24)   ' Extra spaces
    lngPos = 1               ' Starting position in output string

    ' Perform "XoR" binary comparisons
    For lngIndex = 1 To HEX_LENGTH Step 2

        lngValue = ("&H" & Mid$(strHex1, lngIndex, 2)) Xor _
                   ("&H" & Mid$(strHex2, lngIndex, 2))

        Mid$(strOutput, lngPos, 2) = Right$("0" & Hex$(lngValue), 2)
        lngPos = lngPos + 2   ' Increment position pointer

    Next lngIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strOutput = vbNullString
    Else
        strOutput = UCase$(TrimStr(strOutput))   ' Return hex string in uppercase
    End If

w64Hex_XoR_CleanUp:
    w64Hex_XoR = strOutput  ' Return new value
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

w64Hex_XoR_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strOutput = vbNullString
    gblnStopProcessing = True
    Resume w64Hex_XoR_CleanUp

End Function

' ***************************************************************************
' Routine:       Shift_64 (Used for 64-bit computations only)
'
' Description:  Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift do not wrap around.  Fill bits
'                coming in from the opposite side are zeros.  Some common
'                languages like C/C++ or Java have an operator for this
'                job: ">>" and "<<".
'
' Parameters:    lngLeftVal  - value to be manipulated (input)
'                lngRightVal - value to be manipulated (input)
'                lngBitShift - Number of shift positions
'                              Positive value = left shift
'                              Negative value = right shift
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Shift_64(ByRef lngLeftVal As Long, _
                    ByRef lngRightVal As Long, _
                    ByVal lngBitShift As Long)

    Dim strHex As String

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Join long integers and convert to hex string
    strHex = Right$(ZEROES_8 & Hex$(lngLeftVal), 8) & _
             Right$(ZEROES_8 & Hex$(lngRightVal), 8)

    ' Call w64Shift() routine
    strHex = w64Shift(strHex, lngBitShift)

    ' Convert back to long integers
    lngLeftVal = ("&H" & Mid$(strHex, 1, 8))
    lngRightVal = ("&H" & Mid$(strHex, 9, 8))

End Sub

' ***************************************************************************
' Routine:       Rotate_64 (Used for 64-bit computations only)
'
' Description:   Shifts the bits to the right/Left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the direction of the shift wrap around to the opposite end.
'                Some common languages like C/C++ or Java have an operator
'                for this job: ">>>" or "<<<".  Sometimes referred to as
'                circular bit shifting or bit rotation.
'
' Parameters:    lngLeftVal  - value to be manipulated
'                lngRightVal - value to be manipulated
'                lngBitShift - Number of shift positions
'                              Positive value = left shift
'                              Negative value = right shift
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Rotate_64(ByRef lngLeftVal As Long, _
                     ByRef lngRightVal As Long, _
                     ByVal lngBitShift As Long)

    Dim strHex As String

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Join long integers and convert to hex string
    strHex = Right$(ZEROES_8 & Hex$(lngLeftVal), 8) & _
             Right$(ZEROES_8 & Hex$(lngRightVal), 8)

    ' Call w64Rotate() routine
    strHex = w64Rotate(strHex, lngBitShift)

    ' Convert back to long integers
    lngLeftVal = ("&H" & Mid$(strHex, 1, 8))
    lngRightVal = ("&H" & Mid$(strHex, 9, 8))

End Sub

' ***************************************************************************
' Routine:       Xor_64 (Used for 64-bit computations only)
'
' Description:   The Xor operator performs as both a logical and bitwise
'                operator. A bit-wise comparison of two expressions using
'                exclusive-or logic to form the result.
'
' Parameters:    lngReturn1  - Manipulated value1 (input/output)
'                lngReturn2  - Manipulated value2 (input/output)
'                lngLeftVal  - value to be compared (input)
'                lngRightVal - value to be compared (input)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Xor_64(ByRef lngReturn1 As Long, _
                  ByRef lngReturn2 As Long, _
                  ByVal lngLeftVal As Long, _
                  ByVal lngRightVal As Long)

    lngReturn1 = lngReturn1 Xor lngLeftVal
    lngReturn2 = lngReturn2 Xor lngRightVal

End Sub

' ***************************************************************************
' Routine:       Or_64 (Used for 64-bit computations only)
'
' Description:   The Or operator performs as both a logical and bitwise
'                operator. A bit-wise comparison of two expressions using
'                exclusive-or logic to form the result.
'
' Parameters:    lngReturn1  - Manipulated value1 (input/output)
'                lngReturn2  - Manipulated value2 (input/output)
'                lngLeftVal  - value to be compared (input)
'                lngRightVal - value to be compared (input)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Or_64(ByRef lngReturn1 As Long, _
                 ByRef lngReturn2 As Long, _
                 ByVal lngLeftVal As Long, _
                 ByVal lngRightVal As Long)

    lngReturn1 = lngReturn1 Or lngLeftVal
    lngReturn2 = lngReturn2 Or lngRightVal

End Sub

' ***************************************************************************
' Routine:       And_64 (Used for 64-bit computations only)
'
' Description:   The And operator performs as both a logical and bitwise
'                operator. A bit-wise comparison of two expressions using
'                exclusive-or logic to form the result.
'
' Parameters:    lngReturn1  - Manipulated value1 (input/output)
'                lngReturn2  - Manipulated value2 (input/output)
'                lngLeftVal  - value to be compared (input)
'                lngRightVal - value to be compared (input)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub And_64(ByRef lngReturn1 As Long, _
                  ByRef lngReturn2 As Long, _
                  ByVal lngLeftVal As Long, _
                  ByVal lngRightVal As Long)

    lngReturn1 = lngReturn1 And lngLeftVal
    lngReturn2 = lngReturn2 And lngRightVal

End Sub

' ***************************************************************************
' Routine:       Not_64 (Used for 64-bit computations only)
'
' Description:   NOT expression used for bit comparison does not exist
'                in Visual Basic.  This function looks at the binary
'                representation of the values of the expression and does
'                a bitwise negation operation on it.
'
' Parameters:    lngReturn1 - Manipulated value1 (input/output)
'                lngReturn2 - Manipulated value2 (input/output)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Not_64(ByRef lngReturn1 As Long, _
                  ByRef lngReturn2 As Long)

    lngReturn1 = Not lngReturn1
    lngReturn2 = Not lngReturn2

End Sub

' ***************************************************************************
' Routine:       Add_64 (Used for 64-bit computations only)
'
' Description:   Perform Hex addition on two 16 char hex values that have
'                been convert from four long integers.
'
'                            Long         Hex
'                Input:   3049323471   b5c0fbcf    Left side
'                         3964484399   ec4d3b2f    Right Side
'
'                         1116352408   428a2f98    Left side
'                         3609767458   d728ae22    Right side
'                         ----------   --------
'                Output:  4165675880   f84b2b68    Left side
'                         3279284561   c375e951    Right side
'
' Parameters:    lngLeftVal  - left side value (input/output)
'                lngRightVal - right side value (input/output)
'                lngTmpLeft  - Value to be added (input)
'                lngTmpRight - Value to be added (input)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Add_64(ByRef lngLeftVal As Long, _
                  ByRef lngRightVal As Long, _
                  ByVal lngTmpLeft As Long, _
                  ByVal lngTmpRight As Long)

    Dim strHex1   As String
    Dim strHex2   As String
    Dim strOutput As String

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Convert longs to hex and concatenate hex strings
    strHex1 = Right$(ZEROES_8 & Hex$(lngLeftVal), 8) & _
              Right$(ZEROES_8 & Hex$(lngRightVal), 8)

    strHex2 = Right$(ZEROES_8 & Hex$(lngTmpLeft), 8) & _
              Right$(ZEROES_8 & Hex$(lngTmpRight), 8)

    strOutput = w64HexAdd(strHex1, strHex2)                 ' Perform hex addition
    strOutput = UCase$(Right$(ZEROES_16 & strOutput, 16))   ' Format full 16 characters
    lngLeftVal = ("&H" & Left$(strOutput, 8))               ' Split and convert to long integers
    lngRightVal = ("&H" & Right$(strOutput, 8))

End Sub

' ***************************************************************************
' Routine:       Subtract_64 (Used for 64-bit computations only)
'
' Description:   Perform Hex subtraction on two 16 char hex values (64-bit)
'                that have been converted from four long integers.
'
'                            Long         Hex
'                Input:   3049323471   b5c0fbcf    Left side
'                         3964484399   ec4d3b2f    Right Side
'
'                         1116352408   428a2f98    Left side
'                         3609767458   d728ae22    Right side
'                         ----------   --------
'                Output:  7336CC37    1932971063   Left side
'                         15248D0D    354716941    Right side
'
' Parameters:    lngLeftVal  - left side value (input/output)
'                lngRightVal - right side value (input/output)
'                lngTmpLeft  - Value to be added (input)
'                lngTmpRight - Value to be added (input)
'
' Returns:       Two long integers
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-Apr-2013  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Subtract_64(ByRef lngLeftVal As Long, _
                       ByRef lngRightVal As Long, _
                       ByVal lngTmpLeft As Long, _
                       ByVal lngTmpRight As Long)

    Dim strHex1   As String
    Dim strHex2   As String
    Dim strOutput As String

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    ' Convert longs to hex and concatenate hex strings
    strHex1 = Right$(ZEROES_8 & Hex$(lngLeftVal), 8) & _
              Right$(ZEROES_8 & Hex$(lngRightVal), 8)

    strHex2 = Right$(ZEROES_8 & Hex$(lngTmpLeft), 8) & _
              Right$(ZEROES_8 & Hex$(lngTmpRight), 8)

    strOutput = w64HexSubtract(strHex1, strHex2)            ' Perform hex subtraction
    strOutput = UCase$(Right$(ZEROES_16 & strOutput, 16))   ' Format full 16 characters
    lngLeftVal = ("&H" & Left$(strOutput, 8))               ' Split and convert to long integers
    lngRightVal = ("&H" & Right$(strOutput, 8))

End Sub

' ***************************************************************************
' Routine:       Multiply_64
'
' Description:   Perform Hex multiplication on two 16 char hex values (64-bit).
'
'                                Hex
'                Input:     428a2f98d728ae22
'                           b5c0fbcfec4d3b2f
'                           ----------------
'                Output:    2f3ddcbe4a93effec7c951e798d3ce3e  Complete output
'                Return:    2f3ddcbe4a93effe                  Return first 16 chars only
'
' Parameters:    strHex1 - First hex string (64-bit) (Multiplicand)
'                strHex2 - Second hex string (64-bit)(Multiplier)
'
' Returns:       Left-most 16 chars of newly calculated hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-May-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Sub Multiply_64(ByRef lngLeftVal As Long, _
                       ByRef lngRightVal As Long, _
                       ByVal lngTmpLeft As Long, _
                       ByVal lngTmpRight As Long)

    Dim strHex1   As String
    Dim strHex2   As String
    Dim strOutput As String

    strOutput = vbNullString

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo Multiply_64_CleanUp
    End If

    ' Convert longs to hex and concatenate hex strings
    strHex1 = Right$(ZEROES_8 & Hex$(lngLeftVal), 8) & _
              Right$(ZEROES_8 & Hex$(lngRightVal), 8)

    strHex2 = Right$(ZEROES_8 & Hex$(lngTmpLeft), 8) & _
              Right$(ZEROES_8 & Hex$(lngTmpRight), 8)

    strOutput = w64HexMultiply(strHex1, strHex2)


Multiply_64_CleanUp:
    ' Convert to long integers
    lngLeftVal = ("&H" & Right$(ZEROES_8 & Mid$(strOutput, 1, 8), 8))
    lngRightVal = ("&H" & Right$(ZEROES_8 & Mid$(strOutput, 9, 8), 8))

End Sub

' ***************************************************************************
' Routine:       BigAdd
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Adds strNumber2 to strNumber1 and
'                returns the result in a string.
'
' Example:       13096744586834688815
'                 4794697086780616226
'                --------------------
'                17891441673615305041
'
' Parameters:    strNumber1 - Numeric string
'                strNumber1 - Numeric string to be added to strNumber1
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Oct-2005  HAANDI
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62738&lngWId=1
' 29-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function BigAdd(ByVal strNumber1 As String, _
                       ByVal strNumber2 As String) As String

    Dim intIndex   As Integer
    Dim intLengthA As Integer
    Dim intLengthB As Integer
    Dim intLengthT As Integer
    Dim intCarry   As Integer
    Dim dblTemp    As Double
    Dim strTemp    As String
    Dim strResult  As String

    Const ROUTINE_NAME As String = "BigAdd"

    On Error GoTo BigAdd_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    strResult = vbNullString
    strNumber1 = TrimStr(strNumber1)
    strNumber2 = TrimStr(strNumber2)

    If strNumber2 = "0" Or Len(strNumber2) = 0 Then
        strResult = strNumber1
        GoTo BigAdd_CleanUp
    End If

    If strNumber1 = "0" Or Len(strNumber1) = 0 Then
        strResult = strNumber2
        GoTo BigAdd_CleanUp
    End If

    intLengthA = Len(strNumber1)
    intLengthB = Len(strNumber2)

    If intLengthA >= intLengthB Then

        strNumber1 = String$(14 - (intLengthA Mod 14), "0") & strNumber1
        intLengthA = Len(strNumber1)

        strNumber2 = String$(intLengthA - intLengthB, "0") & strNumber2
        intLengthB = intLengthA

    Else

        strNumber2 = String$(14 - (intLengthB Mod 14), "0") & strNumber2
        intLengthB = Len(strNumber2)

        strNumber1 = String$(intLengthB - intLengthA, "0") & strNumber1
        intLengthA = intLengthB

    End If

    For intIndex = intLengthA - 13 To 1 Step -14

        dblTemp = Int(Mid$(strNumber1, intIndex, 14)) + Int(Mid$(strNumber2, intIndex, 14)) + intCarry
        strTemp = LTrim$(Str$(dblTemp))
        intLengthT = Len(strTemp)
        intCarry = 0

        If intLengthT < 14 Then
            strTemp = String$(14 - intLengthT, "0") & strTemp
        End If

        strResult = Right$(strTemp, 14) & strResult

        If intLengthT = 15 Then
            intCarry = 1
        End If

    Next intIndex

    If intCarry Then
        strResult = intCarry & strResult
    End If

BigAdd_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        BigAdd = "0"
    Else
        BigAdd = DataTrim(strResult)
    End If

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigAdd_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strResult = "0"
    gblnStopProcessing = True
    Resume BigAdd_CleanUp

End Function

' ***************************************************************************
' Routine:       BigSubtract
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Subtracts strNumber2 from strNumber1
'                and returns the result in a string.
'
' Example:       13096744586834688815
'                 4794697086780616226
'                --------------------
'                 8302047500054072589
'
' Parameters:    strNumber1 - Numeric string
'                strNumber1 - Numeric string to be Subtracted
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Oct-2005  HAANDI
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62738&lngWId=1
' 29-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function BigSubtract(ByVal strNumber1 As String, _
                            ByVal strNumber2 As String) As String

    Dim intIndex   As Integer
    Dim intLengthA As Integer
    Dim intLengthB As Integer
    Dim intLengthT As Integer
    Dim intCarry   As Integer
    Dim dblTemp    As Double
    Dim strTemp    As String
    Dim strResult  As String

    Const ROUTINE_NAME As String = "BigSubtract"

    On Error GoTo BigSubtract_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    strResult = vbNullString
    strNumber1 = TrimStr(strNumber1)
    strNumber2 = TrimStr(strNumber2)

    If strNumber2 = "0" Or Len(strNumber2) = 0 Then
        strResult = strNumber1
        GoTo BigSubtract_CleanUp
    End If

    intLengthA = Len(strNumber1)
    intLengthB = Len(strNumber2)

    strNumber1 = String$(14 - (intLengthA Mod 14), "0") & strNumber1
    intLengthA = Len(strNumber1)

    strNumber2 = String$(intLengthA - intLengthB, "0") & strNumber2
    intLengthB = intLengthA

    For intIndex = (intLengthA - 13) To 1 Step -14

        dblTemp = CDbl(Mid$(strNumber1, intIndex, 14)) - Val(Mid$(strNumber2, intIndex, 14)) - intCarry

        If dblTemp < 0 Then
            intCarry = 1
            dblTemp = dblTemp + 100000000000000#
        Else
            intCarry = 0
        End If

        strTemp = LTrim$(Str$(dblTemp))
        intLengthT = Len(strTemp)
        strTemp = String$(14 - intLengthT, "0") & strTemp
        strResult = Right$(strTemp, 14) & strResult

    Next intIndex

BigSubtract_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        BigSubtract = "0"
    Else
        BigSubtract = DataTrim(strResult)
    End If

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigSubtract_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strResult = "0"
    gblnStopProcessing = True
    Resume BigSubtract_CleanUp

End Function

' ***************************************************************************
' Routine:       BigMultiply
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Multiplies strNumber1 with strNumber2
'                and returns the result in a string.
'
' Example:                       13096744586834688815
'                                 4794697086780616226
'              --------------------------------------
'              62794923116806087757258117158649712190
'
' Parameters:    strNumber1 - Numeric string as multiplicand
'                strNumber1 - Numeric string as multiplier
'
' Returns:       New numeric string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Oct-2005  HAANDI
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62738&lngWId=1
' 29-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function BigMultiply(ByVal strNumber1 As String, _
                            ByVal strNumber2 As String) As String

    Dim intIdx      As Integer
    Dim intIndex    As Integer
    Dim intLengthA  As Integer
    Dim intLengthB  As Integer
    Dim intLengthT  As Integer
    Dim intCarry    As Integer
    Dim dblTemp     As Double
    Dim strTemp     As String
    Dim strResult   As String
    Dim astrData(9) As String

    Const ROUTINE_NAME As String = "BigMultiply"

    On Error GoTo BigMultiply_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    strResult = "0"
    strNumber1 = TrimStr(strNumber1)
    strNumber2 = TrimStr(strNumber2)

    If strNumber2 = "0" Or Len(strNumber2) = 0 Then
        strResult = "0"
        GoTo BigMultiply_CleanUp
    End If

    If strNumber1 = "0" Or Len(strNumber1) = 0 Then
        strResult = "0"
        GoTo BigMultiply_CleanUp
    End If

    'intLengthA = Len(strNumber1)
    strNumber1 = String$(14 - (intLengthA Mod 14), "0") & strNumber1

    intLengthA = Len(strNumber1)
    intLengthB = Len(strNumber2)

    For intIdx = 1 To 9

        For intIndex = (intLengthA - 13) To 1 Step -14

            dblTemp = CDbl(Mid$(strNumber1, intIndex, 14)) * intIdx + intCarry
            strTemp = LTrim$(Str$(dblTemp))
            intLengthT = Len(strTemp)
            intCarry = 0

            If intLengthT < 14 Then
                strTemp = String$(14 - intLengthT, "0") & strTemp
            End If

            astrData(intIdx) = Right$(strTemp, 14) & astrData(intIdx)

            If intLengthT = 15 Then
                intCarry = CInt(Left$(strTemp, 1))
            End If

        Next intIndex

        If intCarry Then
            astrData(intIdx) = intCarry & astrData(intIdx)
            intCarry = 0
        End If

    Next intIdx

    For intIdx = intLengthB To 1 Step -1

        If Mid$(strNumber2, intIdx, 1) <> "0" Then
            strResult = BigAdd(strResult, astrData(Int(Mid$(strNumber2, intIdx, 1))) & String$(intLengthB - intIdx, "0"))
        End If

    Next intIdx

BigMultiply_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        BigMultiply = "0"
    Else
        BigMultiply = DataTrim(strResult)
    End If


    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigMultiply_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strResult = "0"
    gblnStopProcessing = True
    Resume BigMultiply_CleanUp

End Function

' ***************************************************************************
' Routine:       BigDivide
'
' Description:   A solution for handling large numbers (up to 32768 digits)
'                in Visual Basic 6.  Divides strDividend by strDivisor and
'                returns truncated result (no decimals) in a string.
'
' Example:       Full return with decimals:
'                1060894660739950.4912920210611584 = 13096744586834688815 / 12345
'
'                Only the truncated value is returned in this routine
'                1060894660739950 = 13096744586834688815 \ 12345
'
' Parameters:    strDividend - Numeric string as dividend
'                strDividend - Numeric string as divisor
'
' Returns:       New numeric string as quotient
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-Oct-2005  HAANDI
'              http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62738&lngWId=1
' 29-May-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function BigDivide(ByVal strDividend As String, _
                          ByVal strDivisor As String) As String

    Dim strCurrent  As String
    Dim strTemp     As String
    Dim strCount    As String
    Dim strQuotient As String
    Dim intLength   As Integer
    Dim lngValue    As Long

    Const ROUTINE_NAME As String = "BigDivide"

    On Error GoTo BigDivide_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    If Not IsNumeric(strDividend) Or Not IsNumeric(strDivisor) Then
        InfoMsg "Values to be divided must be numeric." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    If Val(strDividend) <= 0 Or Val(strDivisor) <= 0 Then
        InfoMsg "Both values must be greater than zero." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        Exit Function
    End If

    strCurrent = strDividend
    strQuotient = vbNullString

    If Val(strDivisor) > Val(strDividend) Then
        strQuotient = "0"
        GoTo BigDivide_CleanUp
    End If

    intLength = Len(strDivisor)

    Do While Val(strDivisor) < Val(strCurrent)

        lngValue = Len(strCurrent) - intLength + _
                   CBool(Val(strDivisor) > Val(Left$(strCurrent, intLength)))

        strTemp = strDivisor & String$(lngValue, "0")
        strCount = "1" & String$(lngValue, "0")

        Do While (Val(strCurrent) > Val(strTemp)) Or strCurrent = strTemp

            strCurrent = BigSubtract(strCurrent, strTemp)
            strQuotient = BigAdd(strQuotient, strCount)

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

        Loop

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            Exit Do   ' exit Do..Loop
        End If

        strCurrent = DataTrim(strCurrent)

    Loop

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strQuotient = "0"
    End If

BigDivide_CleanUp:
    BigDivide = strQuotient

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

BigDivide_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strQuotient = "0"
    gblnStopProcessing = True
    Resume BigDivide_CleanUp

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

' ***************************************************************************
' Routine:       SwapEndian
'
' Description:   Converts hex data from BIG_ENDIAN to LITTLE_ENDIAN or
'                LITTLE_ENDIAN to BIG_ENDIAN format.
'
' Parameters:    vntData - Hex data to be manipulated
'                lngRetLength - Optional - Desired length of the return
'                         string.  Default = 16
'
' Returns:       Hex string in new format
'
' Example:         LITTLE_ENDIAN  to   BIG_ENDIAN
'                0123456789ABCDEF    EFCDAB8967452301
'
'                  BIG_ENDIAN     to   LITTLE_ENDIAN
'                EFCDAB8967452301    0123456789ABCDEF
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
                  Optional ByVal lngRetLength As Long = 16) As String

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
' Routine:       HexRepresentation
'
' Description:   This is where the hex representation of an array of numbers
'                is calculated and returned in a string array.  If the
'                parameter lngSpecificNumber is greater than zero then only a single
'                value is processed.  If the parameter lngSpecificNumber equal zero
'                then an array of prime numbers are determined in sequence
'                and their hex representation is calculated accordingly.
'
'                These types of hex values are generally used in the work
'                and constant arrays for the SHA2 family of hash algorithms.
'
' Parameters:    lngQtyReq - Number of items to be returned
'                lngRetFmt - Optional - Return format.  Return all 16 hex
'                            values, just the first 8 or the last 8.
'                            Default = e16Chars (All 16 characters)
'                lngHexRep - Optional - Type of hex representation to
'                            perform.  Either Cube Roots or Square Roots.
'                            Default = eCubeRoots
'                lngSpecificNumber - Optional - This value is greater than zero
'                            if a single value is to be processed.
'                            Default = 0
'                blngDebugPrint - Optional - Only recognized if lngSpecificNumber
'                            is greater than zero. Used to print detailed
'                            data about number in the immediate window.
'                            To see this, press CTRL+G to open the window
'                            after processing has completed.
'
' Returns:       A string array of hex values
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added parameter testing
' 17-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated to allow print detailed structure to VB6 IDE Immediate
'              Window of a specific Prime Number.
' ***************************************************************************
Public Function HexRepresentation(ByVal lngQtyReq As Long, _
                         Optional ByVal lngRetFmt As enumHEX_RETURN_FORMAT = e16Chars, _
                         Optional ByVal lngHexRep As enumHEX_REPRESENTATION = eCubeRoots, _
                         Optional ByVal lngSpecificNumber As Long = 0, _
                         Optional ByVal blnDebugPrint As Boolean = False) As String()

    Dim lngIndex      As Long
    Dim lngPrimePos   As Long
    Dim astrHex()     As String
    Dim alngNumbers() As Long

    Const ROUTINE_NAME As String = "HexRepresentation"

    On Error GoTo HexRepresentation_Error

    Erase astrHex()     ' Always start with empty arrays
    Erase alngNumbers()
    lngPrimePos = 0

    ' If a negative value is passed
    ' then present a message
    If lngQtyReq < 1 Then
        InfoMsg "Quantity requested must be a positive value greater than zero." & _
                vbNewLine & vbNewLine & "Quantity:   " & CStr(lngQtyReq) & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo HexRepresentation_CleanUp
    End If

    Select Case lngRetFmt
           Case e16Chars, eLeft8, eRight8
                ' Valid values
           Case Else
                InfoMsg "Cannot identify number of characters to be returned." & _
                        vbNewLine & vbNewLine & "Valid values:" & vbNewLine & _
                        Space$(5) & "e16Chars or 16" & vbNewLine & _
                        Space$(5) & "eLeft8 or 8" & vbNewLine & _
                        Space$(5) & "eRight8 or 8" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
                GoTo HexRepresentation_CleanUp
    End Select

    Select Case lngHexRep
           Case eCubeRoots, eSquareRoots
                ' Valid values
           Case Else
                InfoMsg "Hex representation must be either CUBE ROOT or SQUARE ROOT." & _
                        vbNewLine & vbNewLine & "Valid values:" & vbNewLine & _
                        Space$(5) & "eCubeRoots or 0" & vbNewLine & _
                        Space$(5) & "eSquareRoots or 1" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
                GoTo HexRepresentation_CleanUp
    End Select

    ' Evaluate if one or more numbers are to be processed
    Select Case lngSpecificNumber

           Case Is > 0
                ' Use an array of prime numbers
                '
                ' load prime number array
                alngNumbers() = GetPrimeNumbers(lngQtyReq)

                ' Test to see if the array has any data
                If Not CBool(IsArrayInitialized(alngNumbers())) Then
                    InfoMsg "Failed to load prime number array." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                    GoTo HexRepresentation_CleanUp
                End If

                For lngIndex = 0 To UBound(alngNumbers) - 1
                    If lngSpecificNumber = alngNumbers(lngIndex) Then
                        lngPrimePos = lngIndex + 1
                        Exit For
                    End If
                Next lngIndex

                ' lngSpecificNumber is greater than zero therefore
                ' a single value is to be processed
                lngQtyReq = 1
                ReDim alngNumbers(lngQtyReq)
                CopyMemory alngNumbers(0), lngSpecificNumber, 4&

           Case 0
                ' Use an array of prime numbers
                '
                ' load prime number array
                alngNumbers() = GetPrimeNumbers(lngQtyReq)

                ' Test to see if the array has any data
                If Not CBool(IsArrayInitialized(alngNumbers())) Then
                    InfoMsg "Failed to load prime number array." & _
                            vbNewLine & vbNewLine & _
                            "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                    GoTo HexRepresentation_CleanUp
                End If

           Case Is < 0
                InfoMsg "Number to be evaluated must be a positive value greater than zero." & _
                        vbNewLine & vbNewLine & "Number:   " & CStr(lngSpecificNumber) & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                        gblnStopProcessing = True
                        GoTo HexRepresentation_CleanUp
    End Select

    ' Fill hex array with computed values
    Select Case lngHexRep

           Case eCubeRoots
                astrHex() = LoadCubeRoots(lngQtyReq, alngNumbers(), lngRetFmt, lngSpecificNumber, blnDebugPrint, lngPrimePos)

           Case eSquareRoots
                astrHex() = LoadSquareRoots(lngQtyReq, alngNumbers(), lngRetFmt, lngSpecificNumber, blnDebugPrint, lngPrimePos)
    End Select

    HexRepresentation = astrHex()

HexRepresentation_CleanUp:
    Erase astrHex()     ' Always empty arrays when not needed
    Erase alngNumbers()

    On Error GoTo 0  ' Nullify this error trap
    Exit Function

HexRepresentation_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume HexRepresentation_CleanUp

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
'                in Modified form, is still an important tool in number
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
' Parameters:    lngQtyNeeded - Optional - Number of prime numbers to return
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

    Erase alngPrimes()   ' Always start with empty arrays
    lngPrimeCnt = 0      ' Init array index

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
        ReDim Preserve alngPrimes(lngQtyNeeded)   ' Resize prime number array
        GetPrimeNumbers = alngPrimes()            ' Return prime number array
        GoTo GetPrimeNumbers_CleanUp              ' Time to leave
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

    ReDim Preserve alngPrimes(lngQtyNeeded)   ' Resize prime number array
    GetPrimeNumbers = alngPrimes()            ' Return prime number array

GetPrimeNumbers_CleanUp:
    Erase alngPrimes()   ' Always empty arrays when not needed
    On Error GoTo 0      ' Nullify this error trap
    Exit Function

GetPrimeNumbers_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
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
'                in Modified form, is still an important tool in number
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

    IsPrime = False              ' Preset to FALSE for input testing
    dblNumber = Fix(dblNumber)   ' Remove any decimal positions

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

' ***************************************************************************
' Routine:       ConvertBase
'
' Description:   These routines will handle both positive/negative values.
'                When making conversions, always double check your work.
'                I searched the web for various routines that would handle
'                both positive and negative values.  Believe me when I say
'                this is like looking for hen's teeth.  Some I found and
'                made some modifications.  Others, I wrote.  Still others,
'                I do not remember where I found them.  If one of the
'                conversion routines is yours, let me know and I will update
'                the documentation.  Be sure to include the URL where your
'                original code is located.  Thank you.
'
'                Bases used here:
'                    2 (Binary)  -> 10 (Numeric) or 16 (Hex)
'                   10 (Numeric) ->  2 (Binary)  or 16 (Hex)
'                   16 (Hex)     ->  2 (Binary)  or 10 (Numeric)
'
'                Correct values for these routines are tested by inputing the
'                output values from a previous test and producing the same
'                results.
'
'                Number:  -18446744073709551615                                             (Min value allowed)
'                   Hex:  FFFFFFFFFFFFFFFF                                                  (16 chars)
'                Binary:  1111111111111111111111111111111111111111111111111111111111111111  (64 chars)
'
'                Number:  9223372036854775807                                               (Max value allowed)
'                   Hex:  7FFFFFFFFFFFFFFF                                                  (16 chars)
'                Binary:  0111111111111111111111111111111111111111111111111111111111111111  (64 chars)
'
' ===========================================================================
'
' Parameters:    strInput    - Data to be converted
'                lngFromBase - Input data base (see above)
'                lngToBase   - Output data base (see above)
'
' Returns:       Converted data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 26-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote all called routines
' 11-Jan-2016  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug.  Invalid evaluation of large input number.
' ***************************************************************************
Private Function ConvertBase(ByVal strInput As String, _
                             ByVal lngFromBase As enum64_BaseValues, _
                             ByVal lngToBase As enum64_BaseValues) As String

    ' Called by w64NumberToHex()
    '           w64NumberToBinary()
    '           w64HexToNumber()
    '           w64HexToBinary()
    '           w64BinaryToHex()
    '           w64BinaryToNumber()

    Dim strData As String

    On Error GoTo ConvertBase_Error

    ConvertBase = vbNullString   ' Empty return string

    If gblnStopProcessing Then
        GoTo ConvertBase_CleanUp
    End If

    strInput = UCase$(TrimStr(strInput))   ' Format input data
    strData = strInput                     ' Copy input data here for restart/debug purposes

    Select Case lngFromBase

           Case e64_Binary   ' Base 2
                ' Verify a full 64 character string
                strData = Right$(String$(BIN_LENGTH, "0") & strData, BIN_LENGTH)

                ' Verify this is binary data
                If Not IsBinaryData(strData) Then
                    gblnStopProcessing = True   ' Set flag to stop processing
                    GoTo ConvertBase_CleanUp
                End If

                Select Case lngToBase
                       Case e64_Numeric: ConvertBase = Base2_to_Base10(strData)   ' Base 2 -> 10
                       Case e64_Hex:     ConvertBase = Base2_to_Base16(strData)   ' Base 2 -> 16
                End Select

           Case e64_Numeric   ' Base 10
                ' Verify this is numeric data
                If Not IsNumericData(strData) Then
                    gblnStopProcessing = True   ' Set flag to stop processing
                    GoTo ConvertBase_CleanUp
                End If

                ' Input numeric comparison
                If (CDec(strData) >= CDec(MIN_RANGE)) And _
                   (CDec(strData) <= CDec(MAX_RANGE)) Then

                    Select Case lngToBase
                           Case e64_Binary: ConvertBase = Base10_to_Base2(strData)    ' Base 10 -> 2
                           Case e64_Hex:    ConvertBase = Base10_to_Base16(strData)   ' Base 10 -> 16
                    End Select
                Else
                    InfoMsg "Numeric data to be converted exceeds" & vbNewLine & _
                            "acceptable range limitations." & vbNewLine & vbNewLine & _
                            "Input data:" & "  " & strData & vbNewLine & vbNewLine & _
                            "Min range:" & "   " & MIN_RANGE & vbNewLine & _
                            "Max range:" & "   " & MAX_RANGE
                    GoTo ConvertBase_Error
                End If

           Case e64_Hex   ' Base 16
                ' Verify a full 16 character string
                strData = Right$(String$(HEX_LENGTH, "0") & strData, HEX_LENGTH)

                ' Verify this is hex data
                If Not IsHexData(strData) Then
                    gblnStopProcessing = True   ' Set flag to stop processing
                    GoTo ConvertBase_CleanUp
                End If

                Select Case lngToBase
                       Case e64_Binary:  ConvertBase = Base16_to_Base2(strData)    ' Base 16 -> 2
                       Case e64_Numeric: ConvertBase = Base16_to_Base10(strData)   ' Base 16 - > 10
                End Select
    End Select

ConvertBase_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ConvertBase = vbNullString
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

ConvertBase_Error:
    Err.Clear                   ' Remove error code
    gblnStopProcessing = True   ' Set flag to stop processing
    Resume ConvertBase_CleanUp

End Function

' ***************************************************************************
' Routine:       Base10_to_Base16
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
Private Function Base10_to_Base16(ByVal strNumber As String) As String

    ' Called by ConvertBase()

    Dim strBinary As String

    If gblnStopProcessing Then
        Exit Function
    End If

    Base10_to_Base16 = vbNullString

    If IsNumericData(strNumber) Then

        ' Test number ranges
        If (CDec(strNumber) < CDec(MIN_RANGE)) Or _
           (CDec(strNumber) > CDec(MAX_RANGE)) Then

            InfoMsg "Incoming value exceeds numeric ranges." & vbNewLine & vbNewLine & _
                    "Input value:  " & strNumber & vbNewLine & vbNewLine & _
                    "Min range:  " & MIN_RANGE & vbNewLine & _
                    "Max range:  " & MAX_RANGE

            gblnStopProcessing = True   ' Set flag to stop processing
            Exit Function
        End If

        strBinary = Base10_to_Base2(strNumber)   ' First convert to binary

        If gblnStopProcessing Then
            Exit Function
        End If

        Base10_to_Base16 = Base2_to_Base16(strBinary)    ' Convert binary to hex

    End If

End Function

' ***************************************************************************
' Routine:       Base10_to_Base2
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
' Parameters:    strNumber       - Number to be converted
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
Private Function Base10_to_Base2(ByVal strNumber As String) As String

    ' Called by ConvertBase()
    '           Base10_to_Base16()

    Dim blnNegative As Boolean

    On Error GoTo Base10_to_Base2_Error

    Base10_to_Base2 = vbNullString

    If gblnStopProcessing Then
        GoTo Base10_to_Base2_CleanUp
    End If

    If IsNumericData(strNumber) Then

       ' Test number range
       If (CDec(strNumber) < CDec(MIN_RANGE)) Or _
          (CDec(strNumber) > CDec(MAX_RANGE)) Then

           InfoMsg "Incoming value exceeds numeric ranges." & vbNewLine & vbNewLine & _
                   "Input value:  " & strNumber & vbNewLine & vbNewLine & _
                   "Min allowed:  " & MIN_RANGE & vbNewLine & _
                   "Max allowed:  " & MAX_RANGE

           gblnStopProcessing = True   ' Set flag to stop processing
           GoTo Base10_to_Base2_CleanUp
       End If

       ' See if number is less than zero
       If CDec(strNumber) < 0 Then
           blnNegative = True                 ' for negative numbers
           strNumber = Abs(CDec(strNumber))   ' change to positive number
       Else
           blnNegative = False  ' for positive numbers
       End If

       Do While CDec(strNumber) > 0
           Base10_to_Base2 = Trim$(Str$(CDec(strNumber) - 2 * Int(CDec(strNumber) / 2))) & Base10_to_Base2
           strNumber = Int(CDec(strNumber) / 2)
       Loop

       Base10_to_Base2 = Right$(String$(BIN_LENGTH, "0") & Base10_to_Base2, BIN_LENGTH)

       ' If value is a negative then
       ' verify first position is "1"
       If blnNegative Then
           Mid$(Base10_to_Base2, 1, 1) = "1"
       End If

    End If

Base10_to_Base2_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Base10_to_Base2_Error:
    Err.Clear                        ' Remove error code
    gblnStopProcessing = True        ' Set flag to stop processing
    Base10_to_Base2 = vbNullString   ' Empty return string
    Resume Base10_to_Base2_CleanUp

End Function

' ***************************************************************************
' Routine:       Base16_to_Base2
'
' Description:   Convert a hexadecimal value to a binary string.
'                Please note that it is faster to load an output string
'                with spaces and then insert the data.  Constantly
'                appending data will cause the system to have to
'                reallocate space, re-evaluate new string size and
'                possibly designate a new internal address to hold this
'                data.  All of which takes extra cycles and time.
'
' Parameters:    strHex - Hex value to be converted
'
' Returns:       Binary string representing converted hex value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function Base16_to_Base2(ByVal strHex As String) As String

    ' Called by ConvertBase()
    '           Base16_to_Base10()

    Dim lngPos     As Long
    Dim lngPointer As Long
    Dim strBinary  As String

    On Error GoTo Base16_to_Base2_Error

    If gblnStopProcessing Then
        GoTo Base16_to_Base2_CleanUp
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

Base16_to_Base2_CleanUp:
    Base16_to_Base2 = strBinary   ' Return new value
    On Error GoTo 0               ' Nullify this error trap
    Exit Function

Base16_to_Base2_Error:
    Err.Clear                        ' Remove error code
    gblnStopProcessing = True        ' Set flag to stop processing
    strBinary = vbNullString         ' Empty return string
    Resume Base16_to_Base2_CleanUp

End Function

' ***************************************************************************
' Routine:       Base16_to_Base10
'
' Description:   Convert a hexadecimal value to a numeric value.  Will
'                return positive or negative value.
'
' Parameters:    strHex - Hex value to be converted
'
' Returns:       Numeric value representing converted hex value
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-Aug-2007  Logophobic
'              Convert huge Hex number to Decimal
'              http://www.vbforums.com/showthread.php?483207-Convert-huge-Hex-number-to-Decimal
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function Base16_to_Base10(ByVal strHex As String) As String

    ' Called by ConvertBase()
    '           Base2_to_Base10()

    Dim blnNegative As Boolean
    Dim strBinary   As String
    Dim lngCarry    As Long   ' For carry-over during additions
    Dim lngDigit    As Long
    Dim lngPointer  As Long
    Dim lngMaxDigit As Long   ' To keep track of actual number of decimal digits
    Dim bytTemp     As Byte
    Dim abytData()  As Byte

    Base16_to_Base10 = vbNullString   ' Preset return value
    lngMaxDigit = 0

    If gblnStopProcessing Then
        Exit Function
    End If

    strHex = UCase$(strHex)               ' Convert to uppercase
    strBinary = Base16_to_Base2(strHex)   ' First convert to binary

    ' Determine if incoming value is a negative value
    If Left$(strBinary, 1) = "1" Then
        blnNegative = True
    Else
        blnNegative = False
    End If

    strBinary = vbNullString    ' No longer needed

    ' Make array large enough
    ' to hold all digits
    ReDim abytData(128)

    ' Loop through HEX string
    For lngPointer = 1 To Len(strHex)

        ' Assign appropriate value to Carry
        lngDigit = Asc(Mid$(strHex, lngPointer, 1))

        Select Case lngDigit
               Case 48 To 57 ' 0 to 9
                    lngCarry = lngDigit - 48
               Case 65 To 70 ' A to F
                    lngCarry = lngDigit - 55
               Case 97 To 102 ' a to f
                    lngCarry = lngDigit - 87
        End Select

        ' Multiply each decimal digit by 16 and add previous carry.
        ' This effectively multiplies the entire number by 16
        ' and adds the value of the current hexidecimal digit.
        For lngDigit = 0 To lngMaxDigit
            abytData(lngDigit) = abytData(lngDigit) * 16 + lngCarry
            lngCarry = abytData(lngDigit) \ 10
            abytData(lngDigit) = abytData(lngDigit) Mod 10
        Next lngDigit

        ' Continue until there is nothing to
        ' carry over, adding digits as necessary
        Do Until lngCarry = 0
            lngMaxDigit = lngMaxDigit + 1
            abytData(lngMaxDigit) = lngCarry Mod 10
            lngCarry = lngCarry \ 10
        Loop

    Next lngPointer

    ' Resize array
    If lngMaxDigit < UBound(abytData) Then
        ReDim Preserve abytData(lngMaxDigit)
    End If

    ' Convert digit value to character code
    For lngPointer = 0 To lngMaxDigit
        abytData(lngPointer) = abytData(lngPointer) + 48
    Next lngPointer

    ' Reverse the order of the digits
    lngPointer = 0
    lngDigit = lngMaxDigit

    Do While lngPointer < lngDigit

        bytTemp = abytData(lngPointer)
        abytData(lngPointer) = abytData(lngDigit)
        abytData(lngDigit) = bytTemp
        lngPointer = lngPointer + 1
        lngDigit = lngDigit - 1

    Loop

    ' Convert byte array to string
    Base16_to_Base10 = StrConv(abytData, vbUnicode)

    ' If a negative value then
    ' prefix with minus sign
    If blnNegative Then
        Base16_to_Base10 = "-" & Base16_to_Base10
    End If

End Function

' ***************************************************************************
' Routine:       Base2_to_Base16
'
' Description:   Convert a binary string to hexadecimal format.
'                Please note that it is faster to load an output string
'                with spaces and then insert the data.  Constantly
'                appending data will cause the system to have to
'                reallocate space, re-evaluate new string size and
'                possibly designate a new internal address to hold this
'                data.  All of which takes extra cycles and time.
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
Private Function Base2_to_Base16(ByVal strBinary As String) As String

    ' Called by ConvertBase()
    '           Base10_to_Base16()

    Dim lngPos     As Long
    Dim lngPointer As Long
    Dim strHex     As String
    Dim strBits    As String

    On Error GoTo Base2_to_Base16_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Verify this is binary data
    If Not IsBinaryData(strBinary) Then
        GoTo Base2_to_Base16_CleanUp
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

    strHex = TrimStr(strHex)  ' Remove excess blank spaces

Base2_to_Base16_CleanUp:
    Base2_to_Base16 = strHex   ' Return new value
    On Error GoTo 0            ' Nullify this error trap
    Exit Function

Base2_to_Base16_Error:
    Err.Clear                        ' Remove error code
    gblnStopProcessing = True        ' Set flag to stop processing
    strHex = vbNullString            ' Empty return string
    Resume Base2_to_Base16_CleanUp

End Function

' ***************************************************************************
' Routine:       Base2_to_Base10
'
' Description:   Convert a binary string to a number.  Will return positive
'                or negative value.  I keep the conversions in two steps
'                for easier trouble shooting.
'
' Parameters:    strBinary - Binary value to be converted
'
' Returns:       Numeric value representing converted binary data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 30-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function Base2_to_Base10(ByVal strBinary As String) As String

    ' Called by ConvertBase()

    Dim strHex As String

    If gblnStopProcessing Then
        Exit Function
    End If

    strHex = Base2_to_Base16(strBinary)

    If gblnStopProcessing Then
        Exit Function
    End If

    Base2_to_Base10 = Base16_to_Base10(strHex)

End Function

Private Function DataTrim(ByVal strNumber As String) As String

    ' Drop all leading "0"s
    '
    ' Called by BigAdd()
    '           BigSubtract()
    '           BigMultiply()
    '           BigDivide()

    strNumber = TrimStr(strNumber)   ' Remove unwanted leading/trailing chars

    ' No data to work with
    If Len(strNumber) = 0 Then
        DataTrim = "0"
    Else
        ' Loop thru input number
        ' and remove leading zeroes
        Do While Left$(strNumber, 1) = "0"
            strNumber = Right$(strNumber, Len(strNumber) - 1)
        Loop

        ' Is there any data left?
        If Len(strNumber) = 0 Then
            strNumber = "0"
        End If

        DataTrim = strNumber   ' Return number without leading zeroes
    End If

End Function

' ***************************************************************************
' Routine:       LoadSquareRoots
'
' Description:   Got this explanation from Ask Dr. Math web site.
'                http://mathforum.org/dr.math/ 
'                 
'                There is a simple process for converting a base 10 decimal 
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex will be our new base (16).  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'                 
' Example:   1st Prime number = 2
'            SQUARE ROOT of 2 = 1.4142135623730950488016887242
'                               |
'                               +-- Whole number is dropped
'
'            0.4142135623730950488016887242 * 16 =  6.6274169979695207808270195872
'            0.6274169979695207808270195872 * 16 = 10.038671967512332493232313395
'            0.038671967512332493232313395  * 16 =  0.61875148019731989171701432
'            0.61875148019731989171701432   * 16 =  9.90002368315711826747222912
'            0.90002368315711826747222912   * 16 = 14.40037893051389227955566592
'            0.40037893051389227955566592   * 16 =  6.40606288822227647289065472
'            0.40606288822227647289065472   * 16 =  6.49700621155642356625047552
'            0.49700621155642356625047552   * 16 =  7.95209938490277706000760832
'            0.95209938490277706000760832   * 16 = 15.23359015844443296012173312
'            0.23359015844443296012173312   * 16 =  3.73744253511092736194772992
'            0.73744253511092736194772992   * 16 = 11.79908056177483779116367872
'            0.79908056177483779116367872   * 16 = 12.78528898839740465861885952
'            0.78528898839740465861885952   * 16 = 12.56462381435847453790175232
'            0.56462381435847453790175232   * 16 =  9.03398102973559260642803712
'            0.03398102973559260642803712   * 16 =  0.54369647576948170284859392
'            0.54369647576948170284859392   * 16 =  8.69914361231170724557750272
'                                                   |
'            Whole numbers are converted to hex ----+
'
'            The hex representation of the fractional parts of the
'            SQUARE ROOT of 2 is:   6a09e667f3bcc908
'
' Parameters:  lngQtyReq       - Number of items to return
'              alngNumbers()   - An array of numbrs to be processed
'              lngRetFmt       - OPTIONAL - Return format.
'                                   e16Chars - All 16 characters (Default)
'                                   eLeft8   - First 8 characters
'                                   eRight8  - Last 8 characters
'              lngSpecificNumber - Optional - This value is greater than zero
'                              if a single value is to be processed.
'                              Default = 0
'              blngDebugPrint  - Optional - Ignored unless lngSpecificNumber
'                              is greater than zero. Used to print detailed
'                              data about number in the immediate window.
'                              To see this, press CTRL+G to open the window
'                              after processing has completed.
'              lngPrimePos     - Optional - Ignored unless lngSpecificNumber
'                              is greater than zero.  Exact position of this
'                              specific prime number when primes are calculated.
'
' Returns:     An array of hex strings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 08-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed to display output details as shown in flowerbox above
'              if user hard codes the exact prime number below.
' 17-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated to allow print detailed structure to VB6 IDE Immediate
'              Window of a specific Prime Number.
' ***************************************************************************
Private Function LoadSquareRoots(ByVal lngQtyReq As Long, _
                                 ByRef alngNumbers() As Long, _
                        Optional ByVal lngRetFmt As enumHEX_RETURN_FORMAT = e16Chars, _
                        Optional ByVal lngSpecificNumber As Long = 0, _
                        Optional ByVal blnDebugPrint As Boolean = False, _
                        Optional ByVal lngPrimePos As Long = 0) As String()

    ' Called by HexRepresentation()

    Dim blnPrintDetail As Boolean   ' Flag for flowerbox output
    Dim lngPtr         As Long      ' Building hex string
    Dim lngLoop        As Long      ' loop counter
    Dim lngIndex       As Long      ' loop counter
    Dim lngNumber      As Long      ' Temp number
    Dim lngWholeNum    As Long      ' Whole number only
    Dim vntBase        As Variant   ' fractional part of number
    Dim vntTemp        As Variant   ' new computed value
    Dim avntHex        As Variant   ' base 16 hex values
    Dim strHex         As String    ' hex character
    Dim astrHex()      As String    ' hex data

    Const ROUTINE_NAME As String = "LoadSquareRoots"

    On Error GoTo LoadSquareRoots_Error

    avntHex = Empty  ' Always start with empty variants and arrays
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()
    blnPrintDetail = False

    ReDim astrHex(lngQtyReq)

    ' Load hex array
    avntHex = Array("0", "1", "2", "3", "4", "5", "6", "7", _
                    "8", "9", "a", "b", "c", "d", "e", "f")

    ' Loop thru array of prime numbers
    ' passed to this routine
    For lngIndex = 0 To lngQtyReq - 1

        lngNumber = CDec(alngNumbers(lngIndex))   ' Convert to decimal base
        vntBase = SquareRoot(lngNumber)           ' Get square root of a number
        strHex = vbNullString

        '---------------------------------------------
        ' If lngSpecificNumber is greater than zero
        ' and blnDebugPrint = True then additional
        ' output will be directed to VB6 IDE Immediate
        ' Window reflecting the detailed data as shown
        ' in above flowerbox.
        '---------------------------------------------
        DoEvents
        If lngSpecificNumber > 0 And blnDebugPrint Then
            ' Set flag to print to immediate (Debug) window
            blnPrintDetail = True
        End If

        ' Print to immediate (Debug) window
        If blnPrintDetail Then
            Debug.Print CStr(lngPrimePos) & "th Prime number = " & CStr(lngSpecificNumber)
            Debug.Print "SQUARE ROOT of " & CStr(lngSpecificNumber) & " = " & CStr(vntBase)
            Debug.Print Space$(20) & "|"
            Debug.Print Space$(20) & "+-- Whole number is dropped"
            Debug.Print ""
        End If

        strHex = Space$(32)   ' Preload with blank spaces
        lngPtr = 1            ' Set output pointer to first position

        For lngLoop = 1 To HEX_LENGTH

            vntBase = CDec(CDec(vntBase) - CDec(Fix(vntBase)))   ' Drop the whole number
            vntTemp = CDec(CDec(vntBase) * CDec(HEX_BASE))       ' Multiply by hex base

            ' Print to immediate (Debug) window
            If blnPrintDetail Then
                Debug.Print Format$(vntBase, "!" & String$(30, "@")) & " * 16 = " & _
                            Format$(vntTemp, "!" & String$(30, "@"))
            End If

            lngWholeNum = CDec(Fix(vntTemp))                 ' Capture whole number
            Mid$(strHex, lngPtr, 1) = avntHex(lngWholeNum)   ' Append hex character to output string
            lngPtr = lngPtr + 1                              ' Increment pointer

            vntTemp = CDec(vntTemp)   ' Convert to decimal base
            vntBase = CDec(vntTemp)   ' Transfer to base value

        Next lngLoop

        strHex = TrimStr(strHex)   ' Remove trailing blanks

        ' Print to immediate (Debug) window
        If blnPrintDetail Then
            Debug.Print Space$(39) & "|"
            Debug.Print "Whole numbers are converted to hex ----+"
            Debug.Print ""
            Debug.Print "The hex representation of the fractional parts of the"
            Debug.Print "SQUARE ROOT of " & CStr(lngSpecificNumber) & " is:   " & strHex
            Debug.Print String$(60, "*")
            Exit For   ' exit For...Next loop
        End If

        Select Case lngRetFmt
               Case e16Chars: strHex = Left$(strHex, HEX_LENGTH)   ' Save all 16 characters
               Case eLeft8:   strHex = Left$(strHex, 8)            ' Save first 8 characters
               Case eRight8:  strHex = Right$(strHex, 8)           ' Save last 8 characters
        End Select

        astrHex(lngIndex) = UCase$(strHex)   ' verify data in uppercase

    Next lngIndex

    LoadSquareRoots = astrHex()

LoadSquareRoots_CleanUp:
    avntHex = Empty   ' Always empty variants when not needed
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()   ' Always empty arrays when not needed

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadSquareRoots_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume LoadSquareRoots_CleanUp

End Function

' ***************************************************************************
' Routine:       LoadCubeRoots
'
' Description:   Got this explanation from Ask Dr. Math web site.
'                http://mathforum.org/dr.math/ 
'                 
'                There is a simple process for converting a base 10 decimal 
'                fraction to a "decimal" fraction in another base.  In this
'                case, hex will be our new base (16).  We repeatedly multiply
'                our decimal fraction by the new base, picking off the whole
'                number part each time as the next digit of the final output. 
'                 
' Example:   1st Prime number = 2
'            CUBE ROOT of 2 = 1.2599210498948731647672106073 
'                             |
'                             +-- Whole number is dropped
' 
'            0.2599210498948731647672106073 * 16 =  4.1587367983179706362753697168 
'            0.1587367983179706362753697168 * 16 =  2.5397887730875301804059154688 
'            0.5397887730875301804059154688 * 16 =  8.636620369400482886494647501 
'            0.636620369400482886494647501  * 16 = 10.185925910407726183914360016 
'            0.185925910407726183914360016  * 16 =  2.974814566523618942629760256 
'            0.974814566523618942629760256  * 16 = 15.597033064377903082076164096 
'            0.597033064377903082076164096  * 16 =  9.552529030046449313218625536 
'            0.552529030046449313218625536  * 16 =  8.840464480743189011498008576 
'            0.840464480743189011498008576  * 16 = 13.447431691891024183968137216 
'            0.447431691891024183968137216  * 16 =  7.158907070256386943490195456 
'            0.158907070256386943490195456  * 16 =  2.542513124102191095843127296 
'            0.542513124102191095843127296  * 16 =  8.680209985635057533490036736 
'            0.680209985635057533490036736  * 16 = 10.883359770160920535840587776 
'            0.883359770160920535840587776  * 16 = 14.133756322574728573449404416 
'            0.133756322574728573449404416  * 16 =  2.140101161195657175190470656 
'            0.140101161195657175190470656  * 16 =  2.241618579130514803047530496 
'                                                   |
'            Whole numbers are converted to hex ----+
' 
'            The hex representation of the fractional parts of the
'            CUBE ROOT of 2 is:   428a2f98d728ae22 
'
' Parameters:  lngQtyReq       - Number of items to return
'              alngNumbers()   - An array of numbrs to be processed
'              lngRetFmt       - OPTIONAL - Return format.
'                                   e16Chars - All 16 characters (Default)
'                                   eLeft8   - First 8 characters
'                                   eRight8  - Last 8 characters
'              lngSpecificNumber - Optional - This value is greater than zero
'                              if a single value is to be processed.
'                              Default = 0
'              blngDebugPrint  - Optional - Ignored unless lngSpecificNumber
'                              is greater than zero. Used to print detailed
'                              data about number in the immediate window.
'                              To see this, press CTRL+G to open the window
'                              after processing has completed.
'              lngPrimePos     - Optional - Ignored unless lngSpecificNumber
'                              is greater than zero.  Exact position of this
'                              specific prime number when primes are calculated.
'
' Returns:     An array of hex strings
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 08-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 08-Jun-2011  Kenneth Ives  kenaso@tx.rr.com
'              Modifed to display output details as shown in flowerbox above
'              if user hard codes the exact prime number below.
' 17-Jan-2017  Kenneth Ives  kenaso@tx.rr.com
'              Updated to allow print detailed structure to VB6 IDE Immediate
'              Window of a specific Prime Number.
' ***************************************************************************
Private Function LoadCubeRoots(ByVal lngQtyReq As Long, _
                               ByRef alngNumbers() As Long, _
                      Optional ByVal lngRetFmt As enumHEX_RETURN_FORMAT = e16Chars, _
                      Optional ByVal lngSpecificNumber As Long = 0, _
                      Optional ByVal blnDebugPrint As Boolean = False, _
                      Optional ByVal lngPrimePos As Long = 0) As String()

    ' Called by HexRepresentation()

    Dim blnPrintDetail As Boolean   ' Flag for flowerbox output
    Dim lngPtr         As Long      ' Building hex string
    Dim lngLoop        As Long      ' loop counter
    Dim lngIndex       As Long      ' loop counter
    Dim lngNumber      As Long      ' Temp number
    Dim lngWholeNum    As Long      ' Whole number only
    Dim vntBase        As Variant   ' fractional part of number
    Dim vntTemp        As Variant   ' new computed value
    Dim avntHex        As Variant   ' base 16 hex values
    Dim strHex         As String    ' hex character
    Dim astrHex()      As String    ' hex data

    Const ROUTINE_NAME As String = "LoadCubeRoots"

    On Error GoTo LoadCubeRoots_Error

    avntHex = Empty  ' Always start with empty variants and arrays
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()
    blnPrintDetail = False

    ReDim astrHex(lngQtyReq)

    ' Load hex array
    avntHex = Array("0", "1", "2", "3", "4", "5", "6", "7", _
                    "8", "9", "a", "b", "c", "d", "e", "f")

    ' Loop thru array of prime numbers
    ' passed to this routine
    For lngIndex = 0 To lngQtyReq - 1

        lngNumber = CDec(alngNumbers(lngIndex))   ' Convert to decimal base
        vntBase = CubeRoot(lngNumber)             ' Get cube root of a number
        strHex = vbNullString

        '---------------------------------------------
        ' If lngSpecificNumber is greater than zero
        ' and blnDebugPrint = True then additional
        ' output will be directed to VB6 IDE Immediate
        ' Window reflecting the detailed data as shown
        ' in above flowerbox.
        '---------------------------------------------
        DoEvents
        If lngSpecificNumber > 0 And blnDebugPrint Then
            ' Set flag to print to immediate (Debug) window
            blnPrintDetail = True
        End If

        ' Print to immediate (Debug) window
        If blnPrintDetail Then
            Debug.Print CStr(lngPrimePos) & "th Prime number = " & CStr(lngSpecificNumber)
            Debug.Print "CUBE ROOT of " & CStr(lngSpecificNumber) & " = " & CStr(vntBase)
            Debug.Print Space$(18) & "|"
            Debug.Print Space$(18) & "+-- Whole number is dropped"
            Debug.Print ""
        End If

        strHex = Space$(32)   ' Preload with blank spaces
        lngPtr = 1            ' Set output pointer to first position

        For lngLoop = 1 To HEX_LENGTH

            vntBase = CDec(CDec(vntBase) - CDec(Fix(vntBase)))   ' Drop the whole number
            vntTemp = CDec(CDec(vntBase) * CDec(HEX_BASE))       ' Multiply by hex base

            ' Print to immediate (Debug) window
            If blnPrintDetail Then
                Debug.Print Format$(vntBase, "!" & String$(30, "@")) & " * 16 = " & _
                            Format$(vntTemp, "!" & String$(30, "@"))
            End If

            lngWholeNum = CDec(Fix(vntTemp))                 ' Capture whole number
            Mid$(strHex, lngPtr, 1) = avntHex(lngWholeNum)   ' Append hex character to output string
            lngPtr = lngPtr + 1                              ' Increment pointer

            vntTemp = CDec(vntTemp)   ' Convert to decimal base
            vntBase = CDec(vntTemp)   ' Transfer to base value

        Next lngLoop

        strHex = TrimStr(strHex)   ' Remove trailing blanks

        ' Print to immediate (Debug) window
        If blnPrintDetail Then
            Debug.Print Space$(39) & "|"
            Debug.Print "Whole numbers are converted to hex ----+"
            Debug.Print ""
            Debug.Print "The hex representation of the fractional parts of the"
            Debug.Print "CUBE ROOT of " & CStr(lngSpecificNumber) & " is:   " & strHex
            Debug.Print String$(60, "-")
            Exit For   ' exit For...Next loop
        End If

        Select Case lngRetFmt
               Case e16Chars: strHex = Left$(strHex, HEX_LENGTH)   ' Save all 16 characters
               Case eLeft8:   strHex = Left$(strHex, 8)            ' Save first 8 characters
               Case eRight8:  strHex = Right$(strHex, 8)           ' Save last 8 characters
        End Select

        astrHex(lngIndex) = UCase$(strHex)   ' verify data in uppercase

    Next lngIndex

    LoadCubeRoots = astrHex()

LoadCubeRoots_CleanUp:
    avntHex = Empty   ' Always empty variants when not needed
    vntBase = Empty
    vntTemp = Empty
    Erase astrHex()   ' Always empty arrays when not needed

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

LoadCubeRoots_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume LoadCubeRoots_CleanUp

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
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function SquareRoot(ByVal vntInput As Variant) As Variant

    ' Called by LoadSquareRoots()

    Dim vntTemp       As Variant   ' Argument - May be positive or negative value
    Dim vntWork1      As Variant   ' Any general approximation to cube root
    Dim vntWork2      As Variant   ' Next successive approximation to cube root
    Dim lngCounter    As Long      ' Cycle loop control counter
    Dim strSqrMinus1  As String    ' Represents the square root of minus 1

    Const ROUTINE_NAME As String = "SquareRoot"

    On Error GoTo SquareRoot_Error

    strSqrMinus1 = vbNullString
    vntWork1 = CDec(0)   ' Initialize (vntWork1) as decimal data type
    vntWork2 = CDec(0)   ' Initialize (vntWork2) as decimal data type
    vntTemp = vntInput   ' Check for invalid numeric argument

    If Not IsNumeric(vntTemp) Then
        GoTo SquareRoot_Error
    End If

    vntTemp = CDec(vntTemp)   ' Convert argument into decimal data type

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
    lngCounter = 50           ' Set limit of iterations to 50 max

    ' A loop to grind out the square root using a series of
    ' successive approximations, starting with (vntWork1).
    Do
        ' Compute next approx (vntWork2) from (vntWork1)
        vntWork2 = (vntWork1 + vntTemp / vntWork1) / 2

        ' Check if finished
        If (vntWork2 = vntWork1) Or (lngCounter <= 0) Then
            Exit Do   ' exit Do..Loop
        End If

        vntWork1 = vntWork2           ' Update approx to current value
        lngCounter = lngCounter - 1   ' Update limit counter
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
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=21048&lngWId=1
' 17-Nov-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function CubeRoot(ByVal vntInput As Variant) As Variant

    ' Called by LoadCubeRoots()

    Dim vntTemp    As Variant   ' Argument - May be positive or negative value
    Dim vntWork1   As Variant   ' Any general approximation to cube root
    Dim vntWork2   As Variant   ' Next successive approximation to cube root
    Dim lngCounter As Long      ' Cycle loop control counter
    Dim strSign    As String    ' Sign of argument - Attached to result

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
        vntWork1 = vntTemp ^ (1 / 3)   ' Use VB cube root as 1st approximation
        lngCounter = 50                ' Set limit of iterations to 50 max

        ' A loop to grind out the cube root using a series of
        ' successive approximations, starting with (vntWork1).
        Do
            ' Compute next approx (vntWork2) from (vntWork1)
            vntWork2 = ((2 * vntWork1) + vntTemp / (vntWork1 * vntWork1)) / 3

            ' Check if finished
            If (vntWork2 = vntWork1) Or (lngCounter <= 0) Then
                Exit Do   ' exit Do..Loop
            End If

            vntWork1 = vntWork2           ' Update approx to current value
            lngCounter = lngCounter - 1   ' Update limit counter
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
    ' Enum enumHEX_REPRESENTATION
    Const eCubeRoots   As Long = 0&
    Const eSquareRoots As Long = 1&

    ' Enum enumHEX_RETURN_FORMAT
    Const e16Chars     As Long = 0&
    Const eLeft8       As Long = 1&
    Const eRight8      As Long = 2&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    StopProcessing = False

End Sub
