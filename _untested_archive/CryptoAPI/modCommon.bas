Attribute VB_Name = "modCommon"
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME As String = "modCommon"


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

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
' Routine:       ByteArrayToHex
'
' Description:   Convert a byte array into a hex string of data.
'
' Parameters:    abytData() - Array of string data in byte format
'
' Returns:       hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function ByteArrayToHex(ByRef abytData() As Byte) As String

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strHex     As String

    Const ROUTINE_NAME As String = "ByteArrayToHex"
    
    If IsArrayInitialized(abytData()) Then

        lngLength = UBound(abytData)        ' capture length of incoming data
        strHex = Space$(lngLength * 2 + 2)  ' Preload output string with blanks
        lngPointer = 1                      ' index pointer for output string

        ' Convert byte array to hex string
        For lngIndex = 0 To lngLength
            Mid$(strHex, lngPointer, 2) = Right$("0" & Hex$(abytData(lngIndex)), 2)
            lngPointer = lngPointer + 2
        Next lngIndex

        strHex = TrimStr(strHex)   ' Remove unwanted characters

        ' Verify that this is hex data
        If Not IsHexData(strHex) Then
            InfoMsg "Failed to convert data to hex." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
            strHex = vbNullString
            gblnStopProcessing = True
        End If

    Else
        InfoMsg "Incoming data not available for conversion." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        strHex = vbNullString
        gblnStopProcessing = True
    End If

    ByteArrayToHex = strHex  ' Return hex string
    strHex = vbNullString

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
' ***************************************************************************
Public Function HexToByteArray(ByVal strHex As String) As Byte()

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim abytData() As Byte

    Const ROUTINE_NAME As String = "HexToByteArray"
    
    Erase abytData()           ' Always start with empty arrays
    strHex = TrimStr(strHex)   ' Remove leading\trailing blanks

    If Len(strHex) = 0 Then
        InfoMsg "Invalid hex string length for conversion (1)." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        Exit Function
    End If

    lngLength = Len(strHex)   ' Capture length of hex string

    If lngLength = 0 Then
        InfoMsg "No data to process (2)." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        Exit Function
    End If

    ' Verify this is hex data
    If Not IsHexData(strHex) Then
        InfoMsg "Invalid hex string for conversion." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME, , , 3
        Exit Function
    End If

    lngLength = Len(strHex)   ' Capture length of hex string

    If lngLength Mod 2 <> 0 Then
        strHex = "0" & strHex   ' Adjust to be divisible by 2
    End If

    lngLength = Len(strHex)     ' Capture length of hex string

    ' String must be divisable by 2
    If lngLength Mod 2 = 0 Then

        ReDim abytData(lngLength \ 2)  ' resize output array

        ' start converting data string two
        ' characters at a time into an ASCII
        ' decimal value
        For lngIndex = 0 To UBound(abytData) - 1
            abytData(lngIndex) = CByte("&H" & Mid$(strHex, lngIndex * 2 + 1, 2))
        Next lngIndex

        ReDim Preserve abytData(lngIndex - 1)  ' resize to actual size

    Else
        ReDim abytData(1)
    End If

    HexToByteArray = abytData()

    Erase abytData()  ' Always empty arrays when not needed

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

