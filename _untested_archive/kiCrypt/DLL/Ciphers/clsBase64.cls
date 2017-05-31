VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:       cBase64  (clsBase64.cls)
'
' Description:  What is BASE64?  A method of encoding binary data within
'               text. If you'll remember, binary data is a full 8-bits per
'               byte, whereas text uses a little more than 6 bits per byte.
'               A 6-bit number has 64 combinations, hence the term "BASE64".
'
'               The way it works is that every three 8-bit bytes are stored
'               in four 6-bit characters, where the characters are in the
'               range [A-Z][a-z][0-9][+/]. (Count 'em up; that's 64 total
'               characters).  Since this doesn't exactly line up, pad
'               characters of [=] are used at the very end.
'
'               Where Base64 Is used?  HTTP "Basic" Authentication and PGP
'               signatures And keys.  One of many ways of encoding e-mail.
'
'               5.2  Base64 Content-Transfer-Encoding
'
'               The  Base64   Content-Transfer-Encoding   is   designed   to
'               represent  arbitrary  sequences  of octets in a form that is
'               not humanly readable.  The encoding and decoding  algorithms
'               are simple, but the encoded data are consistently only about
'               33 percent larger than the unencoded data.  This encoding is
'               based on the one used in Privacy Enhanced Mail applications,
'               as defined in RFC 1113.   The  base64  encoding  is  adapted
'               from  RFC  1113, with one change:  base64 eliminates the "*"
'               mechanism for embedded clear text.
'
'               A 65-character subset of US-ASCII is used, enabling  6  bits
'               to  be  represented per printable character. (The extra 65th
'               character, "=", is used  to  signify  a  special  processing
'               function.)
'
'               NOTE:  This subset has the important  property  that  it  is
'               represented   identically   in  all  versions  of  ISO  646,
'               including US ASCII, and all characters  in  the  subset  are
'               also  represented  identically  in  all  versions of EBCDIC.
'               Other popular encodings, such as the encoding  used  by  the
'               UUENCODE  utility  and the base85 encoding specified as part
'               of Level 2 PostScript, do not share  these  properties,  and
'               thus  do  not  fulfill the portability requirements a binary
'               transport encoding for mail must meet.
'
'               The encoding process represents 24-bit groups of input  bits
'               as  output  strings of 4 encoded characters. Proceeding from
'               left  to  right,  a  24-bit  input  group   is   formed   by
'               concatenating  3  8-bit input groups. These 24 bits are then
'               treated as 4 concatenated 6-bit groups,  each  of  which  is
'               translated  into a single digit in the base64 alphabet. When
'               encoding a bit stream  via  the  base64  encoding,  the  bit
'               stream  must  be  presumed  to  be  ordered  with  the most-
'               significant-bit first.  That is, the first bit in the stream
'               will be the high-order bit in the first byte, and the eighth
'               bit will be the low-order bit in the first byte, and so on.
'
'               Each 6-bit group is used as an index into  an  array  of  64
'               printable  characters. The character referenced by the index
'               is placed in the output string. These characters, identified
'               in  Table  1,  below,  are  selected so as to be universally
'               representable,  and  the  set   excludes   characters   with
'               particular  significance to SMTP (e.g., ".", "CR", "LF") and
'               to the encapsulation boundaries  defined  in  this  document
'               (e.g., "-").
'
'                               Table 1: The Base64 Alphabet
'
'                     Value           Value           Value           Value
'                   Encoding        Encoding        Encoding        Encoding
'                      0 A            17 R            34 i            51 z
'                      1 B            18 S            35 j            52 0
'                      2 C            19 T            36 k            53 1
'                      3 D            20 U            37 l            54 2
'                      4 E            21 V            38 m            55 3
'                      5 F            22 W            39 n            56 4
'                      6 G            23 X            40 o            57 5
'                      7 H            24 Y            41 p            58 6
'                      8 I            25 Z            42 q            59 7
'                      9 J            26 a            43 r            60 8
'                     10 K            27 b            44 s            61 9
'                     11 L            28 c            45 t            62 +
'                     12 M            29 d            46 u            63 /
'                     13 N            30 e            47 v
'                     14 O            31 f            48 w         (pad) =
'                     15 P            32 g            49 x
'                     16 Q            33 h            50 y
'
'               The output stream (encoded bytes)  must  be  represented  in
'               lines  of  no more than 76 characters each.  All line breaks
'               or other characters not found in Table 1 must be ignored  by
'               decoding  software.   In  base64 data, characters other than
'               those in  Table  1,  line  breaks,  and  other  white  space
'               probably  indicate  a  transmission  error,  about  which  a
'               warning  message  or  even  a  message  rejection  might  be
'               appropriate under some circumstances.''
'
'               Special processing is performed if fewer than  24  bits  are
'               available  at  the  end  of  the data being encoded.  A full
'               encoding quantum is always completed at the end of  a  body.
'               When  fewer  than  24  input  bits are available in an input
'               group, zero bits  are  added  (on  the  right)  to  form  an
'               integral number of 6-bit groups.  Output character positions
'               which are not required to represent actual  input  data  are
'               set  to  the  character  "=".   Since all base64 input is an
'               integral number of octets,  only  the  following  cases  can
'               arise:  (1)  the  final  quantum  of  encoding  input  is an
'               integral multiple of  24  bits;  here,  the  final  unit  of
'               encoded  output will be an integral multiple of 4 characters
'               with no "=" padding, (2) the final quantum of encoding input
'               is  exactly  8  bits; here, the final unit of encoded output
'               will  be  two  characters  followed  by  two   "="   padding
'               characters,  or  (3)  the final quantum of encoding input is
'               exactly 16 bits; here, the final unit of encoded output will
'               be three characters followed by one "=" padding character.
'
'               Care must be taken to use the proper octets for line  breaks
'               if base64 encoding is applied directly to text material that
'               has not been converted to  canonical  form.  In  particular,
'               text  line  breaks  should  be converted into CRLF sequences
'               prior to base64 encoding. The important  thing  to  note  is
'               that this may be done directly by the encoder rather than in
'               a prior canonicalization step in some implementations.
'
' NOTE:         There is no need to worry about quoting apparent encapsulation
'               boundaries within base64-encoded parts of multipart entities
'               because no hyphen characters are used in the base64 encoding.
'
'               Base64 was designed for handling messages and files attached
'               to messages.  Most mail systems have a limitation of five
'               megabytes for attachments being sent and received. I normally
'               set the limitation of five megabytes prior to calling this
'               module.  If you call this module directly, then set your
'               restrictions within this module otherwise you will have to FTP
'               the encrypted file because the mail systems will not allow it
'               to overload their primary processes.
'
' References:   RFC 1341 (RFC1341)
'               http://www.faqs.org/rfcs/rfc1341.html
'
'               Comintern
'               http://www.vbforums.com/showthread.php?t=379072
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2006  Comintern
'              http://www.vbforums.com/showthread.php?t=379072
' 28-JUN-2006  Kenneth Ives  kenaso@tx.rr.com
'              Documented and modified
' 21-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Arrays are not sized until routines EncryptData() or
'              DecryptData() are called. Saves resources.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
'              Removed obsolete code.  Updated documentation.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME     As String = "clsBase64"
  Private Const ONE_MASK        As Long = 16515072   ' 000000 111111 111111 111111
  Private Const TWO_MASK        As Long = 258048     ' 111111 000000 111111 111111
  Private Const THREE_MASK      As Long = 4032       ' 111111 111111 000000 111111
  Private Const FOUR_MASK       As Long = 63         ' 111111 111111 111111 000000
  Private Const HIGH_MASK       As Long = 16711680   ' 11111111 00000000 00000000
  Private Const MID_MASK        As Long = 65280      ' 00000000 11111111 00000000
  Private Const LOW_MASK        As Long = 255        ' 00000000 00000000 11111111
  Private Const TWO_TO_THE_6TH  As Long = 64         ' 2 to the 6th power
  Private Const TWO_TO_THE_8TH  As Long = 256        ' 2 to the 8th power
  Private Const TWO_TO_THE_12TH As Long = 4096       ' 2 to the 12th power
  Private Const TWO_TO_THE_16TH As Long = 65536      ' 2 to the 16th power
  Private Const TWO_TO_THE_18TH As Long = 262144     ' 2 to the 18th power

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Currency)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m cur MaxSize
' Variable name:     mcurMaxSize
'
' ***************************************************************************
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************
    
' Update progress bar
Public Event CipherProgress(ByVal lngProgress As Long)


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let CreateNewFile(ByVal blnNewFile As Boolean)
   
    ' 21-Feb-2012 Added Property to desiginate if input file
    '             is to be overwritten after encryption/decryption.
    '
    ' TRUE - Create new file to hold encrypted/decrypted data
    ' FALSE - Overwrite input file after encryption/decryption
    
    mblnCreateNewFile = blnNewFile

End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Base64_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objBase64.Base64_File abytFile(), blnEncrypt                 ' Encrypt/Decrypt data
' 
' Parameters:    abytFile() - Path/filename in byte array to be hashed
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files less than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' ***************************************************************************
Public Function Base64_File(ByRef abytFile() As Byte, _
                   Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String  ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim strExtension As String
    Dim hFile        As Long
    Dim curPointer   As Currency
    Dim abytData()   As Byte

    Const ROUTINE_NAME As String = "Base64_File"
    
    On Error GoTo Base64_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal
    
    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        
        strExtension = vbNullString
        strExtension = GetFilenameExt(strFileName)

        Select Case blnEncrypt
               
               Case True
                    ' First verify file name has an extension
                    If Len(strExtension) > 0 Then
                        strTarget = Mid$(strFileName, 1, Len(strFileName) - (Len(strExtension) + 1)) & ENCRYPT_EXT
                    Else
                        ' No file extension
                        strTarget = strFileName & ENCRYPT_EXT
                    End If
               
               Case False
                    ' First verify file name has an extension
                    If Len(strExtension) > 0 Then
                        strTarget = Mid$(strFileName, 1, Len(strFileName) - (Len(strExtension) + 1)) & DECRYPT_EXT
                    Else
                        ' No file extension
                        strTarget = strFileName & DECRYPT_EXT
                    End If
        End Select
    End If
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo Base64_File_CleanUp
    End If

    Erase abytData()                ' Always start with empty arrays
    curPointer = 1                  ' set to first position in source file
    strTempFile = CreateTempFile()  ' Create temp target file
    
    ' Make sure the target file is empty
    hFile = FreeFile
    Open strTempFile For Output As #hFile
    Close #hFile
    
    ' Open source file and capture its contents
    hFile = FreeFile
    Open strSource For Binary Access Read As #hFile
    ReDim abytData(LOF(hFile) - 1)
    Get #hFile, , abytData()
    Close #hFile
    
    mcurMaxSize = CCur(UBound(abytData))   ' Save file size for progress calcs
    mcurAccrued = 0@                       ' Init accrued progress
    
    If blnEncrypt Then

        ' Encrypt data
        If Not EncryptData(abytData()) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If
    
    Else
        ' Decrypt data
        If Not DecryptData(abytData()) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If
    
    End If
            
    ' Open the target file to be updated
    hFile = FreeFile
    Open strTempFile For Binary Access Write As #hFile
    
    ' Excessive DoEvents is to keep the write
    ' process in sync with the CPU speed
    DoEvents
    Put #hFile, curPointer, abytData()
    DoEvents
    
    Close          ' Close all files opened by this application
    
    ' Move temp file to destination location
    DoEvents
    On Error Resume Next
    DoEvents

    ' See if a new file is to be created after
    ' performing encryption or decryption
    If mblnCreateNewFile Then
        ' Create new file
        MoveFileEx strTempFile, strTarget, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    Else
        ' Overwrite source file
        MoveFileEx strTempFile, strSource, _
                   MOVEFILE_COPY_ALLOWED Or _
                   MOVEFILE_REPLACE_EXISTING
    End If

    RaiseEvent CipherProgress(100)   ' Update progress bar

Base64_File_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Base64_File = False
    Else
        Base64_File = True
    End If

    Close             ' close all files opened by this application
    Erase abytData()  ' Always empty arrays when not needed
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Base64_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Base64_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Base64_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)      ' convert string data to byte array
'                objBase64.Base64_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
' 
' Parameters:    abytData() - data string in byte array
'                blnEncrypt - [Optional] - TRUE - Encrypt incoming data
'                             FALSE - Decrypt incoming data
'                             Default=TRUE
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-NOV-2004  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 19-APR-2006  Kenneth Ives  kenaso@tx.rr.com
'              Changed input parameter to byte array for security purposes
' ***************************************************************************
Public Function Base64_String(ByRef abytData() As Byte, _
                     Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Const ROUTINE_NAME As String = "Base64_String"

    On Error GoTo Base64_String_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True    ' Set flag to stop processing
        Exit Function
    End If

    mcurMaxSize = CCur(UBound(abytData))   ' Save file size for progress calcs
    mcurAccrued = 0@                       ' Init accrued progress
    
    If blnEncrypt Then
        
        ' Encrypt data
        If Not EncryptData(abytData()) Then
            gblnStopProcessing = True
            Base64_String = False
            GoTo Base64_String_CleanUp
        End If
    
    Else
        ' Decrypt data
        If Not DecryptData(abytData()) Then
            gblnStopProcessing = True
            Base64_String = False
            GoTo Base64_String_CleanUp
        End If
    
    End If
    
Base64_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Base64_String = False
    Else
        Base64_String = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Base64_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Base64_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

' ***************************************************************************
' Routine:      EncryptData
'
' Description:  Encodes data with MIME Base64. Base64 encoding is designed
'               to make abytInputary data survive transport through transport
'               layers that are not 8-bit clean, such as mail bodies.
'
'               Base64 processes data as 24-bit groups, mapping this data
'               to four encoded characters. Base64 encoding is sometimes
'               referred to as 3-to-4 encoding. Each 6 bits of the 24-bit
'               group is used as an index into a mapping table (the base64
'               alphabet) to obtain a character for the encoded data.
'               Padding at the end of the data is performed using the
'               additional "=" character. The encoded data are consistently
'               only about 33 percent larger than the unencoded data.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2006  Comintern
'              http://www.vbforums.com/showthread.php?t=379072
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified, documented, renamed variables
' 21-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Arrays are not sized until routine is called. Saves resources.
' ***************************************************************************
Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    Dim strData        As String
    Dim intPadding     As Integer
    Dim lngChar        As Long
    Dim lngTrip        As Long
    Dim lngTemp        As Long
    Dim lngIndex       As Long
    Dim lngLength      As Long
    Dim lngPosition    As Long
    Dim lngProgress    As Long
    Dim lngMsgLength   As Long
    Dim lngOutputSize  As Long
    Dim alngPowers8()  As Long
    Dim alngPowers16() As Long
    Dim abytInput()    As Byte
    Dim abytTrans()    As Byte
    Dim abytOutput()   As Byte
    
    Const ROUTINE_NAME As String = "EncryptData"
    Const BLOCK_SIZE   As Long = 3
    
    On Error GoTo EncryptData_Error

    ReDim abytTrans(64)   ' Size work arrays
    ReDim alngPowers8(256)
    ReDim alngPowers16(256)
    
    ' Fill the translation table.
    For lngChar = 0 To 63
        Select Case lngChar
               Case 0 To 25:  abytTrans(lngChar) = 65 + lngChar   ' A-Z
               Case 26 To 51: abytTrans(lngChar) = 71 + lngChar   ' a-z
               Case 52 To 61: abytTrans(lngChar) = lngChar - 4    ' 0-9
               Case 62:       abytTrans(lngChar) = 43             ' Chr$(43) = "+"
               Case 63:       abytTrans(lngChar) = 47             ' Chr$(47) = "/"
        End Select
    Next lngChar

    ' Fill the 2^8 and 2^16 lookup tables.
    For lngIndex = 0 To 255
        
        alngPowers8(lngIndex) = lngIndex * TWO_TO_THE_8TH
        alngPowers16(lngIndex) = lngIndex * TWO_TO_THE_16TH
        
    Next lngIndex

    strData = ByteArrayToString(abytData())   ' convert incoming data to string data
    
    If Len(strData) <= 0 Then
        InfoMsg "Input data is missing." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo EncryptData_CleanUp
    End If
    
    intPadding = Len(strData) Mod 3   ' See if the length is divisible by 3
    
    ' If not, figure out the end
    ' pad and resize the input.
    If intPadding Then
        intPadding = 3 - intPadding
        strData = strData & String$(intPadding, Chr$(0))
    End If

    abytInput() = StringToByteArray(strData)           ' Load the input string.
    lngLength = ((UBound(abytInput) + 1) \ 3) * 4      ' Length of resulting string.
    lngMsgLength = lngLength                           ' length of message
    lngTemp = lngLength \ 72                           ' Added space for vbNewLines.
    lngOutputSize = ((lngTemp * 2) + lngLength) - 1    ' Calculate the size of the output buffer.
    ReDim abytOutput(lngOutputSize)                    ' Make the output buffer.
    lngLength = 0                                      ' Reusing this one, so reset it.
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo EncryptData_CleanUp
    End If
    
    For lngIndex = LBound(abytInput) To UBound(abytInput) Step BLOCK_SIZE
    
        ' Compute the 3 bytes
        lngTrip = alngPowers16(abytInput(lngIndex)) + _
                  alngPowers8(abytInput(lngIndex + 1)) + _
                  abytInput(lngIndex + 2)
                  
        lngTemp = lngTrip And ONE_MASK   ' Mask for the first 6 bits
        
        ' Shift it down to the low 6 bits and get the value
        abytOutput(lngPosition) = abytTrans(lngTemp \ TWO_TO_THE_18TH)
        lngTemp = lngTrip And TWO_MASK                                      ' Mask for the second set.
        abytOutput(lngPosition + 1) = abytTrans(lngTemp \ TWO_TO_THE_12TH)  ' Shift it down and translate.
        lngTemp = lngTrip And THREE_MASK                                    ' Mask for the third set.
        abytOutput(lngPosition + 2) = abytTrans(lngTemp \ TWO_TO_THE_6TH)   ' Shift it down and translate.
        abytOutput(lngPosition + 3) = abytTrans(lngTrip And FOUR_MASK)      ' Mask for the low set.
        
        ' Ready for a newline
        If lngLength = 68 Then
            abytOutput(lngPosition + 4) = 13   ' Chr$(13) = vbCr
            abytOutput(lngPosition + 5) = 10   ' Chr$(10) = vbLf
            lngLength = 0                      ' Reset the counter
            lngPosition = lngPosition + 6
        Else
            lngLength = lngLength + 4
            lngPosition = lngPosition + 4
        End If
        
        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + BLOCK_SIZE
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For
        End If
    
    Next lngIndex
    
    ' Shift the padding chars down if it ends with CrLf.
    If abytOutput(lngOutputSize) = 10 Then
        lngOutputSize = lngOutputSize - 2
    End If
    
    ' Add the padding chars if any.  Chr$(61) = "="
    If intPadding = 1 Then
        abytOutput(lngOutputSize) = 61
        
    ElseIf intPadding = 2 Then
        abytOutput(lngOutputSize) = 61
        abytOutput(lngOutputSize - 1) = 61
    End If
    
    ReDim abytData(UBound(abytOutput))
    CopyMemory abytData(0), abytOutput(0), UBound(abytOutput) + 1
    
EncryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False
    Else
        EncryptData = True
    End If

    On Error GoTo 0      ' Nullify this error trap
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

' ***************************************************************************
' Routine:      DecryptData
'
' Description:  Decodes data encoded with MIME Base64. This is the reverse
'               of Base64 encoding.
'
'               Base64 processes data as 24-bit groups, mapping this data
'               to four encoded characters. Base64 encoding is sometimes
'               referred to as 3-to-4 encoding. Each 6 bits of the 24-bit
'               group is used as an index into a mapping table (the base64
'               alphabet) to obtain a character for the encoded data.
'               Padding at the end of the data is performed using the
'               additional "=" character. The encoded data are consistently
'               only about 33 percent larger than the unencoded data.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2006  Comintern
'              http://www.vbforums.com/showthread.php?t=379072
' 19-DEC-2006  Kenneth Ives  kenaso@tx.rr.com
'              Modified, documented, renamed variables
' 21-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Arrays are not sized until routine is called. Saves resources.
' ***************************************************************************
Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    Dim strData        As String
    Dim strTemp        As String
    Dim intPadding     As Integer
    Dim lngChar        As Long
    Dim lngQuad        As Long
    Dim lngTemp        As Long
    Dim lngIndex       As Long
    Dim lngProgress    As Long
    Dim lngPosition    As Long
    Dim lngMsgLength   As Long
    Dim alngPowers6()  As Long
    Dim alngPowers12() As Long
    Dim alngPowers18() As Long
    Dim abytOutput()   As Byte
    Dim abytInput()    As Byte
    Dim abytTrans()    As Byte

    Const ROUTINE_NAME As String = "DecryptData"
    Const BLOCK_SIZE   As Long = 4

    On Error GoTo DecryptData_Error

    ReDim abytTrans(256)   ' Size work arrays
    ReDim alngPowers6(64)
    ReDim alngPowers12(64)
    ReDim alngPowers18(64)
    
    strData = ByteArrayToString(abytData())   ' convert incoming data to string data
    
    If Len(strData) <= 0 Then
        InfoMsg "Input data is missing." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    ' Get rid of the vbNewLines.
    strData = Replace(strData, vbCr, "")
    strData = Replace(strData, vbLf, vbNullString)

    lngMsgLength = Len(strData) Mod 4    ' Test for valid input.
    
    If lngMsgLength > 0 Then
        
        InfoMsg "Input string is not valid Base64." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        gblnStopProcessing = True
        GoTo DecryptData_CleanUp
        
    End If
    
    lngMsgLength = Len(strData)  ' capture message length
    lngPosition = 0
    
    ' InStrRev is faster when you know it's at the end.
    ' Note:  These translate to 0, so you can leave them
    '        in the string and just resize the output.
    If InStrRev(strData, "==") Then
        intPadding = 2
    ElseIf InStrRev(strData, "=") Then
        intPadding = 1
    End If
    
    ' Fill the translation table.
    For lngChar = 0 To 255
        Select Case lngChar
               Case 65 To 90:  abytTrans(lngChar) = lngChar - 65   ' A-Z
               Case 97 To 122: abytTrans(lngChar) = lngChar - 71   ' a-z
               Case 48 To 57:  abytTrans(lngChar) = lngChar + 4    ' 0-9
               Case 43:        abytTrans(lngChar) = 62             ' Chr$(43) = "+"
               Case 47:        abytTrans(lngChar) = 63             ' Chr$(47) = "/"
        End Select
    Next lngChar

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo DecryptData_CleanUp
    End If

    ' Fill the 2^6, 2^12, and 2^18 lookup tables.
    For lngIndex = 0 To 63
        
        alngPowers6(lngIndex) = lngIndex * TWO_TO_THE_6TH
        alngPowers12(lngIndex) = lngIndex * TWO_TO_THE_12TH
        alngPowers18(lngIndex) = lngIndex * TWO_TO_THE_18TH
    
    Next lngIndex

    abytInput() = StringToByteArray(strData)                    ' Load the input byte array.
    ReDim abytOutput((((UBound(abytInput) + 1) \ 4) * 3) - 1)   ' Prepare the output buffer.
    
    For lngIndex = 0 To UBound(abytInput) Step BLOCK_SIZE
        
        ' Rebuild the bits.
        lngQuad = alngPowers18(abytTrans(abytInput(lngIndex))) + _
                  alngPowers12(abytTrans(abytInput(lngIndex + 1))) + _
                  alngPowers6(abytTrans(abytInput(lngIndex + 2))) + _
                  abytTrans(abytInput(lngIndex + 3))
                  
        lngTemp = lngQuad And HIGH_MASK                          ' Mask for the first byte
        abytOutput(lngPosition) = lngTemp \ TWO_TO_THE_16TH      ' Shift it down
        lngTemp = lngQuad And MID_MASK                           ' Mask for the second byte
        abytOutput(lngPosition + 1) = lngTemp \ TWO_TO_THE_8TH   ' Shift it down
        abytOutput(lngPosition + 2) = lngQuad And LOW_MASK       ' Mask for the third byte
        lngPosition = lngPosition + 3
        
        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + BLOCK_SIZE
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For
        End If

    Next lngIndex

    strTemp = ByteArrayToString(abytOutput())    ' Convert work array to string
    
    If intPadding Then
        strTemp = Left$(strTemp, Len(strTemp) - intPadding)   ' Chop off any extra bytes.
    End If
    
    abytData() = StringToByteArray(strTemp)   ' Convert string to byte array

DecryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False
    Else
        DecryptData = True
    End If

    On Error GoTo 0       ' Nullify this error trap
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

