VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTwofish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cTwofish  (clsTwofish.cls)
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
' *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING *** WARNING ***
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
'    You acknowledge that this software is subject to the export control
'    laws and regulations of the United States ("U.S.") and agree to abide
'    by those laws and regulations. Under U.S. law, this software may not
'    be downloaded or otherwise exported, reexported, or transferred to
'    restricted countries, restricted end-users, or for restricted
'    end-uses. The U.S. currently has embargo restrictions against Cuba,
'    Iran, Iraq, Libya, North Korea, Sudan, and Syria. The lists of
'    restricted end-users are maintained on the U.S. Commerce Department's
'    Denied Persons List, the Commerce Department's Entity List, the
'    Commerce Department's List of Unverified Persons, and the U.S.
'    Treasury Department's List of Specially Designated Nationals and
'    Blocked Persons. In addition, this software may not be downloaded or
'    otherwise exported, reexported, or transferred to an end-user engaged
'    in activities related to weapons of mass destruction.
'
' *=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'
' Description:   Twofish is a block cipher by Counterpane Labs. It was one
'                of the five AES (Advanced Encryption Standard) finalists.
'                The Twofish cipher has not been patented and the reference
'                implementation has been placed in the public domain. As a
'                result, the Twofish algorithm is free for anyone to use
'                without any restrictions whatsoever.
'
'                Twofish is a 128-bit block cipher that accepts a variable
'                length key up to 256 bits.  The Twofish block cipher is
'                Counterpane Labs' candidate for the new AES. It was one of
'                the five finalists chosen by NIST (National Institute of
'                Standards and Technology ) from a field of 15 candidates as
'                explained above. Twofish is designed to be highly secure
'                and highly flexible. It is well suited for large
'                microprocessors, 8-bit smart card microprocessors, and
'                dedicated hardware. Counterpane Labs has spent over one
'                thousand hours cryptanalyzing Twofish, and has found no
'                attacks that can break the full 16 round version of the
'                algorithm. The algorithm is very secure when the full 16
'                rounds are used. It was also the fastest AES candidate, and
'                one of the most compact. Its conservative design allows the
'                ability to trade off key setup time for encryption speed, as
'                well as sacrificing smaller memory requirements to obtain
'                greater encryption speed. Algorithm implementers are given a
'                lot of flexibility to play around with to produce a version
'                of Twofish that is just right for the job at hand.
'
'                Twofish was designed by Bruce Schneier, John Kelsey,
'                Doug Whiting, David Wagner, Chris Hall, and Niels Ferguson.
'                Bruce Schneier also designed the Blowfish algorithm, which
'                remains unbroken after many years of cryptanalysis and has
'                been implemented in over 130 commercial applications.
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length is expanded to meet and exceed AES
'                  standards.  Key lengths are 32 to 448 in 32 bit
'                  increments.  See property KeyLength().
'                **********************************************************
'
' Warning:       The weakest link in any encryption is the password. This
'                includes the security, length and components that make up
'                the password.  Password length and components are usually
'                dictated by organizational standards.  Security is the
'                responsibility of the user. Do not write down or discuss
'                your password with others.  Lock your screensaver when not
'                at your workstation.
'
' Suggestions:   To strengthen your encryption, prefix each block of data with
'                4-16 bytes of salt (random) data.  After decrypting each
'                block do not forget to drop these random bytes of data.
'
'                                     or
'
'                Compress the file or data string first using your favorite
'                archive utility (Zip, Rar, etc.).  Now encrypt the data.
'
'                                     or
'
'                First encode data with Base64 and then perform your encryption.
'                After decryption, decode your data with Base64.  Most messaging
'                is done with string data and not file data.
'
'                The above are just suggestions.  Happy coding.
'
' Note:          After encryption, data sizes will not match original sizes.
'                This is due to internal padding and the storing of information
'                required to decrypt the data later.
'
' References:    Twofish was created and analyzed by:
'                Bruce Schneier, John Kelsey, Doug Whiting, David Wagner,
'                Chris Hall, Niels Ferguson
'                http://www.schneier.com/Twofish.html
'
'                Tropical software (Portions of description above)
'                http://www.tropsoft.com/strongenc/Twofish.htm
'
'                Visual Basic translation by Fredrik Qvarfort
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=12023&lngWId=1
'                12-Oct-2000
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-JUN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Documented, modified, updated variable names
' 05-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Consolidated routines that performed bit shifting or bit
'              rotating to call w32Shift or w32Rotate respectfully.  Removed
'              obsolete bit shifting and rotating routines.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified Twofish_File() and Twofish_String() routines to
'              encrypt appended data information for additional security.
'              Moved ExpandData() routine to basCommon.bas module.
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet and exceed AES standards.
'              (128, 192, 256, 384, 512, 1024 bits)  Modified EvaluateKey()
'              routine to use utilize the new key lengths.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed OffSet range in EncryptData() and DecryptData() routines.
'              Moved GetBlockSize() routine to common module.
'              Set key length to increments of 32 bits for 128 to 448.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in Twofish_String().
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within Twofish_File() routine.
'                4. Split size testing in Twofish_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Moved progress bar code from BlockEncrypt() and
'                BlockDecrypt() routines to EncryptData() and DecryptData()
'                routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Twofish_File() routine to reference new property.
' 05-Dec-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified ProcessKey() routine to perform multiple rounds of
'              mixing the values for the primary key [malngKey()] and
'              S-box [malngSBox()] arrays.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsTwofish"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MB_5           As Long = &H500000           ' 5242880 bytes
  Private Const CIPHER_ROUNDS  As Long = 16
  Private Const BLOCK_SIZE     As Long = 16
  Private Const ROUND_SUBKEYS  As Long = 4
  Private Const MAX_ROUNDS     As Long = 10
  Private Const MAX_BYTE       As Long = 256

' ***************************************************************************
' Module Variables
'                    +---------------- Module level designator
'                    | +-------------- Array designator
'                    | |  +----------- Data type (Byte)
'                    | |  |     |----- Variable subname
'                    - - --- ---------
' Naming standard:   m a byt Password
' Variable name:     mabytPassword
' ***************************************************************************
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency
  Private mcurOrigLength    As Currency
  Private mcurCurrByteCnt   As Currency
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private malngKey()        As Long
  Private malngMDS()        As Long
  Private malngSBox()       As Long
  Private mabytPBox()       As Byte
  Private mabytPassword()   As Byte
  Private mobjMath32        As cMath32


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event CipherProgress(ByVal lngProgress As Long)

' Track amount of data being processed.
Public Event CurrentBlockSize(ByVal lngBlockSize As Long)


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

Public Property Let KeyLength(ByVal lngKeyLength As Long)

    ' 01-Jan-2009 Set increments of 32 bits for 32 to 448
    Select Case lngKeyLength
           Case 448:  mlngKeyLength = 448
           Case 416:  mlngKeyLength = 416
           Case 384:  mlngKeyLength = 384
           Case 352:  mlngKeyLength = 352
           Case 320:  mlngKeyLength = 320
           Case 288:  mlngKeyLength = 288
           Case 256:  mlngKeyLength = 256
           Case 224:  mlngKeyLength = 224
           Case 192:  mlngKeyLength = 192
           Case 160:  mlngKeyLength = 160
           Case 128:  mlngKeyLength = 128
           Case 96:   mlngKeyLength = 96
           Case 64:   mlngKeyLength = 64
           Case 32:   mlngKeyLength = 32
           Case Else: mlngKeyLength = 256    ' Default value
    End Select

End Property

Public Property Let EncryptRounds(ByVal lngData As Long)

    ' 26-Dec-2008 Added multiple pass functionality.
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngRounds = lngData   ' Good selection
           Case Else:            mlngRounds = 1         ' Default value
    End Select

End Property

Public Property Let PrimaryKeyRounds(ByVal lngData As Long)

    ' 21-Feb-2011 Number of rounds to mix primary key array data
    '             within EvaluateKey() routine
    Select Case lngData
           Case 1 To 5: mlngKeyRounds = lngData   ' Good selection
           Case Else:   mlngKeyRounds = 1         ' Default value
    End Select

End Property

Public Property Let Password(ByVal vntData As Variant)

    ' Used when assigning a value to the password property, on the
    ' left side of an assignment.
    '
    ' Syntax:     X.Password = "User Supplied Password"
    '                            or
    '             X.Password = abytPwd()  <-- Highly recommend using a byte array

    Dim intIndex As Integer

    Const ROUTINE_NAME As String = "Password Property"

    On Error GoTo Password_Error

    Erase mabytPassword()

    ' Is password data in an array
    If CBool(IsArrayInitialized(vntData)) Then

        ReDim mabytPassword(UBound(vntData))

        For intIndex = 0 To UBound(vntData) - 1
            mabytPassword(intIndex) = CByte(vntData(intIndex))
        Next intIndex

    ' Is password data in raw string format
    ElseIf Len(vntData) > 0 Then
        mabytPassword() = StringToByteArray(CStr(vntData))

    ' no data passed
    Else
        InfoMsg "Missing or invalid password." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    End If

Password_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Password_CleanUp

End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Twofish_File
'
' Description:   Function to encrypt/decrypt a file. 
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objTwofish.Twofish_File abytFile(), blnEncrypt               ' Encrypt/Decrypt data
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
'              Modified to handle files larger than 2gb
' 08-MAY-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call for file verification.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' 19-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated progress bar display if using multiple pass
'              functionality for encryption.  See BlockEncrypt() and
'              BlockDecrypt() routines.
' ***************************************************************************
Public Function Twofish_File(ByRef abytFile() As Byte, _
                    Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String  ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim strExtension As String
    Dim hFile1       As Long
    Dim hFile2       As Long
    Dim lngBlockSize As Long
    Dim lngLastBlock As Long
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curAmtLeft   As Currency
    Dim abytData()   As Byte
    Dim abytSize(11) As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "Twofish_File"

    On Error GoTo Twofish_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles   ' Instantiate class object
    ResetVariables                    ' Reset class variables
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()

    If Not EvaluateKey() Then
        GoTo Twofish_File_CleanUp
    End If

    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal

    ' See if a new file is to be created after
    ' performing encryption or decryption.
    ' ex:  Append "ENC" extension to target file
    '      designating the file is encrypted.
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
        GoTo Twofish_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Twofish_File_CleanUp
        End If

        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)  ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                          ' Set pointer to capture the last 12 bytes

            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If

            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 12 bytes of source file
            .API_CloseFile hFile1                             ' close source file

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo Twofish_File_CleanUp
            End If

            MixAppendedData abytSize()                        ' Decrypt appended data
            CopyMemory lngLastBlock, abytSize(0), 4&          ' Size of last block of data without padding [0-3]
            CopyMemory mcurOrigLength, abytSize(4), 8&        ' Capture original file size [4-11]

            ' Decryption error test
            If (lngLastBlock < 0) Or _
               (lngLastBlock > MAX_FILE_SIZE) Or _
               (curAmtLeft < 1) Or _
               (mcurOrigLength < 0) Or _
               (mcurOrigLength > MAX_FILE_SIZE) Then

                InfoMsg "This file cannot be Decrypted, it is" & vbNewLine & _
                        "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        "Max allowed size:  " & Format$(MAX_FILE_SIZE, "#,##0") & " bytes" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                gblnStopProcessing = True
                GoTo Twofish_File_CleanUp

            End If
        End If

        mcurMaxSize = mcurOrigLength * mlngRounds  ' Save file size for progress calcs

        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Twofish_File_CleanUp
        End If

        strTempFile = CreateTempFile()  ' Create a temp target file

        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Twofish_File_CleanUp
        End If

        Do
            Erase abytData()                              ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array

            If blnEncrypt Then
                lngLastBlock = lngBlockSize               ' Save size of last block of data
            End If

            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            curFilePos1 = curFilePos1 + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly

            If blnEncrypt Then

                ' Encrypt data
                If Not EncryptData(abytData()) Then
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If

            Else
                ' Decrypt data
                If DecryptData(abytData()) Then

                    If curAmtLeft < 1 Then
                        ReDim Preserve abytData(lngLastBlock - 1)   ' Resize last data chunk
                    End If

                Else
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If

            End If

            ' Write to target file
            If Not .API_WriteFile(hFile2, curFilePos2, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)    ' Adjust pointers accordingly

            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do  ' exit Do..Loop
            End If

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do   ' exit Do..Loop
            End If

        Loop

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo Twofish_File_CleanUp
        End If

        .API_CloseFile hFile1    ' Close all files opened by this routine
        .API_CloseFile hFile2

        ' Add important data to end of encrypted file
        If blnEncrypt Then

            CopyMemory abytSize(0), lngLastBlock, 4&         ' Save size of last block of data without padding [0-3]
            CopyMemory abytSize(4), mcurOrigLength, 8&       ' Save original file size [4-11]
            .CalcFileSize strTempFile, curFilePos2           ' Get size of target file
            MixAppendedData abytSize()                       ' Encrypt appended data

            ' Excessive DoEvents are to keep the write
            ' process in sync with the CPU speed
            DoEvents
            .OpenReadWrite strTempFile, hFile2               ' open target file
            .API_WriteFile hFile2, curFilePos2, abytSize()   ' Append to target file
            .API_CloseFile hFile2                            ' Close target file
            DoEvents
        Else
            ' Final decryption verification
            curAmtLeft = 0@                         ' Zero variable
            .CalcFileSize strTempFile, curAmtLeft   ' Get size of decrypted file

            ' Compare current file size with
            ' what it is supposed to be
            If (curAmtLeft <> mcurOrigLength) Then

                InfoMsg "This file cannot be Decrypted." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                gblnStopProcessing = True
                GoTo Twofish_File_CleanUp

            End If
        End If
    End With

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

Twofish_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Close all files opened by this routine
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory

    ResetVariables    ' Reset class variables
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Twofish_File = False
    Else
        Twofish_File = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Twofish_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Twofish_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Twofish_String
'
' Description:   Function to encrypt/decrypt a string of data.
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)        ' convert string data to byte array
'                objTwofish.Twofish_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added call to MixAppendedData() routine to encrypt/decrypt
'              size information that is to be appended to original data.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Split size testing as per Joe Sova's suggestion.
' ***************************************************************************
Public Function Twofish_String(ByRef abytData() As Byte, _
                      Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte

    Const ROUTINE_NAME As String = "Twofish_String"

    On Error GoTo Twofish_String_Error

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

    ResetVariables     ' Reset class variables
    Erase abytSize()   ' Always start with empty arrays

    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds   ' Save file size for progress calcs
    mcurAccrued = 0@                                    ' Init accrued progress

    If Not EvaluateKey() Then
        GoTo Twofish_String_CleanUp
    End If

    ' encrypt the data
    If blnEncrypt Then

        mcurOrigLength = CCur(UBound(abytData))         ' Get string length for progress display
        CopyMemory abytSize(0), mcurOrigLength, 8&      ' Copy original length to temp array
        MixAppendedData abytSize()                      ' Encrypt appended data

        If EncryptData(abytData()) Then

            curDataLength = UBound(abytData) + 1        ' data length after encryption
            ReDim Preserve abytData(curDataLength + 7)  ' Resize to hold original length

            ' Append original data length to end of encrypted data
            CopyMemory abytData(curDataLength), abytSize(0), 8&

        Else
            gblnStopProcessing = True
        End If

    Else
        ' Decrypt the data
        CopyMemory abytSize(0), abytData(UBound(abytData) - 7), 8&   ' capture the original data length
        ReDim Preserve abytData(UBound(abytData) - 8)                ' Adjust data length
        MixAppendedData abytSize()                                   ' Decrypt appended data
        CopyMemory mcurOrigLength, abytSize(0), 8&                   ' capture the original data length

        ' Decrypt the data
        If DecryptData(abytData()) Then

            ReDim Preserve abytData(mcurOrigLength)  ' Resize data to exact length
            curDataLength = CCur(UBound(abytData))   ' data length after decryption

            ' Compare data length
            If curDataLength <> mcurOrigLength Then

                InfoMsg "Encrypted data string is corrupted.  Cannot Decrypt." & _
                        vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curDataLength, "#,##0") & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                gblnStopProcessing = True

            End If
        Else
            gblnStopProcessing = True
        End If

    End If

Twofish_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Twofish_String = False
    Else
        Twofish_String = True
    End If

    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Twofish_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Twofish_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Twofish_File()
    '           Twofish_String()

    Dim lngHold   As Long
    Dim lngKeyLen As Long
    Dim abytPwd() As Byte

    Const ROUTINE_NAME As String = "EvaluateKey"

    On Error GoTo EvaluateKey_Error

    EvaluateKey = False  ' Preset flag

    If Not CBool(IsArrayInitialized(mabytPassword)) Then
        InfoMsg "Need a password before data can be encrypted/decrypted." & _
                vbNewLine & vbNewLine & _
                "Min length - " & MIN_PWD_LENGTH & vbNewLine & _
                "Max length - " & MAX_PWD_LENGTH & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If

    LoadWorkArrays                                        ' Load constants arrays
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  4 bytes = 32 bits \ 8 bits
    Erase abytPwd()                                      ' Always start with an empty array
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    lngHold = mlngRounds                                 ' 02-Nov-2009 Capture number of rounds for block encryption
    EncryptRounds = mlngKeyRounds                        ' 02-Nov-2009 Temp reset number of rounds of block
                                                         '             encryption (1-5 rounds is a good mix)
    ProcessKey abytPwd(), lngKeyLen                      ' Create key-dependent sboxes

EvaluateKey_CleanUp:
    Erase abytPwd()           ' Always empty arrays when not needed
    EncryptRounds = lngHold   ' 02-Nov-2009 Reset number of rounds of block
                              '             encryption back to original value
    ' An error occurred or user
    ' opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EvaluateKey = False
    Else
        EvaluateKey = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ResetVariables
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Twofish_File()
    '           Twofish_String()

    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim alngDWord(6)  As Long
    Dim alngCipher(4) As Long
    Dim abytOutput()  As Byte

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array

    ' The encoded length may have to be padded out.
    ' This normally does not take place until the
    ' last chunk of data is to be processed.
    If (lngDataLength Mod BLOCK_SIZE <> 0) Then
        lngDataLength = lngDataLength + BLOCK_SIZE - (lngDataLength Mod BLOCK_SIZE)
        ReDim Preserve abytData(lngDataLength - 1)
    End If

    ReDim abytOutput(lngDataLength)  ' Resize the output array
    CopyMemory abytOutput(0), abytData(0), lngDataLength

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Get the next block
        CopyMemory alngDWord(0), abytOutput(lngOffSet), BLOCK_SIZE

        ' XOR the plaintext with the previous
        ' ciphertext (CBC, Cipher-Block Chaining)
        alngDWord(0) = alngDWord(0) Xor alngCipher(0)
        alngDWord(1) = alngDWord(1) Xor alngCipher(1)
        alngDWord(2) = alngDWord(2) Xor alngCipher(2)
        alngDWord(3) = alngDWord(3) Xor alngCipher(3)

        ' Encrypt block of data
        BlockEncrypt alngDWord()

        ' Store the new block
        CopyMemory abytOutput(lngOffSet), alngDWord(0), BLOCK_SIZE

        ' Store the cipherblock (for CBC)
        CopyMemory alngCipher(0), alngDWord(0), BLOCK_SIZE

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + CCur(BLOCK_SIZE * mlngRounds)
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)

    Next lngOffSet

    ' Prepare original array for return
    ReDim abytData(lngDataLength)

    ' Copy encrypted output array back to original array
    CopyMemory abytData(0), abytOutput(0), lngDataLength
    ReDim Preserve abytData(lngDataLength - 1)

EncryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False
    Else
        EncryptData = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Twofish_File()
    '           Twofish_String()

    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim alngDWord(6)  As Long
    Dim alngCipher(4) As Long
    Dim abytOutput()  As Byte

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array

    ReDim abytOutput(lngDataLength)
    CopyMemory abytOutput(0), abytData(0), lngDataLength

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Get the next block
        CopyMemory alngDWord(2), abytOutput(lngOffSet), BLOCK_SIZE

        ' Decrypt the block
        BlockDecrypt alngDWord()

        ' XOR with the previous cipherblock
        alngDWord(0) = alngDWord(0) Xor alngCipher(0)
        alngDWord(1) = alngDWord(1) Xor alngCipher(1)
        alngDWord(2) = alngDWord(2) Xor alngCipher(2)
        alngDWord(3) = alngDWord(3) Xor alngCipher(3)

        ' Store the current ciphertext to use
        ' XOR with the next block plaintext
        CopyMemory alngCipher(0), abytOutput(lngOffSet), BLOCK_SIZE

        ' Store the block
        CopyMemory abytOutput(lngOffSet), alngDWord(0), BLOCK_SIZE

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + CCur(BLOCK_SIZE * mlngRounds)
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)

    Next lngOffSet

    ' Prepare the return array
    lngDataLength = UBound(abytOutput)
    ReDim abytData(lngDataLength)

    ' Copy decrypted data back to original array
    CopyMemory abytData(0), abytOutput(0), lngDataLength
    ReDim Preserve abytData(lngDataLength - 1)

DecryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False
    Else
        DecryptData = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

Private Sub BlockEncrypt(ByRef alngDWord() As Long)

    ' Called by EncryptData()

    Dim lngIndex As Long
    Dim lngTemp1 As Long
    Dim lngTemp2 As Long
    Dim lngCount As Long
    Dim lngRound As Long

    alngDWord(0) = alngDWord(0) Xor malngKey(4)
    alngDWord(1) = alngDWord(1) Xor malngKey(5)
    alngDWord(2) = alngDWord(2) Xor malngKey(6)
    alngDWord(3) = alngDWord(3) Xor malngKey(7)

    With mobjMath32
        For lngRound = 1 To mlngRounds

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

            lngCount = ROUND_SUBKEYS

            For lngIndex = 0 To (CIPHER_ROUNDS - 1) Step 2

                lngTemp1 = Fe32(alngDWord(0), 0)
                lngTemp2 = Fe32(alngDWord(1), 3)
                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)

                alngDWord(2) = .w32Rotate(alngDWord(2) Xor _
                               (.UnsignedAdd(lngTemp1, malngKey(lngCount))), 1)
                lngCount = lngCount + 1

                alngDWord(3) = .w32Rotate(alngDWord(3), 31) Xor _
                               (.UnsignedAdd(.UnsignedAdd(lngTemp1, lngTemp2), malngKey(lngCount)))
                lngCount = lngCount + 1

                lngTemp1 = Fe32(alngDWord(2), 0)
                lngTemp2 = Fe32(alngDWord(3), 3)
                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)

                alngDWord(0) = .w32Rotate(alngDWord(0) Xor _
                               (.UnsignedAdd(lngTemp1, malngKey(lngCount))), 1)
                lngCount = lngCount + 1

                alngDWord(1) = .w32Rotate(alngDWord(1), 31) Xor _
                               (.UnsignedAdd(.UnsignedAdd(lngTemp1, lngTemp2), malngKey(lngCount)))
                lngCount = lngCount + 1

            Next lngIndex

        Next lngRound
    End With

    alngDWord(2) = alngDWord(2) Xor malngKey(4)
    alngDWord(3) = alngDWord(3) Xor malngKey(5)
    alngDWord(4) = alngDWord(0) Xor malngKey(6)
    alngDWord(5) = alngDWord(1) Xor malngKey(7)

    CopyMemory alngDWord(0), alngDWord(2), 16&

End Sub

Private Sub BlockDecrypt(ByRef alngDWord() As Long)

    ' Called by DecryptData()

    Dim lngIndex As Long
    Dim lngTemp1 As Long
    Dim lngTemp2 As Long
    Dim lngCount As Long
    Dim lngRound As Long

    alngDWord(2) = alngDWord(2) Xor malngKey(4)
    alngDWord(3) = alngDWord(3) Xor malngKey(5)
    alngDWord(0) = alngDWord(4) Xor malngKey(6)
    alngDWord(1) = alngDWord(5) Xor malngKey(7)

    With mobjMath32
        For lngRound = 1 To mlngRounds

            ' An error occurred or user opted to STOP processing
            If gblnStopProcessing Then
                Exit For    ' exit For..Next loop
            End If

            lngCount = ROUND_SUBKEYS + 2 * CIPHER_ROUNDS - 1

            For lngIndex = 0 To CIPHER_ROUNDS - 1 Step 2

                lngTemp1 = Fe32(alngDWord(2), 0)
                lngTemp2 = Fe32(alngDWord(3), 3)
                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)

                alngDWord(1) = .w32Rotate(alngDWord(1) Xor _
                               (.UnsignedAdd(.UnsignedAdd(lngTemp1, lngTemp2), malngKey(lngCount))), 1)
                lngCount = lngCount - 1

                alngDWord(0) = .w32Rotate(alngDWord(0), 31) Xor _
                               (.UnsignedAdd(lngTemp1, malngKey(lngCount)))
                lngCount = lngCount - 1

                lngTemp1 = Fe32(alngDWord(0), 0)
                lngTemp2 = Fe32(alngDWord(1), 3)
                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)

                alngDWord(3) = .w32Rotate(alngDWord(3) Xor _
                               (.UnsignedAdd(.UnsignedAdd(lngTemp1, lngTemp2), malngKey(lngCount))), 1)
                lngCount = lngCount - 1

                alngDWord(2) = .w32Rotate(alngDWord(2), 31) Xor _
                               (.UnsignedAdd(lngTemp1, malngKey(lngCount)))
                lngCount = lngCount - 1

            Next lngIndex

        Next lngRound
    End With

    alngDWord(0) = alngDWord(0) Xor malngKey(4)
    alngDWord(1) = alngDWord(1) Xor malngKey(5)
    alngDWord(2) = alngDWord(2) Xor malngKey(6)
    alngDWord(3) = alngDWord(3) Xor malngKey(7)

End Sub

Private Function Fe32(ByVal lngValue As Long, _
                      ByVal lngPosition As Long) As Long

    ' Called by BlockEncrypt()
    '           BlockDecrypt()

    Dim abytTemp(4) As Byte

    ' Extract the byte sequence
    CopyMemory abytTemp(0), lngValue, 4&

    ' Calculate the FE32 function
    Fe32 = malngSBox(2 * abytTemp(lngPosition Mod 4)) Xor _
           malngSBox(2 * abytTemp((lngPosition + 1) Mod 4) + 1) Xor _
           malngSBox(&H200 + 2 * abytTemp((lngPosition + 2) Mod 4)) Xor _
           malngSBox(&H200 + 2 * abytTemp((lngPosition + 3) Mod 4) + 1)

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
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Called by Twofish_File()
    '           Twofish_String()

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select

    RaiseEvent CurrentBlockSize(GetBlockSize)

End Function

' ***************************************************************************
' Routine:       LoadWorkArrays
'
' Description:   Load primary key arrays.
'
'                Enhanced the P-Box (primary key) array data by making
'                them more dynamic. This was obtained by mixing ASCII
'                values (0-255) in hex format using Knuth Shuffle
'                100-200 iterations with no duplicates.
'
'                See application BuildTables.vbp to create additional
'                values.  See Twofish_Pbox.txt file for additional
'                tables.
'
' --------------------------------------------------------------------------
'    ' Original data for mabytPBox(0, 255) array (Do not modify)
'    strData = strData & "A9 67 B3 E8 04 FD A3 76 9A 92 80 78 E4 DD D1 38 "
'    strData = strData & "0D C6 35 98 18 F7 EC 6C 43 75 37 26 FA 13 94 48 "
'    strData = strData & "F2 D0 8B 30 84 54 DF 23 19 5B 3D 59 F3 AE A2 82 "
'    strData = strData & "63 01 83 2E D9 51 9B 7C A6 EB A5 BE 16 0C E3 61 "
'    strData = strData & "C0 8C 3A F5 73 2C 25 0B BB 4E 89 6B 53 6A B4 F1 "
'    strData = strData & "E1 E6 BD 45 E2 F4 B6 66 CC 95 03 56 D4 1C 1E D7 "
'    strData = strData & "FB C3 8E B5 E9 CF BF BA EA 77 39 AF 33 C9 62 71 "
'    strData = strData & "81 79 09 AD 24 CD F9 D8 E5 C5 B9 4D 44 08 86 E7 "
'    strData = strData & "A1 1D AA ED 06 70 B2 D2 41 7B A0 11 31 C2 27 90 "
'    strData = strData & "20 F6 60 FF 96 5C B1 AB 9E 9C 52 1B 5F 93 0A EF "
'    strData = strData & "91 85 49 EE 2D 4F 8F 3B 47 87 6D 46 D6 3E 69 64 "
'    strData = strData & "2A CE CB 2F FC 97 05 7A AC 7F D5 1A 4B 0E A7 5A "
'    strData = strData & "28 14 3F 29 88 3C 4C 02 B8 DA B0 17 55 1F 8A 7D "
'    strData = strData & "57 C7 8D 74 B7 C4 9F 72 7E 15 22 12 58 07 99 34 "
'    strData = strData & "6E 50 DE 68 65 BC DB F8 C8 A8 2B 40 DC FE 32 A4 "
'    strData = strData & "CA 10 21 F0 D3 5D 0F 00 6F 9D 36 42 4A 5E C1 E0 "
'
'    ' Original data for mabytPBox(1, 255) array (Do not modify)
'    strData = strData & "75 F3 C6 F4 DB 7B FB C8 4A D3 E6 6B 45 7D E8 4B "
'    strData = strData & "D6 32 D8 FD 37 71 F1 E1 30 0F F8 1B 87 FA 06 3F "
'    strData = strData & "5E BA AE 5B 8A 00 BC 9D 6D C1 B1 0E 80 5D D2 D5 "
'    strData = strData & "A0 84 07 14 B5 90 2C A3 B2 73 4C 54 92 74 36 51 "
'    strData = strData & "38 B0 BD 5A FC 60 62 96 6C 42 F7 10 7C 28 27 8C "
'    strData = strData & "13 95 9C C7 24 46 3B 70 CA E3 85 CB 11 D0 93 B8 "
'    strData = strData & "A6 83 20 FF 9F 77 C3 CC 03 6F 08 BF 40 E7 2B E2 "
'    strData = strData & "79 0C AA 82 41 3A EA B9 E4 9A A4 97 7E DA 7A 17 "
'    strData = strData & "66 94 A1 1D 3D F0 DE B3 0B 72 A7 1C EF D1 53 3E "
'    strData = strData & "8F 33 26 5F EC 76 2A 49 81 88 EE 21 C4 1A EB D9 "
'    strData = strData & "C5 39 99 CD AD 31 8B 01 18 23 DD 1F 4E 2D F9 48 "
'    strData = strData & "4F F2 65 8E 78 5C 58 19 8D E5 98 57 67 7F 05 64 "
'    strData = strData & "AF 63 B6 FE F5 B7 3C A5 CE E9 68 44 E0 4D 43 69 "
'    strData = strData & "29 2E AC 15 59 A8 0A 9E 6E 47 DF 34 35 6A CF DC "
'    strData = strData & "22 C9 C0 9B 89 D4 ED AB 12 A2 0D 52 BB 02 2F A9 "
'    strData = strData & "D7 61 1E B4 50 04 F6 C2 16 25 86 56 55 09 BE 91 "
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2017  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Sub LoadWorkArrays()

    Dim strData    As String
    Dim avntData   As Variant
    Dim lngIndex   As Long
    Dim lngPointer As Long
    Dim alngM1()   As Long
    Dim alngM2()   As Long
    Dim alngM3()   As Long

    ReDim alngM1(2)
    ReDim alngM2(2)
    ReDim alngM3(2)
    ReDim malngMDS(0 To 3, MAX_BYTE)
    ReDim mabytPBox(0 To 1, MAX_BYTE)

    avntData = Empty         ' Always start with empty variants
    strData = vbNullString   ' Verify data string is empty

    ' Data for mabytPBox(0, 255) array is ASCII
    ' decimal 0-255 mixed 100-200 times with no
    ' duplicates displayed here in hex format
    strData = strData & "15 34 D8 6C 71 97 53 45 B0 4B 4C 27 96 D1 30 39 "
    strData = strData & "61 1E 2C 7B 1C 51 79 18 03 2A 84 DD FB 72 C9 CC "
    strData = strData & "33 5F 3E 21 76 E5 50 DF EE 9F CE D4 D0 DB 5A 99 "
    strData = strData & "2E 4D C8 6A F6 9D B1 86 A1 83 D5 56 6B C6 3B 5C "
    strData = strData & "10 AA 0A B2 A5 5E 67 7E E4 E0 BB 65 9A 90 BC 04 "
    strData = strData & "3D 57 D9 0F D2 32 7C 1B 8B FC AC F9 BF A6 22 BA "
    strData = strData & "3C 82 1A 62 55 9E E2 7D CF 19 05 06 02 C4 D3 8C "
    strData = strData & "EB 0D 0C 47 6D CB 91 14 69 8D 36 EA 81 FA C3 B4 "
    strData = strData & "4E 89 D7 B5 7F 93 52 17 6F 88 54 5D A8 F5 B3 B7 "
    strData = strData & "35 E8 D6 CD 60 EF A0 E7 8E 94 11 95 F7 A2 25 37 "
    strData = strData & "48 8F F3 9B 09 66 4A 98 E3 EC C5 38 63 B9 13 AF "
    strData = strData & "BE DC 70 C0 85 FE 3A 77 20 F0 7A A9 2B 58 29 A3 "
    strData = strData & "AE DE 74 0E 3F 44 01 B8 AB 24 68 F1 46 AD 23 C1 "
    strData = strData & "FD 26 16 43 00 4F A7 1D DA 75 2D CA B6 40 5B 41 "
    strData = strData & "42 E6 E1 1F F8 28 12 08 ED 49 E9 07 9C 8A F4 31 "
    strData = strData & "A4 0B 73 87 80 78 59 BD 64 FF 2F 92 C2 6E C7 F2 "

    ' Separate into a variant array
    avntData = Split(strData, Chr$(32))

    ' Transfer variant array to the P-box array
    For lngIndex = 0 To (MAX_BYTE - 1)
        mabytPBox(0, lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex

    avntData = Empty         ' Always start with empty variants
    strData = vbNullString   ' Verify data string is empty

    ' Data for mabytPBox(1, 255) array is ASCII
    ' decimal 0-255 mixed 100-200 times with no
    ' duplicates displayed here in hex format
    strData = strData & "F2 67 69 34 96 1D EC 5E B6 60 80 4C 63 55 39 CC "
    strData = strData & "7A 50 73 F9 14 78 E6 9D CF 53 1E 5D 3A E1 00 CE "
    strData = strData & "64 5B D8 DD B5 51 82 CD 27 87 C6 3B 59 C2 49 3E "
    strData = strData & "BD 90 24 AF E9 CA 72 66 4F AD F3 7E 29 28 7F F7 "
    strData = strData & "CB E7 DB 5F 94 62 3C 95 0F F4 DA EF EE 6C 42 BE "
    strData = strData & "93 2A 4D 45 11 0C A2 76 E4 2B FB C9 2D 3D 9B 75 "
    strData = strData & "F5 17 FE A8 26 DC D7 C7 56 EB 4B D2 EA FA 58 10 "
    strData = strData & "C8 2F A0 7B 47 86 C4 9C 05 0A 8B 21 A4 06 E3 A7 "
    strData = strData & "02 A3 37 18 1F E0 77 B1 33 D9 B8 70 54 0B B0 1C "
    strData = strData & "0D D1 97 6A A9 48 35 38 85 F1 15 A5 DF 0E 8C F6 "
    strData = strData & "7C 71 E2 ED 01 31 B3 9A E8 09 FC 13 BC D3 9F B7 "
    strData = strData & "C5 FD BF 52 DE 20 68 AB 12 5A 88 6B 08 C1 46 83 "
    strData = strData & "B2 2C FF 8E 98 99 16 F0 C0 2E 7D AE B9 79 19 5C "
    strData = strData & "40 57 91 4A E5 03 A1 92 32 8A B4 AA C3 43 74 30 "
    strData = strData & "6D 61 22 81 6F 1A D0 44 BA 41 3F AC 07 D4 9E 89 "
    strData = strData & "36 4E 8D A6 23 D5 84 BB 8F 25 1B 04 65 D6 F8 6E "

    avntData = Split(strData, Chr$(32))

    ' Transfer variant array to the P-box array
    For lngIndex = 0 To (MAX_BYTE - 1)
        mabytPBox(1, lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex

    avntData = Empty         ' Always empty variants when not needed
    strData = vbNullString   ' Verify data string is empty

    ' Load malngMDS() array
    With mobjMath32
        For lngIndex = 0 To (MAX_BYTE - 1)

            lngPointer = mabytPBox(0, lngIndex)
            alngM3(0) = lngPointer
            alngM1(0) = lngPointer Xor (.w32Shift(lngPointer, -2))
            alngM2(0) = lngPointer Xor _
                        (.w32Shift(lngPointer, -1)) Xor _
                        (.w32Shift(lngPointer, -2))

            lngPointer = mabytPBox(1, lngIndex)
            alngM3(1) = lngPointer
            alngM1(1) = lngPointer Xor (.w32Shift(lngPointer, -2))
            alngM2(1) = lngPointer Xor _
                        (.w32Shift(lngPointer, -1)) Xor _
                        (.w32Shift(lngPointer, -2))

            malngMDS(0, lngIndex) = alngM3(1) Or _
                                    (.w32Shift(alngM1(1), 8)) Or _
                                    (.w32Shift(alngM2(1), 16)) Or _
                                    (.w32Shift(alngM2(1), 24))

            malngMDS(1, lngIndex) = alngM2(0) Or _
                                    (.w32Shift(alngM2(0), 8)) Or _
                                    (.w32Shift(alngM1(0), 16)) Or _
                                    (.w32Shift(alngM3(0), 24))

            malngMDS(2, lngIndex) = alngM1(1) Or _
                                    (.w32Shift(alngM2(1), 8)) Or _
                                    (.w32Shift(alngM3(1), 16)) Or _
                                    (.w32Shift(alngM2(1), 24))

            malngMDS(3, lngIndex) = alngM1(0) Or _
                                    (.w32Shift(alngM3(0), 8)) Or _
                                    (.w32Shift(alngM2(0), 16)) Or _
                                    (.w32Shift(alngM1(0), 24))

        Next lngIndex
    End With

LoadWorkBoxes_CleanUp:
    avntData = Empty         ' Always empty variants when not needed
    strData = vbNullString   ' Verify data string is empty

End Sub

Private Sub ProcessKey(ByRef abytPwdKey() As Byte, _
                       ByVal lngKeyLength As Long)

    ' Called by EvaluateKey()

    Dim lngMax        As Long
    Dim lngIdx        As Long
    Dim lngIndex      As Long
    Dim lngTemp1      As Long
    Dim lngTemp2      As Long
    Dim lngK64Count   As Long
    Dim lngSubKeyCnt  As Long
    Dim alngK32E()    As Long
    Dim alngK32O()    As Long
    Dim alngSBoxKey() As Long
    Dim abytTemp()    As Byte
    Dim abytKey(4, 4) As Byte

    Const SK_STEP As Long = &H2020202
    Const SK_BUMP As Long = &H1010101

    lngK64Count = lngKeyLength \ 8
    lngSubKeyCnt = ROUND_SUBKEYS + (2 * CIPHER_ROUNDS)
    lngMax = IIf(lngKeyLength < 32, lngKeyLength * 8 - 1, lngKeyLength)

    ReDim alngSBoxKey(lngMax)     ' Size temp work arrays
    ReDim abytTemp(lngMax)
    ReDim alngK32E(lngMax)
    ReDim alngK32O(lngMax)

    ReDim malngSBox(1024)         ' Size module arrays
    ReDim malngKey(lngSubKeyCnt)

    ' 05-Dec-2013  Modified for possible additional
    '              mixing of primary key and s-box arrays
    With mobjMath32
        For lngIdx = 1 To mlngRounds     ' Possible additional mixing

            For lngIndex = 0 To lngMax   ' Prepare malngKey() and malngSBox() arrays

                If lngIdx = 1 Then
                    CopyMemory alngK32E(lngIndex), abytPwdKey((lngIndex * 8) Mod lngKeyLength), 4&
                    CopyMemory alngK32O(lngIndex), abytPwdKey((lngIndex * 8 + 4) Mod lngKeyLength), 4&
                End If

                lngTemp1 = RS_Rem(alngK32O(lngIndex))
                lngTemp1 = RS_Rem(lngTemp1)
                lngTemp1 = RS_Rem(lngTemp1)
                lngTemp1 = RS_Rem(lngTemp1)

                lngTemp1 = lngTemp1 Xor alngK32E(lngIndex)
                lngTemp1 = RS_Rem(lngTemp1)
                lngTemp1 = RS_Rem(lngTemp1)
                lngTemp1 = RS_Rem(lngTemp1)

                ' load array from last element to first (ex: 127 to 0)
                alngSBoxKey(lngMax - lngIndex) = RS_Rem(lngTemp1)

            Next lngIndex

            lngMax = ((lngSubKeyCnt / 2) - 1)

            For lngIndex = 0 To lngMax

                lngTemp1 = F32(lngK64Count, lngIndex * SK_STEP, alngK32E())
                lngTemp2 = F32(lngK64Count, lngIndex * SK_STEP + SK_BUMP, alngK32O())
                lngTemp2 = (.w32Shift(lngTemp2, 8)) Or (.w32Shift(lngTemp2, -24))

                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)
                malngKey(2 * lngIndex) = lngTemp1

                lngTemp1 = .UnsignedAdd(lngTemp1, lngTemp2)
                malngKey(2 * lngIndex + 1) = (.w32Shift(lngTemp1, 9)) Or _
                                             (.w32Shift(lngTemp1, -23))

            Next lngIndex

            ' Initial load of temp key array
            CopyMemory abytKey(0, 0), alngSBoxKey(0), 16&

            For lngIndex = 0 To (MAX_BYTE - 1)

                If ((lngK64Count And 3) = 1) Then

                    malngSBox(2 * lngIndex) = malngMDS(0, mabytPBox(0, lngIndex) Xor abytKey(0, 0))
                    malngSBox(2 * lngIndex + 1) = malngMDS(1, mabytPBox(0, lngIndex) Xor abytKey(1, 0))
                    malngSBox(2 * lngIndex + 512) = malngMDS(2, mabytPBox(1, lngIndex) Xor abytKey(2, 0))
                    malngSBox(2 * lngIndex + 512 + 1) = malngMDS(3, mabytPBox(1, lngIndex) Xor abytKey(3, 0))

                Else

                    abytTemp(0) = lngIndex
                    abytTemp(1) = lngIndex
                    abytTemp(2) = lngIndex
                    abytTemp(3) = lngIndex

                    If ((lngK64Count And 3) = 0) Then
                        abytTemp(0) = mabytPBox(1, abytTemp(0)) Xor abytKey(0, 3)
                        abytTemp(1) = mabytPBox(0, abytTemp(1)) Xor abytKey(1, 3)
                        abytTemp(2) = mabytPBox(0, abytTemp(2)) Xor abytKey(2, 3)
                        abytTemp(3) = mabytPBox(1, abytTemp(3)) Xor abytKey(3, 3)
                    End If

                    ' (exception = True)
                    If ((lngK64Count And 3) = 3) Or ((lngK64Count And 3) = 0) Then
                        abytTemp(0) = mabytPBox(1, abytTemp(0)) Xor abytKey(0, 2)
                        abytTemp(1) = mabytPBox(1, abytTemp(1)) Xor abytKey(1, 2)
                        abytTemp(2) = mabytPBox(0, abytTemp(2)) Xor abytKey(2, 2)
                        abytTemp(3) = mabytPBox(0, abytTemp(3)) Xor abytKey(3, 2)
                    End If

                    malngSBox(2 * lngIndex) = malngMDS(0, mabytPBox(0, mabytPBox(0, abytTemp(0)) Xor abytKey(0, 1)) Xor abytKey(0, 0))
                    malngSBox(2 * lngIndex + 1) = malngMDS(1, mabytPBox(0, mabytPBox(1, abytTemp(1)) Xor abytKey(1, 1)) Xor abytKey(1, 0))
                    malngSBox(2 * lngIndex + 512) = malngMDS(2, mabytPBox(1, mabytPBox(0, abytTemp(2)) Xor abytKey(2, 1)) Xor abytKey(2, 0))
                    malngSBox(2 * lngIndex + 512 + 1) = malngMDS(3, mabytPBox(1, mabytPBox(1, abytTemp(3)) Xor abytKey(3, 1)) Xor abytKey(3, 0))

                End If

            Next lngIndex
        Next lngIdx
    End With

End Sub

Private Function RS_Rem(ByVal lngValue As Long) As Long

    ' Called by ProcessKey()

    Dim lngTemp  As Long
    Dim lngWord1 As Long
    Dim lngWord2 As Long

    With mobjMath32
        lngTemp = (.w32Shift(lngValue, -24) And &HFF)

        lngWord1 = ((.w32Shift(lngTemp, 1) Xor _
                   (lngTemp And &H80) / &H80 * &H14D) And &HFF)

        lngWord2 = (.w32Shift(lngTemp, -1) Xor _
                   ((lngTemp And &H1) * .w32Shift(&H14D, -1)) Xor lngWord1)

        RS_Rem = .w32Shift(lngValue, 8) Xor _
                 .w32Shift(lngWord2, 24) Xor _
                 .w32Shift(lngWord1, 16) Xor _
                 .w32Shift(lngWord2, 8) Xor lngTemp
    End With

End Function

Private Function F32(ByVal lngCount As Long, _
                     ByVal lngValue As Long, _
                     ByRef alngData() As Long) As Long

    ' Called by ProcessKey()

    Dim abytTemp(4)   As Byte
    Dim abytKey(4, 4) As Byte

    CopyMemory abytTemp(0), lngValue, 4&
    CopyMemory abytKey(0, 0), alngData(0), 16&

    If ((lngCount And 3) = 1) Then

        F32 = malngMDS(0, mabytPBox(0, abytTemp(0)) Xor abytKey(0, 0)) Xor _
              malngMDS(1, mabytPBox(0, abytTemp(1)) Xor abytKey(1, 0)) Xor _
              malngMDS(2, mabytPBox(1, abytTemp(2)) Xor abytKey(2, 0)) Xor _
              malngMDS(3, mabytPBox(1, abytTemp(3)) Xor abytKey(3, 0))
    Else
        If ((lngCount And 3) = 0) Then
            abytTemp(0) = mabytPBox(1, abytTemp(0)) Xor abytKey(0, 3)
            abytTemp(1) = mabytPBox(0, abytTemp(1)) Xor abytKey(1, 3)
            abytTemp(2) = mabytPBox(0, abytTemp(2)) Xor abytKey(2, 3)
            abytTemp(3) = mabytPBox(1, abytTemp(3)) Xor abytKey(3, 3)
        End If

        If ((lngCount And 3) = 3) Or ((lngCount And 3) = 0) Then
            abytTemp(0) = mabytPBox(1, abytTemp(0)) Xor abytKey(0, 2)
            abytTemp(1) = mabytPBox(1, abytTemp(1)) Xor abytKey(1, 2)
            abytTemp(2) = mabytPBox(0, abytTemp(2)) Xor abytKey(2, 2)
            abytTemp(3) = mabytPBox(0, abytTemp(3)) Xor abytKey(3, 2)
        End If

        F32 = malngMDS(0, mabytPBox(0, mabytPBox(0, abytTemp(0)) Xor abytKey(0, 1)) Xor abytKey(0, 0)) Xor _
              malngMDS(1, mabytPBox(0, mabytPBox(1, abytTemp(1)) Xor abytKey(1, 1)) Xor abytKey(1, 0)) Xor _
              malngMDS(2, mabytPBox(1, mabytPBox(0, abytTemp(2)) Xor abytKey(2, 1)) Xor abytKey(2, 0)) Xor _
              malngMDS(3, mabytPBox(1, mabytPBox(1, abytTemp(3)) Xor abytKey(3, 1)) Xor abytKey(3, 0))
    End If

End Function

Private Sub ResetVariables()

    ' Always empty arrays and variants
    ' when not needed.  Saves resources.
    Erase mabytPBox()
    Erase malngKey()
    Erase malngSBox()
    Erase malngMDS()

    mcurOrigLength = 0@
    mcurCurrByteCnt = 0@

End Sub

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables
    Erase mabytPassword()

    Set mobjMath32 = Nothing
    Set mobjMath32 = New cMath32

    ' Preset property values
    KeyLength = 256        ' Default value
    EncryptRounds = 1      ' Number of rounds of encryption
    PrimaryKeyRounds = 1   ' Number of rounds to mix primary key array
    CreateNewFile = True   ' Default is to create a new output file

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ResetVariables
    Erase mabytPassword()
    Set mobjMath32 = Nothing

End Sub
