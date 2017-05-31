VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBlowFish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cBlowfish  (clsBlowfish.cls)
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
' Description:   Blowfish is a block cipher that was designed in 1993 by
'                Bruce Schneier as a fast, free alternative to existing
'                encryption algorithms. Like Blowfish it can be used as
'                a convenient and much faster substitute for DES. Since
'                then Blowfish has been extensively analyzed and no
'                significant weaknesses have been found. It is considered
'                to be a strong algorithm and has been implemented in over
'                130 commercial applications. Bruce Schneier was also one
'                of the designers of the Blowfish algorithm, which is one of
'                the five finalists selected by NIST in the Advanced Encryption
'                Standard (AES) competition. Blowfish was developed as a
'                successor to the Blowfish cipher, which does not meet the
'                design requirements of the AES.
'
'                Blowfish is a 64-bit block cipher, meaning that data is
'                encrypted and decrypted in 64-bit chunks. The key length
'                can vary from 32 to 448 bits. The algorithm uses 16 rounds,
'                or iterations of the main algorithm. It has been found that
'                the number of rounds is exponentially proportional to the
'                amount of time required to find a key using a brute-force
'                attack. So as the number of rounds increases, the security
'                of the algorithm increases exponentially.
'
'                Private Encryptor's implementation of Blowfish allows the
'                key to vary from 32 to 448 bits, as per the algorithm
'                specification.
'
'                The relative strength of the encryption algorithm is based on
'                key length.  Bruce Schneier, creator of the Blowfish encryption
'                algorithm, has calculated that according to what we know of
'                quantum mechanics today, that the entire energy output of the
'                sun is insufficient to break a 192-bit key.
'
'                The detailed description of the Blowfish algorithm is contained
'                in the original Blowfish paper, written by Bruce Schneier.
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
'
'                - Password key array (p-box) has been enhanced by using
'                  the first thirty-two bits (8 chars) of the fractional
'                  parts of the SQUARE ROOT of the 100th through 351st
'                  prime numbers.  See LoadWorkArrays() routine.
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
' References:    Original algorithm Written by Bruce Schneier
'                http://www.counterpane.com/blowfish.html
'
'                Tropical software (Algorithm description)
'                http://www.tropsoft.com/strongenc/blowfish.htm
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
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 08-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added GetWord() and PutWord() routines.  Fixed a logic bug
'              in BlockDecrypt() routine.
' 05-Sep-2008  Kenneth Ives  kenaso@tx.rr.com
'              Renamed lNewValue() to CalcNewValue().  Modifed CalcNewValue()
'              routine for easier maintenance.
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified Blowfish_File() and Blowfish_String() routines to
'              encrypt appended data information for additional security.
'              Moved ExpandData() routine to basCommon.bas module.
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet design standards. Set key
'              length to increments of 32 bits for 128 to 448.
'              Modified EvaluateKey() routine to use utilize the new key
'              lengths.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed OffSet range in EncryptData() and DecryptData() routines.
'              Moved GetBlockSize() routine to common module.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in Blowfish_String().
' 21-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified EvaluateKey() routine to perform multiple rounds of
'              block encryption when mixing the values for the password key
'              [malngPrimaryKey()] and S-Box [malngSBox()] arrays.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
'              Corrected excess looping in EvaluateKey() routine.
'              Removed obsolete code.  Updated documentation.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within Blowfish_File() routine.
'                4. Split size testing in Blowfish_String() routine.
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
'              - Updated Blowfish_File() routine to reference new property.
'              - Added new property to designate number of rounds to mix
'                primary key array.
' 26-Aug-2013  Kenneth Ives  kenaso@tx.rr.com
'              Enhanced loading of lngWholeWord in EvaluateKey() routine.
' 15-Mar-2017  Kenneth Ives  kenaso@tx.rr.com
'              - Enhanced p-box (primary key) array data based on user
'                selected key length in LoadWorkArrays() routine.
'              - Enhanced s-box (constants) array in LoadWorkArrays() routine.
'              - Added XorBytes() routine to enhance encryption output.
'              - Updated GetWord(), PutWord() routines.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "clsBlowfish"
  Private Const MAX_FILE_SIZE As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MB_5          As Long = &H500000           ' 5242880 bytes
  Private Const BLOCK_SIZE    As Long = 8
  Private Const KEY_ARRAY     As Long = 18
  Private Const CIPHER_ROUNDS As Long = 16
  Private Const MAX_ROUNDS    As Long = 10
  Private Const MAX_BYTE      As Long = 256

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
  Private mobjMath32        As cMath32
  Private mblnCreateNewFile As Boolean
  Private mcurMaxSize       As Currency
  Private mcurAccrued       As Currency
  Private mcurOrigLength    As Currency
  Private mcurCurrByteCnt   As Currency
  Private mlngPtrX          As Long
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private malngSBox()       As Long
  Private malngPrimaryKey() As Long
  Private mabytXBox()       As Byte
  Private mabytPassword()   As Byte


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
    On Error GoTo 0
    Exit Property

Password_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Password_CleanUp

End Property


' ***************************************************************************
' ****                         Methods                                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       Blowfish_File
'
' Description:   Function to encrypt/decrypt a file. 
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objBlowfish.Blowfish_File abytFile(), blnEncrypt             ' Encrypt/Decrypt data
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
' ***************************************************************************
Public Function Blowfish_File(ByRef abytFile() As Byte, _
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

    Const ROUTINE_NAME As String = "Blowfish_File"

    On Error GoTo Blowfish_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Set objBigFiles = New cBigFiles   ' Instantiate class object
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    ResetVariables                    ' Reset class variables

    If Not EvaluateKey() Then
        GoTo Blowfish_File_CleanUp
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
        GoTo Blowfish_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
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
                GoTo Blowfish_File_CleanUp
            End If

            MixAppendedData abytSize()                         ' Decrypt appended data
            CopyMemory lngLastBlock, abytSize(0), 4&           ' Size of last block of data without padding [0-3]
            CopyMemory mcurOrigLength, abytSize(4), 8&         ' Capture original file size [4-11]

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
                GoTo Blowfish_File_CleanUp

            End If
        End If

        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs

        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
        End If

        strTempFile = CreateTempFile()  ' Create a temp target file

        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Blowfish_File_CleanUp
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
            GoTo Blowfish_File_CleanUp
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
            curAmtLeft = 0@                         ' Verify empty variable
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
                GoTo Blowfish_File_CleanUp

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

Blowfish_File_CleanUp:
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
        Blowfish_File = False
    Else
        Blowfish_File = True
    End If

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Blowfish_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Blowfish_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Blowfish_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)          ' convert string data to byte array
'                objBlowfish.Blowfish_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
Public Function Blowfish_String(ByRef abytData() As Byte, _
                       Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte

    Const ROUTINE_NAME As String = "Blowfish_String"

    On Error GoTo Blowfish_String_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True      ' Set flag to stop processing
        Exit Function
    End If

    Erase abytSize()                                    ' Always start with empty arrays
    ResetVariables                                      ' Reset class variables
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds   ' Save file size for progress calcs
    mcurAccrued = 0@                                    ' Init accrued progress

    If Not EvaluateKey() Then
        GoTo Blowfish_String_CleanUp
    End If

    ' encrypt the data
    If blnEncrypt Then

        mcurOrigLength = CCur(UBound(abytData))      ' Get string length for progress display
        CopyMemory abytSize(0), mcurOrigLength, 8&   ' Copy original length to temp array
        MixAppendedData abytSize()                   ' Encrypt appended data

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

Blowfish_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Blowfish_String = False
    Else
        Blowfish_String = True
    End If

    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Blowfish_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Blowfish_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Blowfish_File()
    '           Blowfish_String()

    Dim lngIdx       As Long
    Dim lngHold      As Long
    Dim lngIndex     As Long
    Dim lngKeyLen    As Long
    Dim lngPwdIdx    As Long
    Dim lngWholeWord As Long
    Dim lngLeftWord  As Long
    Dim lngRightWord As Long
    Dim abytPwd()    As Byte

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

    Erase abytPwd()  ' Always start with empty arrays

    LoadWorkArrays                                       ' Preload work arrays
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    mabytXBox() = LoadXBoxArray(abytPwd())               ' Load X-Box array with ASCII 0-255 mixed

    mlngPtrX = 0                    ' Set x-box pointer to zero
    lngPwdIdx = 0                   ' Set temp password array index
    lngHold = mlngRounds            ' 02-Nov-2009 Capture number of rounds for block encryption
    EncryptRounds = mlngKeyRounds   ' 02-Nov-2009 Temp reset number of rounds of block
                                    '             encryption (1-5 rounds is a good mix)
    ' Create key-dependant p-boxes
    For lngIndex = 0 To (KEY_ARRAY - 1)

        lngWholeWord = 0

        ' Create a unique long integer
        For lngIdx = 1 To 4

            ' Overlay last three bytes with first three bytes
            MoveMemory ByVal VarPtr(lngWholeWord) + 1, lngWholeWord, 3&
            lngWholeWord = lngWholeWord Or abytPwd(lngPwdIdx Mod lngKeyLen)
            lngPwdIdx = lngPwdIdx + 1

        Next lngIdx

        ' Update primary key array
        malngPrimaryKey(lngIndex) = malngPrimaryKey(lngIndex) Xor lngWholeWord

    Next lngIndex

    ' Additional mixing of the primary key data
    ' 29-Jan-2010 Removed outside loop.  Updated
    '             number of encrypt rounds instead
    For lngIndex = 0 To (KEY_ARRAY - 2) Step 2

        lngLeftWord = malngPrimaryKey(lngIndex)        ' Capture 2 long integers
        lngRightWord = malngPrimaryKey(lngIndex + 1)
        BlockEncrypt lngLeftWord, lngRightWord         ' Perform block encryption
        malngPrimaryKey(lngIndex) = lngLeftWord        ' Insert data back into array
        malngPrimaryKey(lngIndex + 1) = lngRightWord

    Next lngIndex

    ' Additional mixing of the S-Box data
    ' 29-Jan-2010 Removed outside loop.  Updated
    '             number of encrypt rounds instead
    For lngIndex = 0 To 3
        For lngIdx = 0 To (MAX_BYTE - 2) Step 2

            lngLeftWord = malngSBox(lngIndex, lngIdx)        ' Capture 2 long integers
            lngRightWord = malngSBox(lngIndex, lngIdx + 1)
            BlockEncrypt lngLeftWord, lngRightWord           ' Perform block encryption
            malngSBox(lngIndex, lngIdx) = lngLeftWord        ' Insert data back into array
            malngSBox(lngIndex, lngIdx + 1) = lngRightWord

        Next lngIdx
    Next lngIndex

EvaluateKey_CleanUp:
    Erase abytPwd()           ' Always empty arrays when not needed
    EncryptRounds = lngHold   ' 02-Nov-2009 Reset number of rounds of block
                              '             encryption back to original value
    ' An error occurred or user
    ' opted to STOP processing
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

    ' Called by Blowfish_File()
    '           Blowfish_String()

    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1  ' Get size of incoming array

    ' Data length must be a multiple of 8 bytes
    If (lngDataLength Mod BLOCK_SIZE <> 0) Then
        lngDataLength = (lngDataLength + BLOCK_SIZE) - (lngDataLength Mod BLOCK_SIZE)
        ReDim Preserve abytData(lngDataLength - 1)
    End If

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer incoming data to a temp work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE

        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Encrypt block of data
        BlockEncrypt lngLeftWord, lngRightWord

        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&

        XorBytes abytOutput()   ' Mix with x-box data

        ' Transfer encrypted data back to original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

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

EncryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False   ' An error occurred or user opted to STOP processing
    Else
        EncryptData = True    ' Good Finish
    End If

    On Error GoTo 0
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Blowfish_File()
    '           Blowfish_String()

    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngLeftWord   As Long
    Dim lngRightWord  As Long
    Dim lngDataLength As Long
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1  ' Get size of incoming array

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Transfer incoming data to a temp work array
        CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE

        XorBytes abytInput()   ' Mix with x-box data

        ' Convert bytes to two long integers
        GetWord lngLeftWord, abytInput(), 0&
        GetWord lngRightWord, abytInput(), 4&

        ' Decrypt block of data
        BlockDecrypt lngLeftWord, lngRightWord

        ' Convert two long integers to bytes
        PutWord lngLeftWord, abytOutput(), 0&
        PutWord lngRightWord, abytOutput(), 4&

        ' Transfer decrypted data back to the original array
        CopyMemory abytData(lngOffSet), abytOutput(0), BLOCK_SIZE

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

DecryptData_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False   ' An error occurred or user opted to STOP processing
    Else
        DecryptData = True    ' Good Finish
    End If

    On Error GoTo 0
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

Private Sub BlockEncrypt(ByRef lngLeftWord As Long, _
                         ByRef lngRightWord As Long)

    ' Called by EncryptData()

    Dim lngIdx      As Long
    Dim lngTemp     As Long
    Dim lngIndex    As Long
    Dim lngRound    As Long
    Dim lngTmpLeft  As Long
    Dim lngTmpRight As Long

    lngTmpLeft = lngLeftWord
    lngTmpRight = lngRightWord

    For lngRound = 1 To mlngRounds

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        lngIdx = 0

        For lngIndex = 0 To (CIPHER_ROUNDS \ 2 - 1)

            lngTmpLeft = lngTmpLeft Xor malngPrimaryKey(lngIdx)
            lngTmpRight = CalcNewValue(lngTmpLeft) Xor lngTmpRight

            lngTmpRight = lngTmpRight Xor malngPrimaryKey(lngIdx + 1)
            lngTmpLeft = CalcNewValue(lngTmpRight) Xor lngTmpLeft
            lngIdx = lngIdx + 2

        Next lngIndex

        lngTemp = lngTmpRight
        lngTmpRight = lngTmpLeft Xor malngPrimaryKey(CIPHER_ROUNDS)
        lngTmpLeft = lngTemp Xor malngPrimaryKey(CIPHER_ROUNDS + 1)

    Next lngRound

    lngLeftWord = lngTmpLeft
    lngRightWord = lngTmpRight

End Sub

Private Sub BlockDecrypt(ByRef lngLeftWord As Long, _
                         ByRef lngRightWord As Long)

    ' Called by DecryptData()

    Dim lngIdx      As Long
    Dim lngTemp     As Long
    Dim lngIndex    As Long
    Dim lngRound    As Long
    Dim lngTmpLeft  As Long
    Dim lngTmpRight As Long

    lngTmpLeft = lngLeftWord
    lngTmpRight = lngRightWord

    For lngRound = 1 To mlngRounds

        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        lngTemp = lngTmpRight
        lngTmpRight = lngTmpLeft Xor malngPrimaryKey(CIPHER_ROUNDS + 1)
        lngTmpLeft = lngTemp Xor malngPrimaryKey(CIPHER_ROUNDS)
        lngIdx = CIPHER_ROUNDS - 2

        For lngIndex = 0 To (CIPHER_ROUNDS \ 2 - 1)

            lngTmpLeft = CalcNewValue(lngTmpRight) Xor lngTmpLeft
            lngTmpRight = lngTmpRight Xor malngPrimaryKey(lngIdx + 1)

            lngTmpRight = CalcNewValue(lngTmpLeft) Xor lngTmpRight
            lngTmpLeft = lngTmpLeft Xor malngPrimaryKey(lngIdx)
            lngIdx = lngIdx - 2

        Next lngIndex

    Next lngRound

    lngLeftWord = lngTmpLeft
    lngRightWord = lngTmpRight

End Sub

Private Function CalcNewValue(ByVal lngValue As Long) As Long

    ' Called by BlockEncrypt()
    '           BlockDecrypt()

    Dim lngTemp     As Long
    Dim abytData(4) As Byte

    CopyMemory abytData(0), lngValue, 4&   ' Convert long integer to byte array

    With mobjMath32
        lngTemp = .UnsignedAdd(malngSBox(0, abytData(3)), malngSBox(1, abytData(2)))
        lngTemp = lngTemp Xor malngSBox(2, abytData(1))
        CalcNewValue = .UnsignedAdd(lngTemp, malngSBox(3, abytData(0)))
    End With

End Function

Private Sub XorBytes(ByRef abytData() As Byte)

    ' Called by EncryptData()
    '           DecryptData()
    '
    ' 15-Mar-2017  Added this routine to enhance encryption output

    Dim lngIdx As Long

    For lngIdx = 0 To (BLOCK_SIZE - 1)
        abytData(lngIdx) = abytData(lngIdx) Xor mabytXBox(mlngPtrX)   ' Xor bytes of data
        mlngPtrX = ((mlngPtrX + 1) Mod MAX_BYTE)                      ' Prevent crashing from overflow
    Next lngIdx

End Sub

Private Sub GetWord(ByRef lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)

    ' Called by EncryptData()
    '           DecryptData()

    CopyMemory lngValue, abytData(lngIndex), 4&      ' Create long integer from byte array
    lngValue = mobjMath32.SwapEndianLong(lngValue)   ' Convert to BIG_ENDIAN format

End Sub

Private Sub PutWord(ByVal lngValue As Long, _
                    ByRef abytData() As Byte, _
                    ByVal lngIndex As Long)

    ' Called by EncryptData()
    '           DecryptData()

    lngValue = mobjMath32.SwapEndianLong(lngValue)   ' Convert to LITTLE_ENDIAN format
    CopyMemory abytData(lngIndex), lngValue, 4&      ' Convert long integer to byte array

End Sub

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

    ' Called by Blowfish_File()
    '           Blowfish_String()

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
' Description:   Load primary key and constants arrays.
'
'                Bruce Schneier, author of the Blowfish and Twofish cipher
'                algorithms, has several items on his to do list for
'                Blowfish.  One of which is to make the password key array
'                more dynamic.  Below is my idea for accomplishing this.
'
'                Enhanced the P-Box (primary key) array and S-Box (Constants)
'                array data by making them more dynamic. This was obtained
'                using Microsoft's CryptoAPI Random Number Generator with
'                optional SHA-512 hashed seeding to create cryptograpic values
'                values.
'
'                See application BuildTables.vbp to create additional
'                values.  See Blowfish_Pbox.txt, Blowfish_Sbox.txt file
'                for additional tables.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Mar-2017  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Sub LoadWorkArrays()

    ' Called by EvaluateKey()

    Dim lngIdx    As Long
    Dim lngPos    As Long
    Dim lngIndex  As Long
    Dim lngStart  As Long
    Dim lngFinish As Long
    Dim strData   As String
    Dim avntData  As Variant

    Const SBOX_MAX As Long = 1024&

    ReDim malngPrimaryKey(KEY_ARRAY)    ' Size primary key array
    ReDim malngSBox(0 To 3, MAX_BYTE)   ' Size constant array

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    ' 15-Mar-2017 Kenneth Ives
    ' Enhanced P-box (primary key) data based on user selected key length.
    ' Blowfish_Pbox.txt file for additional P-Box tables and case structures.
    Select Case mlngKeyLength
           Case 32, 256
                strData = strData & "1511A7B4 47B24B34 78A9C1A7 28FBAD5B 59D16BCD 98FE348F "
                strData = strData & "0121388C 7CB22C44 2DEDC19E C7601E79 4E56B61F C13FF5DD "
                strData = strData & "254323A7 942749D5 6C721469 E09D90F0 123CD8DE D4372E8B "

           Case 64, 288
                strData = strData & "2FEECD3A 7E944E6D 392E9324 60D82B1B 611A233F 287E66F5 "
                strData = strData & "C6B6C651 73C4754E 084A89EB CB0ADE0C 55A6FB5B 93DF42E3 "
                strData = strData & "E0C174DC 7EF3175B 2F5F5DFE E8431D9C 7A34660C 305CAF54 "

           Case 96, 320
                strData = strData & "4E85FA4F 99E730D9 7D1C8A01 4E36BC10 ABF5261A 81CCD3E0 "
                strData = strData & "A5EC95D7 C6724345 5E715EA8 7B091BF7 287FEB18 0D808CE3 "
                strData = strData & "CD7605B4 32594FFB 20B5BFFF F6B6C2EE AAA5AB4C 86E0FAD9 "

           Case 128, 352
                strData = strData & "59AF00FF 0AFC5287 0B169A9B 1BF893E3 7DE97E1B D574D159 "
                strData = strData & "0A709585 9D5F100B C9C5DC2D 0161989B C3608850 62BB93A2 "
                strData = strData & "34BFE206 CA937150 421C2810 28C43F4B C852CB61 75D639B5 "

           Case 160, 384
                strData = strData & "E4D75F09 32B986CB 68275173 5B2A30AB 164691F1 D959A0B2 "
                strData = strData & "CE32B6F5 FEF2976D 798A1905 C289738F A6C29FC8 73EA9602 "
                strData = strData & "4C6AA1CC 0565700F C2C86A32 832D1BEC 58AE9130 D629C81B "

           Case 192, 416
                strData = strData & "6C32A8F1 6D449F5D 252B941F 610C382F DE0EBBA3 EA81E31A "
                strData = strData & "8FBC7BCC D9B2EB21 DF08F2C2 48EDCF1F 3F0DAC57 FED82DF1 "
                strData = strData & "AC81AA05 A9765B8B DC15D2ED F11BA2C6 A6AB2294 4F9DD1B3 "

           Case 224, 448
                strData = strData & "5A6A90CF 6EF8EC0B EF76D1AC 30D148B5 CAB92A48 39D67E04 "
                strData = strData & "D62B0E66 F4D7D6E1 A2B89086 D9F213EA 9E1E37E5 717E2C02 "
                strData = strData & "5B4E89CF CC161466 BFFBF6FE 7847525C 1EA31DDB 074B3293 "
    End Select

    avntData = Split(strData, Chr$(32))   ' Separate data into variant array

    ' Transfer data to password key array
    For lngIndex = 0 To (KEY_ARRAY - 1)

        ' Perform Xor bit manipulation
        ' with next value. Last item in
        ' array will be Xor'ed with first
        ' item in variant array.
        malngPrimaryKey(lngIndex) = CLng("&H" & avntData(lngIndex)) Xor _
                                    CLng("&H" & avntData((lngIndex + 1) Mod KEY_ARRAY))
    Next lngIndex

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

    ' 15-Mar-2017 Kenneth Ives
    ' Enhanced S-Box (Constants) array data.  Blowfish_Sbox.txt file
    ' for additional tables.  See application BuildTables.vbp to create
    ' additional values.  Below are 1024 values.
    strData = strData & "2A1FA013 572BFC56 48F97E67 985BBB66 B0515768 6062BA95 CC1D5400 CE71D9E9 "
    strData = strData & "75F77BDD 4390DC2E 1637E1EF 074AEAED 1103ABED 502171D7 8725DC35 E14AE08A "
    strData = strData & "8EB219FB 7282A9D9 1CCC4BF5 3954126F FBFBAD57 AEC74173 F2B817AA B82883A1 "
    strData = strData & "167F7898 9982C4C4 E2860B93 0C5A3823 CBB1C767 5259E9E1 FC5B8587 D628F269 "
    strData = strData & "604217DE 4CA9D5AA 939A2A96 C361AA34 73918F6F 7D170B69 FF4C75A5 E1196288 "
    strData = strData & "C0F55C6B ECC4BF68 646B4341 32F13986 CC4DFD71 8F23B4E0 9A4D29F0 A80F6148 "
    strData = strData & "96D3B31D E02D64A3 0E46F30D CC4AF7C8 6932CF88 A191F182 2E1D155E 80D39C09 "
    strData = strData & "CBDEE4B9 C9C5FB8F 10B28F7F 7594FCAD C9F94695 93B31484 13329757 FF859D03 "
    strData = strData & "80225978 4D32EC84 87598D35 F5AA2804 B359EA91 32419F05 DC0A6B30 9E78E8AC "
    strData = strData & "FBAB3893 EC65B987 7B7F0786 98E47CE3 002D1FB2 4443146F A1E58584 32C10D2D "
    strData = strData & "EA5B384D F9EBB907 9E4E116C E967BF4E E7D4E50E 4193924F 7CA3EACC 70A02420 "
    strData = strData & "27B29A6A 7D800783 E2744911 5B5ACD32 A7614DDA 751305AE 7E3284E1 86930B8F "
    strData = strData & "899C9936 F787F90A C137FCCE 65549170 C9D4A588 B7DAEC38 59611328 C0174FB2 "
    strData = strData & "0CBF0A1A 982D094A 7CCDE318 ED57DF01 DBA58B6C D57D2D6A 6C041E56 C1BFDA90 "
    strData = strData & "C1AE5BE3 9AB75052 C099AE2E 5F1110DE 51AB2184 83E4B4CD 782E4487 A0E3FAA7 "
    strData = strData & "6AF6A22F AE85E402 F90500E7 FEB4C866 1C5BA13D 62BF233B 9DF01EB2 0FB8C8FE "
    strData = strData & "49D00A4D AA4FBFB3 70459590 84D0F427 A64765FC B05E35E7 01956CB3 1DF6928E "
    strData = strData & "E89F0F8B D5FAB81C C500323D C15FAD8C 01E5721B 43A670C8 3568C02A E33C3B1B "
    strData = strData & "F024288C 7C0FEFD6 16E9DB93 EF472168 13D884ED C2AF7279 FE6BE000 E4BE4274 "
    strData = strData & "2D3B04A5 F64597BF 38C0A253 78E9D88A 0D8DA91A 5D753B53 292DD001 8B707DBE "
    strData = strData & "794EF1EA 68E91A1F AE894CC5 FC28E513 BCEE554F C6D114C8 DA5DDC8B 600239EA "
    strData = strData & "1146090B 96146EA4 A65D002A 9C8A5B2B DF6DD52D 019CEEA4 7373D25C 18B60715 "
    strData = strData & "B2A95026 6B44B86A 4AF1E804 5A22B6A3 EEB6F533 E771D405 25672864 C1B7DB06 "
    strData = strData & "90770828 F7B87150 FAB42284 52973A4A FBA819A0 BD084F8A 1C5C37F4 53B8F2C4 "
    strData = strData & "FFCE4148 CD65EDB3 8DB4FFD0 5AE9F51D F91AE7AB 0619109D AE2BF9B9 17A898FE "
    strData = strData & "49B43382 7AB61ADD 2E14698F 3C0ED266 75BEB627 127014CC 7B7DF77A 3FDC99F7 "
    strData = strData & "421D8FA0 EBFF8D1E F5436C5B DCCC8AA0 36CE5489 279C1857 68895A3F D7BD1BE1 "
    strData = strData & "558ACE25 B1043E27 98C31DBF 77A8890B 0DEB6C80 7DC6D1A8 63DA186C CBC753B1 "
    strData = strData & "EF6BE507 B00937E1 0A1E4F4D 702E57A3 565919F3 ECCB9ACD ED4E977B B9C3803A "
    strData = strData & "F260A211 04E03C96 19FA8E16 E6620961 1F1C0845 FF9845EB F5016C97 48D9AE2A "
    strData = strData & "18D5FE54 AEE19DCF DAEE6E5C 6915B727 874F53F7 D5E2FE22 566199FB 3686CF16 "
    strData = strData & "C55554A0 C835342C 8B4B3E88 8109EE51 03F5F0F1 EC5C0E8E E4A716B3 E1081963 "
    strData = strData & "C34339BF 8AB43200 005353F6 627F17AA DA6C460E 04A55B1B 0E0C351D BFFF3673 "
    strData = strData & "DCB23C83 4F4C2386 69F85D3D 50143197 19772FCF 96A244CB D96E7C0E FFABFDF0 "
    strData = strData & "87EEE197 EA5D2EA2 5D200F95 9BB62B2B 3071B77B 20B21B49 2B5624AA C887783D "
    strData = strData & "46784F24 A69D1E43 F22DCECB 9814624C 51E80630 2E926E0B 08AE6DFD FFF0F089 "
    strData = strData & "FD710598 2B2E0217 4902FC87 00793F34 DB75E73D 3CAB81BD DF2F1DB2 80A5CFD4 "
    strData = strData & "E1305BFF EF68D7D9 43E0C0BF 05D52293 DD98E655 A1985771 D18C5AB3 0FE4FFD2 "
    strData = strData & "2288FE36 194C8112 307DA35C 48FD94F4 3BD18631 29F933C0 66157799 7581C61E "
    strData = strData & "ECA418BA 4B22E211 D96CF039 57B0A809 34C4CA14 CC40BCDC BC6878D5 7AC4032D "
    strData = strData & "627E34CB EC45A9F5 C25DEB2D 016663D8 196146EF 41B89337 B9D7D417 B3E12DF8 "
    strData = strData & "AD7DEBB4 797C6D66 08902733 D58D8B35 FD386D12 AAB01714 96B6E177 DD3245C2 "
    strData = strData & "36670922 73ED4051 667E6F27 42ABC108 2592106D BD80BC85 FBFEBEAE EBF63C58 "
    strData = strData & "8B73836E D73A9CAB 0FB57953 08A00EE1 BAEE7EFA ADFF16B1 D49EEE83 342AB46F "
    strData = strData & "E292825B 854FE07A 13C644F9 FEFBCA4B 39EEDCAB 45659176 B90D50CD 8F121FE8 "
    strData = strData & "63F0D9B3 1CA4F31C C0929285 0084B879 9BAE7995 C56FA968 CCE55A25 8667FDED "
    strData = strData & "13AD0865 8B148E00 F6DAB40E 8C08A1DC 1DEC4021 837C00A7 964F0E2F 3085CC51 "
    strData = strData & "2A6A8662 46EEC314 90AB12EB CEA7D179 7179EC82 B3465221 4065E06E 84D51B36 "
    strData = strData & "19C2B947 62EC472F 83031318 9051C12F 6FBE399A 1FB47CC5 FF0A2367 BB311AFC "
    strData = strData & "525C552A 24C41B2B 16352C35 195DB0BE 7435B3DE E2DE190E 3CF1A0BD B9C423C2 "
    strData = strData & "774AA9E4 3786B76D 3ABA2F6F EF8D471E F3FE52A9 99E0E856 101D209D 4185CB10 "
    strData = strData & "81A34389 5E7BF2BA F27155CD FDEECEFE 6C66CDED 4BC63CED 0F47CCF1 71BEF2F0 "
    strData = strData & "6EA2C8F5 B1E57B68 48E679B9 83BE2F39 1CC32C92 05C72807 A7B764B6 07C5D454 "
    strData = strData & "602B8DCB 2FB7A45E 5E3EA21C E301723D 5B8DBFE5 5BAFC4D5 A332118D 5FAE0DEA "
    strData = strData & "6F9D3B20 BDA52C7B 5205E8CF F398805B 1233BC77 3247337A E6B3C367 AE032107 "
    strData = strData & "403347DC 1478CACF CA089B89 5D7C5270 300335D7 12A82027 019BB508 6024DF1B "
    strData = strData & "957A05D7 285EAC58 55EB13B0 7BDDA873 5EF861EA 2F455915 7E83936D DA047E0D "
    strData = strData & "0C101940 5988EF50 6F321FF1 632971E7 EFE64492 0963CF8D EA2CE5E6 213124D1 "
    strData = strData & "CB172774 13575DA3 FA54655E 17F6B394 13D9B996 4B3F9331 0CD1AF3C 58262627 "
    strData = strData & "A016BD74 2D2D9D4B D36B4F60 B38B28DA 62EB152A A3586CE9 F87CCDF9 BC001D6F "
    strData = strData & "0A71DABD 006A1470 3D8D085B 9FD79685 0F2CD8A6 FAD66308 3C57B760 562618D9 "
    strData = strData & "CDB501E8 66894AF8 E3EE216B A262D189 C4F9EB0B 5EF94C4D E17AA344 F82CD27C "
    strData = strData & "55BD287E 3795E214 D50C5A82 6160E6B1 0A9CB6E1 5EC64F17 664E4C5E B67A441B "
    strData = strData & "7C1A8E67 A69345A7 65D36C1E 16FF8EA6 3A1354D1 2035413B 0BF1FC8F B7DD1B24 "
    strData = strData & "8893FFCB F3649565 54FBA3B8 D27C7B79 CF275B07 1542F805 D2B845A8 C5B1E6D0 "
    strData = strData & "EAB5BD73 1DA563B7 8EA7DBCB CCEE1FE8 7F219686 C1C23CF0 BFE17DD8 69333E37 "
    strData = strData & "B346D0D9 863FDC09 23D0FBA4 FABB9178 57098959 B7EDBBDB 83A7AE83 A77BA287 "
    strData = strData & "BF708AD2 75274D43 A40F4C9E 1C763632 9EDBCBE3 AA738B6E C40C9D95 ACB6196D "
    strData = strData & "EEF4F9CE B6BCF1CE DC234D93 5BFBCF9A D10C4410 3F18F5DA 580B2992 825A967A "
    strData = strData & "3224540F CF5BE0DD DC918E12 E5D7F9DC F2EFE44A 6F11A3AA 931A007F DCB13292 "
    strData = strData & "CC721DE0 3D3A5232 08362F6A 025BA921 67A4B3FF 47460D8F 0FA08681 994E1987 "
    strData = strData & "F8CBEF24 DF6D1683 FC48C1BA 5D230D35 AB3925D4 B14A0D64 EEC0D3DC 34B6D4D1 "
    strData = strData & "8F665905 89FC7DE4 2603EF38 B4FF8DA2 AF350686 173066FA F14C5F09 1E0244B8 "
    strData = strData & "1FC60257 B4A4BF31 52F76768 3E507D1E EEC3DEDC 76647C90 FF7CCD28 619A8322 "
    strData = strData & "409577CB DB80BE67 2CD86E8C 02A2EE2B D3545EA6 09412775 C402B7A6 7310A7B1 "
    strData = strData & "2C608B23 5EDED079 FE24612E 0A607CDA 1A3F32F4 8BB2A6F3 E2D12BD3 65822345 "
    strData = strData & "C1E1B0B7 657521C3 A466E033 45C68D6F 1097C2E2 13293086 C1BC9373 4435DFF7 "
    strData = strData & "3A7354FB 7354B99D 115FD31E EE3E9C46 018B98E7 B044E8DD D158C441 CC5135C7 "
    strData = strData & "4A69D7B9 D1A80597 89BDAC53 2F1F50CC CE0B0350 D0DCDB5E 9D4FAFCC BEDBA56F "
    strData = strData & "FBF8BD8E BBA0BC2D DC3181D7 3C8E62DC A7CB295F C63657DD 2280E7DA 001EA6D9 "
    strData = strData & "3B3FA24E E691CCA6 C3215229 6C162A3F 414AF29B 8CC55F0C 671DDBB9 DCC385C4 "
    strData = strData & "25351BDC 73FC71C9 B01342C7 5974FABE 8A8088A8 3EBF98FF D8C24195 06FC7A37 "
    strData = strData & "58F1900B AA9A4D16 5371362E F1C3053F 56ACE83F 3B99D327 5164C980 FD207E15 "
    strData = strData & "69DADE45 F51DB6E2 C9E7ACFA E0F0C298 771B189C A182356F 7F09FABF 28C504AE "
    strData = strData & "401FC56B BE80C1AC AD44115C FA7F6677 DAF0773A 7F6A2120 6CE45588 6C29AB4F "
    strData = strData & "578E3E46 AA105429 E761B483 A895F852 42A3B09F C087092B BEC6E355 DEF96E84 "
    strData = strData & "C9BD8C45 D93CFBA1 372C398F E63D5CF3 043309E5 E191303F CE9FDB04 618A91FA "
    strData = strData & "71266CE7 F9DC445E 70E85228 3EFC4872 63122190 76EB3F0F 6C1C51B5 09BB16C8 "
    strData = strData & "906CE6DC F04895D5 56625A33 B0246B3C EE4BBDD2 3D42A5CB B8E965F8 B5C2B0FA "
    strData = strData & "2C9B061E 0172A318 510E4F6E 7541AD8C 35B128A7 7B345E8A C9D34D99 DDCE3534 "
    strData = strData & "988D8F17 7CF8E51F 682C4BC6 A8B4E8FC A6336297 F36FFDBB 9C927A48 76032DBF "
    strData = strData & "E1D1906E 47094300 6B8C8255 1E12A8FB CD9EC5C3 A3902AAA 79FBED42 FC276BDE "
    strData = strData & "82CA0D47 21BEA7EB EB03037A 4127D219 9E1D9EDE C5F7C33F C895BB7E 67A95F69 "
    strData = strData & "1DBE68BD A4FB7023 7E77994F 5346C4EF E2D16771 0BA67146 5558ECD4 6B29C229 "
    strData = strData & "ECEC78C4 120D5B20 3A7E6D3A 0673D511 07447588 52DFF637 FD47BEAC C6223113 "
    strData = strData & "27469459 97713324 E47A4294 93C054DD 0997B11F 75EB9D71 F443C9C4 33318DF8 "
    strData = strData & "4E33F8A7 5484FF9D 5BFDDC21 EAF8E24F C77005CC D968FFD3 CC19E8FC 2979651D "
    strData = strData & "F3B4A852 36253D38 D35DD76D D5455E47 69A13D08 9F69EAB6 68E87992 B6EFFA69 "
    strData = strData & "FE306610 A9E77275 2C4FF8EC 38B00F32 43A26949 C35CB43F D07BD848 08578401 "
    strData = strData & "3A17B9E0 655711E7 FF036C64 8B1B80E7 2BB6AEEE DBEDA9EF A96A62AB 31A23578 "
    strData = strData & "0C0827B3 4D903B7C 5AF61E73 EEB98D45 16EC4D8C 3F455E7C 7ECB44A2 E68B4232 "
    strData = strData & "EEB13EDD A104CB13 9E24867A D19A408F 5C175F5B 04762547 858E1997 48782B07 "
    strData = strData & "65B7EE9D 134873ED 317BD964 2C916263 E03784B2 024FBEA9 0B817985 FAC20276 "
    strData = strData & "F6FDCF59 26CA18BD 5FCAAA44 B4ED2921 34CD5BB4 15ED7146 608B4D17 DF659CC9 "
    strData = strData & "6462C049 BBDA0EC3 D648ED48 1DE46041 C5747B2C F5FDD03E 4808DC50 1DCFDA46 "
    strData = strData & "BE751070 8ECBD639 E13047CE 3DEEE3C1 8A4226A9 4C07BF31 EF3FE156 2B27CB58 "
    strData = strData & "B8E0A02E 7C772AD3 4959DF60 AD81A8C1 C5568AC5 1BE3DED8 15A8A513 367B999C "
    strData = strData & "202CDA0D 4E9B3466 B73BCE1B 9C26EC52 2A81CA33 8959A530 B3157B51 5EAB0B3E "
    strData = strData & "0AA0486B B7C293BA 1EA6ACBC 9EC9EE33 2D3C0414 D26FBB80 7C1197D4 8A79788A "
    strData = strData & "4F2E7665 CCCAD2F3 808A084D 3871DF40 D6C81D92 8A6FC99F 4F116FA8 5CFB7520 "
    strData = strData & "4E4829E3 759B221B 20EAABC3 667E148F A99C4A3F 4B8C4C51 D382D34E D678CDBB "
    strData = strData & "23887207 1E20975C 7E4AFD66 7A9D3F2A 6B1712A5 874C207F 4922BB57 C732D39D "
    strData = strData & "4B04C884 227858EC CD830464 3417F47C 0DE16C45 00E96C19 14EE8729 582AE95B "
    strData = strData & "0562762E F94B2516 896DF8A3 C3BCE97A 57365A3B A47F1FA8 263A33BC 48ECBC58 "
    strData = strData & "A9601B4C 398E29DF FA28466B 5922848E A8223ECD 4EBCBC08 EAFDAA10 52EC6366 "
    strData = strData & "6064610D 4691D2D5 DD3C054D 3518F4AD 761AE287 B1540002 E6210C81 9FFB1A31 "
    strData = strData & "D3EE6360 45B8CC29 ED7E9F98 2231A91A E524C033 D889B669 FBD41584 9EF9B38E "
    strData = strData & "BF26723D 2544DC90 256ADC75 8B444838 8F6226D8 F05560B2 8B2AA922 18FEF272 "
    strData = strData & "0ADC7E37 EDE5F8F6 942A8200 61D5CF88 C40D3E4F 17CED0A8 64C4017A 937D504F "
    strData = strData & "D22FB26F 066A0DAF A39BDF9C 2D02D8B5 50A21FB1 B6C7821F 257B0A84 CAB7597F "
    strData = strData & "3C7C81DC 72F12D30 5A6E698F 28BC0626 599C3926 2298B9C6 0D2E41C8 D7994190 "
    strData = strData & "04E8449A CC07D69F 0750CC6F 71250CC0 63DD44CA 0A843BF5 ABA692D9 3A325FF5 "
    strData = strData & "991FAF01 4BD8F342 42E2A3F8 5D283140 12B42948 59E80D5D 4B58103C 5C3B939B "
    strData = strData & "FF2571A0 92E99503 CBF6A2EA 6993C723 4A97B401 B4653042 B400F2DF 54B8F15D "
    strData = strData & "2E23DD2C 8BC7A300 382B996C ADDCCE50 2DB3C134 27CCC2B7 F132257A A0930982 "
    strData = strData & "57101B32 19E2EEBD 9C0CF4A5 B4820E95 6BD58BF9 7E3F4E85 9F768A97 ACC1AC16 "
    strData = strData & "794DE6BB 60522ADE 5AEBE2F7 A1352655 6CE329C3 DF3C3353 12F2A045 910D072E "
    strData = strData & "890EFCB7 A385865F F68338B9 2B342CD1 60544F16 A7ABC7DE 30F0C1C1 9CCE0DB5 "

    avntData = Split(strData, Chr$(32))     ' Separate data into variant array
    lngPos = 0                              ' Starting position in S-box array
    lngStart = 0                            ' Starting position in variant array
    lngFinish = lngStart + (MAX_BYTE - 1)   ' Ending position in variant array

    ' Transfer data to constants array
    For lngIdx = 0 To 3

        For lngIndex = lngStart To lngFinish

            ' Perform Xor bit manipulation
            ' with next value.  Last item in
            ' array will be Xor'ed with first
            ' item in variant array.
            malngSBox(lngIdx, lngPos) = CLng("&H" & avntData(lngIndex)) Xor _
                                        CLng("&H" & avntData((lngIndex + 1) Mod SBOX_MAX))
            lngPos = lngPos + 1   ' Increment position in S-box array

        Next lngIndex

        lngPos = 0                              ' Reset S-box position to zero
        lngStart = lngFinish + 1                ' Update starting position in variant array
        lngFinish = lngStart + (MAX_BYTE - 1)   ' Update ending position in variant array

    Next lngIdx

    strData = vbNullString   ' Clear data string
    avntData = Empty         ' Empty variant

End Sub

Private Sub ResetVariables()

    ' Always empty arrays and variants
    ' when not needed.  Saves resources.
    Erase malngPrimaryKey()
    Erase malngSBox()
    Erase mabytXBox()

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

    ' Set property values
    KeyLength = 256        ' Default value
    EncryptRounds = 1      ' Number of rounds of encryption
    PrimaryKeyRounds = 1   ' Number of rounds to mix primary key array
    CreateNewFile = True   ' Default is not to overwrite input file

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    Set mobjMath32 = Nothing
    ResetVariables
    Erase mabytPassword()

End Sub
