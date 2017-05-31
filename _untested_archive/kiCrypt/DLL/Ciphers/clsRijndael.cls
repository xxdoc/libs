VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRijndael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cRijndael  (clsRijndael.cls)
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
' Description:   The block cipher Rijndael was designed by Joan Daemen and
'                Vincent Rijmen as a candidate for the Advanced Encryption
'                Standard. It was chosen by NIST (National Institute of
'                Standards and Technology ) from a field of 15 candidates.
'                The design of Rijndael was strongly influenced by the design
'                of the block cipher Square, which was also created by Joan
'                Daemen and Vincent Rijmen. The name of the algorithm is a
'                combination of the names of its two creators. The Rijndael
'                web page jokes that the name Rijndael was used because we
'                were both fed up with people mutilating the pronunciation of
'                the names 'Daemen' and 'Rijmen'. The algorithm can be
'                implemented very efficiently on a wide range of processors
'                and in hardware.  Like all AES (Advanced Encryption Standard)
'                candidates, Rijndael is very secure and has no known weaknesses.
'
'                Rijndael's key length is defined to be either 128, 192, or
'                256 bits in accordance with the requirements of the AES.
'                Note that unlike Serpent and Twofish, the key size must be
'                one of these values; it is not allowed to be arbitrary. Also
'                unlike other AES candidates, Rijndael has a variable block
'                length of either 128, 192, or 256 bits. All nine combinations
'                of key length and block length are possible, although the
'                official AES block size is 128 bits. Both block length and
'                key length can be extended very easily to multiples of 32
'                bits. The number of rounds, or iterations of the main
'                algorithm, can vary from 10 to 14 and is dependent on the
'                block size and key length.
'
'                Implementation of the AES Rijndael Block Cipher. Inspired
'                by Mike Scott's implementation in C. Permission for free
'                direct or derivative use is granted subject to compliance
'                with any conditions that the originators of the algorithm
'                place on its exploitation.
'
'                Rijndael is an iterated block cipher with a variable block
'                length and a variable key length. The block length and the
'                key length can be independently specified to 128, 192 or
'                256 bits.
'
'                Rijndael is a block cipher, designed by Joan Daemen and
'                Vincent Rijmen as a candidate algorithm for the AES .
'
'                The cipher has a variable block length and key length. We
'                currently specified how to use keys with a length of 128,
'                192, or 256 bits to encrypt blocks with al length of 128,
'                192 or 256 bits (all nine combinations of key length and
'                block length are possible). Both block length and key
'                length can be extended very easily to multiples of 32 bits.
'
'                Rijndael can be implemented very efficiently on a wide
'                range of processors and in hardware.
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length and block size has been expanded
'                  to meet and exceed AES standards.  New sizes are 128,
'                  160, 192, 224, 256 bits.  See properties KeyLength()
'                  and BlockSize().
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
' References:    The block cipher Rijndael (Home page)
'                http://www.ktana.eu/html/theRijndaelPage.htm
'
'                NIST home page for the Rijndael block cipher.  It has links
'                to specifications and source code.
'                http://www.sciencecentral.com/site/497733
'
'                Brian Gladman's Home Page
'                Select cryptography on left panel under technology
'                http://www.gladman.me.uk/
'
'                Tropical software (Portions of description above)
'                http://www.tropsoft.com/strongenc/rijndael.htm
'
'                Rijndael converted to Visual Basic 6.0
'                John Korejwa  <korejwa@tiac.net>
'                09-Jul-2006
'                http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=65906&lngWId=1
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-Jul-2006  John Korejwa  korejwa@tiac.net
'              Rijndael converted to Visual Basic 6.0
'              http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=65906&lngWId=1
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete variables.
'              Updated documentation.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Removed obsolete code.
'              - Updated documentation.
'              - Correct data size capture for decryption in Rijndael_String().
' 18-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote logic selection in GetBlockSize() routine.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 15-Dec-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to perform encryption/decryption multiple times.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while encrypting or
'                decrypting data (EncryptData() and DecryptData() routines).
'              - Removed obsolete code.  Updated documentation.
' 08-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Replaced Do..Loops with For..Next loops in LoadEncKeyArray()
'                routine for consistancy.
'              - Replaced multiple "If" statements with For..Next loop in
'                LoadConstArrays() routine.
'              - Broke out LoadEncKeyArray() routine to multiple routines due
'                to its size.
' 24-Jan-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated EvaluateKey() routine and documented.  Thanks to
'                Peter Stahl for pointing out my flawed process for determining
'                the number of encrypt/decrypt iterations as per Rijndael
'                specifications.
'              - Updated EncryptRounds() property.
'              - Renamed mlngIterations to mlngEncryptRnds.
'              - Updated LoadEncKeyArray(), LoadDecKeyArray() routines.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved "Next lngRound" command down to encompass additional
'                   data in BlockEncrypt_x() and BlockDecrypt_x() routines
'                3. Removed "ReDim abytTemp(mlngDataChunk - 1)" command from
'                   EncryptData() and DecryptData() routines.
'                4. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                5. Moved Decryption size testing to earlier position within
'                   Rijndael_File() routine.
'                6. Split size testing in Rijndael_String() routine.
'              - Moved calculating starting index and incremental step to
'                LoadEncKeyArray() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Moved progress bar code from BlockEncrypt_x() and
'                BlockDecrypt_x() routines to EncryptData() and DecryptData()
'                routines.
'              - Updated progress bar display if using multiple pass
'                functionality for encryption.  See BlockEncrypt_x() and
'                BlockDecrypt_x() routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Rijndael_File() routine to reference new property.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsRijndael"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MAX_ROUNDS     As Long = 10                 ' Modify as you see fit
  Private Const MB_5           As Long = &H500000           ' 5242880 bytes

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
  Private mlngStep          As Long
  Private mlngRounds        As Long
  Private mlngKeyLength     As Long
  Private mlngBlockSize     As Long
  Private mlngDataChunk     As Long
  Private mlngEncryptRnds   As Long
  Private mabytTmp()        As Byte
  Private mabytPassword()   As Byte
  
  ' Used for encryption and decryption
  Private TE0()           As Long
  Private TE1()           As Long
  Private TE2()           As Long
  Private TE3()           As Long
  Private TE4()           As Long
  Private malngEncKey()   As Long
  
  ' Used for decryption only
  Private TD0()           As Long
  Private TD1()           As Long
  Private TD2()           As Long
  Private TD3()           As Long
  Private TD4()           As Long
  Private malngDecKey()   As Long


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

Public Property Let KeyLength(ByVal lngData As Long)

    ' 01-Jan-2009  Meets and exceeds AES standards
    Select Case lngData
           Case 256:  mlngKeyLength = 256
           Case 224:  mlngKeyLength = 224
           Case 192:  mlngKeyLength = 192
           Case 160:  mlngKeyLength = 160
           Case 128:  mlngKeyLength = 128
           Case Else: mlngKeyLength = 128   ' Use default value
    End Select
    
End Property

Public Property Let BlockSize(ByVal lngData As Long)

    ' 01-Jan-2009  Meets and exceeds AES standards
    Select Case lngData
           Case 256:  mlngBlockSize = 256
           Case 224:  mlngBlockSize = 224
           Case 192:  mlngBlockSize = 192
           Case 160:  mlngBlockSize = 160
           Case 128:  mlngBlockSize = 128
           Case Else: mlngBlockSize = 128   ' Use default value
    End Select
    
End Property

Public Property Let EncryptRounds(ByVal lngData As Long)
   
    ' 26-Dec-2008 Added multiple pass functionality.
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngRounds = lngData   ' Good selection
           Case Else:            mlngRounds = 1         ' Default value
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
' Routine:       Rijndael_File (Pronounced rine-doll)
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objRijndael.Rijndael_File abytFile(), blnEncrypt             ' Encrypt/Decrypt data
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
'              Moved Decryption size testing to an earlier position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function Rijndael_File(ByRef abytFile() As Byte, _
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

    Const ROUTINE_NAME As String = "Rijndael_File"
    
    On Error GoTo Rijndael_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles   ' Instantiate class object
    Erase abytSize()                  ' Verify there is no leftover data

    ResetVariables                    ' Reset class variables
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    
    If Not EvaluateKey(blnEncrypt) Then
        GoTo Rijndael_File_CleanUp
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
        GoTo Rijndael_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Rijndael_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            
            mcurOrigLength = curAmtLeft  ' Capture total byte count
        
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
                GoTo Rijndael_File_CleanUp
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
                GoTo Rijndael_File_CleanUp
                
            End If
        End If
    
        mcurMaxSize = mcurOrigLength * mlngRounds  ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Rijndael_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Rijndael_File_CleanUp
        End If
        
        Do
            lngBlockSize = GetBlockSize(curAmtLeft)       ' Process the source file in chunks
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of file left to process
            ReDim abytData(lngBlockSize - 1)              ' Size receiving array
            
            If blnEncrypt Then
                lngLastBlock = lngBlockSize   ' Save size of last block of data
            End If
            
            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do
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
                       
            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)   ' Adjust pointers accordingly
            
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
            GoTo Rijndael_File_CleanUp
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
                GoTo Rijndael_File_CleanUp
                
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

Rijndael_File_CleanUp:
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
        Rijndael_File = False
    Else
        Rijndael_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Rijndael_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Rijndael_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Rijndael_String
'
' Description:   Function to encrypt/decrypt a string of data.
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)          ' convert string data to byte array
'                objRijndael.Rijndael_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
Public Function Rijndael_String(ByRef abytData() As Byte, _
                       Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "Rijndael_String"

    On Error GoTo Rijndael_String_Error
    
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
    
    Erase abytSize()  ' Verify there is no leftover data
    ResetVariables    ' Reset class variables
    
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds  ' Save file size for progress calcs
    mcurAccrued = 0@                                   ' Init accrued progress
    
    If Not EvaluateKey(blnEncrypt) Then
        GoTo Rijndael_String_CleanUp
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
    
Rijndael_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Rijndael_String = False
    Else
        Rijndael_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Verify there is no leftover data
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Rijndael_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Rijndael_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey(ByVal blnEncrypt As Boolean) As Boolean

    ' Called by Rijndael_File()
    '           Rijndael_String()
    
    Dim abytPwd() As Byte

    Const ROUTINE_NAME As String = "EvaluateKey"
    Const THIRTY_TWO   As Long = 32
    
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
            
    LoadConstArrays                      ' Load encryption/decryption constant tables
    mlngDataChunk = (mlngBlockSize \ 8)  ' Calc size of data chunks
    mlngStep = (mlngDataChunk \ 4)       ' Calc starting index and incremental step
    ReDim mabytTmp(mlngDataChunk - 1)
    
    ' Calculate number of Encrypt/Decrypt rounds
    ' (iterations) based on key and block lengths
    '
    '                    Block length in Bit (Nb)
    '                    128  160  192  224  256
    '                   +------------------------+
    '               128 | 10   11   12   13   14 |
    '  Key length   160 | 11   11   12   13   14 |
    '  in Bit (Nk)  192 | 12   12   12   13   14 |
    '               224 | 13   13   13   13   14 |
    '               256 | 14   14   14   14   14 |
    '                   +------------------------+
    '
    Select Case mlngBlockSize
           Case Is > mlngKeyLength: mlngEncryptRnds = ((mlngBlockSize \ 8) \ 4) + 6
           Case Else:               mlngEncryptRnds = ((mlngKeyLength \ 8) \ 4) + 6
    End Select

    'Debug.Print Format$(mlngBlockSize, "B-@@@") & " " & _
                Format$(mlngKeyLength, "@@K-@@@") &  " " & _
                Format$(mlngEncryptRnds, "@@@@")
    
    ' Verify array is empty so
    ' it can accept new data
    Erase abytPwd()
    
    ' Expand password key to maximum length of
    ' 32 elements (32 bytes = 256 bits \ 8 bits).
    ' Password key length will be adjusted
    ' accordingly in LoadEncKeyArray() routine.
    abytPwd() = ExpandData(mabytPassword(), THIRTY_TWO)
    
    ' Load encryption/decryption key arrays
    LoadEncKeyArray abytPwd(), blnEncrypt
            
EvaluateKey_CleanUp:
    ' An error occurred or user opted to STOP processing
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
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Rijndael_File()
    '           Rijndael_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim abytInput()   As Byte
    Dim abytOutput()  As Byte
    Dim abytTemp()    As Byte

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
        
    ' The encoded length may have to be padded out.
    ' This normally does not take place until the
    ' last chunk of data is to be processed.
    If lngDataLength Mod mlngDataChunk <> 0 Then
        lngDataLength = lngDataLength + mlngDataChunk - (lngDataLength Mod mlngDataChunk)
        ReDim Preserve abytData(lngDataLength - 1)
    End If

    ' Size work arrays
    ReDim abytInput(lngDataLength)
    ReDim abytOutput(lngDataLength)
    ReDim abytTemp(mlngDataChunk - 1)
    
    ' Transfer incoming data to temp hold array
    CopyMemory abytInput(0), abytData(0), lngDataLength
    
    ' Encrypt data
    For lngOffSet = 0 To lngDataLength - 1 Step mlngDataChunk

        ' Transfer plain text to temp array
        CopyMemory abytTemp(0), abytInput(lngOffSet), mlngDataChunk
                
        ' Based on block size selected by user,
        ' perform appropriate encryption
        Select Case mlngBlockSize
               Case 128: BlockEncrypt_128 abytTemp()
               Case 160: BlockEncrypt_160 abytTemp()
               Case 192: BlockEncrypt_192 abytTemp()
               Case 224: BlockEncrypt_224 abytTemp()
               Case 256: BlockEncrypt_256 abytTemp()
        End Select
                        
        ' Transfer encrypted data back to original array
        CopyMemory abytOutput(lngOffSet), abytTemp(0), mlngDataChunk
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + CCur(mlngDataChunk * mlngRounds)
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)
            
    Next lngOffSet

    ' Prepare the original array for return.
    ReDim abytData(lngDataLength)

    ' copy the encrypted output array back to the original array
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

    ' Called by Rijndael_File()
    '           Rijndael_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim abytInput()   As Byte
    Dim abytOutput()  As Byte
    Dim abytTemp()    As Byte

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' Work arrays need to be the
    ' same size as the input array
    ReDim abytInput(lngDataLength)
    ReDim abytOutput(lngDataLength)
    ReDim abytTemp(mlngDataChunk - 1)
    
    ' Transfer incoming data to temp hold array
    CopyMemory abytInput(0), abytData(0), lngDataLength

    ' Decrypt data
    For lngOffSet = 0 To lngDataLength - 1 Step mlngDataChunk
    
        ' Transfer encrypted data to temp array
        CopyMemory abytTemp(0), abytInput(lngOffSet), mlngDataChunk
                
        ' Based on block size selected by user,
        ' perform appropriate decryption
        Select Case mlngBlockSize
               Case 128: BlockDecrypt_128 abytTemp()
               Case 160: BlockDecrypt_160 abytTemp()
               Case 192: BlockDecrypt_192 abytTemp()
               Case 224: BlockDecrypt_224 abytTemp()
               Case 256: BlockDecrypt_256 abytTemp()
        End Select
                
        ' Transfer decrypted data to output array
        CopyMemory abytOutput(lngOffSet), abytTemp(0), mlngDataChunk
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If

        ' 19-Feb-2012 Update progress bar. Separate
        '             steps for easier debugging
        mcurAccrued = mcurAccrued + CCur(mlngDataChunk * mlngRounds)
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent CipherProgress(lngProgress)
            
    Next lngOffSet

    ' Prepare the return array
    lngDataLength = UBound(abytOutput)
    ReDim abytData(lngDataLength)

    ' transfer decrypted data to the original array
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

Private Sub BlockEncrypt_128(ByRef abytData() As Byte)
    
    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        
        lngTmp0 = lngTmp0 Xor malngEncKey(0)
        lngTmp1 = lngTmp1 Xor malngEncKey(1)
        lngTmp2 = lngTmp2 Xor malngEncKey(2)
        lngTmp3 = lngTmp3 Xor malngEncKey(3)
        
        lngIdx = mlngStep   ' Set starting index
  
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            
            lngTmp0 = TE0(mabytTmp(0)) Xor _
                      TE1(mabytTmp(5)) Xor _
                      TE2(mabytTmp(10)) Xor _
                      TE3(mabytTmp(15)) Xor _
                      malngEncKey(lngIdx)
    
            lngTmp1 = TE0(mabytTmp(4)) Xor _
                      TE1(mabytTmp(9)) Xor _
                      TE2(mabytTmp(14)) Xor _
                      TE3(mabytTmp(3)) Xor _
                      malngEncKey(lngIdx + 1)
    
            lngTmp2 = TE0(mabytTmp(8)) Xor _
                      TE1(mabytTmp(13)) Xor _
                      TE2(mabytTmp(2)) Xor _
                      TE3(mabytTmp(7)) Xor _
                      malngEncKey(lngIdx + 2)
    
            lngTmp3 = TE0(mabytTmp(12)) Xor _
                      TE1(mabytTmp(1)) Xor _
                      TE2(mabytTmp(6)) Xor _
                      TE3(mabytTmp(11)) Xor _
                      malngEncKey(lngIdx + 3)
    
            lngIdx = lngIdx + mlngStep    ' Increment index
    
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        
        lngTmp0 = (TE4(mabytTmp(0)) And &HFF&) Xor _
                  (TE4(mabytTmp(5)) And &HFF00&) Xor _
                  (TE4(mabytTmp(10)) And &HFF0000) Xor _
                  (TE4(mabytTmp(15)) And &HFF000000) Xor _
                  malngEncKey(lngIdx)
    
        lngTmp1 = (TE4(mabytTmp(4)) And &HFF&) Xor _
                  (TE4(mabytTmp(9)) And &HFF00&) Xor _
                  (TE4(mabytTmp(14)) And &HFF0000) Xor _
                  (TE4(mabytTmp(3)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 1)
    
        lngTmp2 = (TE4(mabytTmp(8)) And &HFF&) Xor _
                  (TE4(mabytTmp(13)) And &HFF00&) Xor _
                  (TE4(mabytTmp(2)) And &HFF0000) Xor _
                  (TE4(mabytTmp(7)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 2)
    
        lngTmp3 = (TE4(mabytTmp(12)) And &HFF&) Xor _
                  (TE4(mabytTmp(1)) And &HFF00&) Xor _
                  (TE4(mabytTmp(6)) And &HFF0000) Xor _
                  (TE4(mabytTmp(11)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 3)
    
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
    
    Next lngRound
        
End Sub

Private Sub BlockDecrypt_128(ByRef abytData() As Byte)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform decryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        
        lngTmp0 = lngTmp0 Xor malngDecKey(0)
        lngTmp1 = lngTmp1 Xor malngDecKey(1)
        lngTmp2 = lngTmp2 Xor malngDecKey(2)
        lngTmp3 = lngTmp3 Xor malngDecKey(3)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            
            lngTmp0 = TD0(mabytTmp(0)) Xor _
                      TD1(mabytTmp(13)) Xor _
                      TD2(mabytTmp(10)) Xor _
                      TD3(mabytTmp(7)) Xor _
                      malngDecKey(lngIdx)
      
            lngTmp1 = TD0(mabytTmp(4)) Xor _
                      TD1(mabytTmp(1)) Xor _
                      TD2(mabytTmp(14)) Xor _
                      TD3(mabytTmp(11)) Xor _
                      malngDecKey(lngIdx + 1)
      
            lngTmp2 = TD0(mabytTmp(8)) Xor _
                      TD1(mabytTmp(5)) Xor _
                      TD2(mabytTmp(2)) Xor _
                      TD3(mabytTmp(15)) Xor _
                      malngDecKey(lngIdx + 2)
      
            lngTmp3 = TD0(mabytTmp(12)) Xor _
                      TD1(mabytTmp(9)) Xor _
                      TD2(mabytTmp(6)) Xor _
                      TD3(mabytTmp(3)) Xor _
                      malngDecKey(lngIdx + 3)
    
            lngIdx = lngIdx + mlngStep    ' Increment index
    
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        
        lngTmp0 = (TD4(mabytTmp(0)) And &HFF&) Xor _
                  (TD4(mabytTmp(13)) And &HFF00&) Xor _
                  (TD4(mabytTmp(10)) And &HFF0000) Xor _
                  (TD4(mabytTmp(7)) And &HFF000000) Xor _
                  malngDecKey(lngIdx)
    
        lngTmp1 = (TD4(mabytTmp(4)) And &HFF&) Xor _
                  (TD4(mabytTmp(1)) And &HFF00&) Xor _
                  (TD4(mabytTmp(14)) And &HFF0000) Xor _
                  (TD4(mabytTmp(11)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 1)
    
        lngTmp2 = (TD4(mabytTmp(8)) And &HFF&) Xor _
                  (TD4(mabytTmp(5)) And &HFF00&) Xor _
                  (TD4(mabytTmp(2)) And &HFF0000) Xor _
                  (TD4(mabytTmp(15)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 2)
    
        lngTmp3 = (TD4(mabytTmp(12)) And &HFF&) Xor _
                  (TD4(mabytTmp(9)) And &HFF00&) Xor _
                  (TD4(mabytTmp(6)) And &HFF0000) Xor _
                  (TD4(mabytTmp(3)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 3)
    
        ' Copy Decrytped back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
    
    Next lngRound
        
End Sub

Private Sub BlockEncrypt_160(ByRef abytData() As Byte)
    
    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        
        lngTmp0 = lngTmp0 Xor malngEncKey(0)
        lngTmp1 = lngTmp1 Xor malngEncKey(1)
        lngTmp2 = lngTmp2 Xor malngEncKey(2)
        lngTmp3 = lngTmp3 Xor malngEncKey(3)
        lngTmp4 = lngTmp4 Xor malngEncKey(4)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            
            lngTmp0 = TE0(mabytTmp(0)) Xor _
                      TE1(mabytTmp(5)) Xor _
                      TE2(mabytTmp(10)) Xor _
                      TE3(mabytTmp(15)) Xor _
                      malngEncKey(lngIdx)
    
            lngTmp1 = TE0(mabytTmp(4)) Xor _
                      TE1(mabytTmp(9)) Xor _
                      TE2(mabytTmp(14)) Xor _
                      TE3(mabytTmp(19)) Xor _
                      malngEncKey(lngIdx + 1)
    
            lngTmp2 = TE0(mabytTmp(8)) Xor _
                      TE1(mabytTmp(13)) Xor _
                      TE2(mabytTmp(18)) Xor _
                      TE3(mabytTmp(3)) Xor _
                      malngEncKey(lngIdx + 2)
    
            lngTmp3 = TE0(mabytTmp(12)) Xor _
                      TE1(mabytTmp(17)) Xor _
                      TE2(mabytTmp(2)) Xor _
                      TE3(mabytTmp(7)) Xor _
                      malngEncKey(lngIdx + 3)
    
            lngTmp4 = TE0(mabytTmp(16)) Xor _
                      TE1(mabytTmp(1)) Xor _
                      TE2(mabytTmp(6)) Xor _
                      TE3(mabytTmp(11)) Xor _
                      malngEncKey(lngIdx + 4)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
    
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        
        lngTmp0 = (TE4(mabytTmp(0)) And &HFF&) Xor _
                  (TE4(mabytTmp(5)) And &HFF00&) Xor _
                  (TE4(mabytTmp(10)) And &HFF0000) Xor _
                  (TE4(mabytTmp(15)) And &HFF000000) Xor _
                  malngEncKey(lngIdx)
    
        lngTmp1 = (TE4(mabytTmp(4)) And &HFF&) Xor _
                  (TE4(mabytTmp(9)) And &HFF00&) Xor _
                  (TE4(mabytTmp(14)) And &HFF0000) Xor _
                  (TE4(mabytTmp(19)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 1)
    
        lngTmp2 = (TE4(mabytTmp(8)) And &HFF&) Xor _
                  (TE4(mabytTmp(13)) And &HFF00&) Xor _
                  (TE4(mabytTmp(18)) And &HFF0000) Xor _
                  (TE4(mabytTmp(3)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 2)
    
        lngTmp3 = (TE4(mabytTmp(12)) And &HFF&) Xor _
                  (TE4(mabytTmp(17)) And &HFF00&) Xor _
                  (TE4(mabytTmp(2)) And &HFF0000) Xor _
                  (TE4(mabytTmp(7)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 3)
    
        lngTmp4 = (TE4(mabytTmp(16)) And &HFF&) Xor _
                  (TE4(mabytTmp(1)) And &HFF00&) Xor _
                  (TE4(mabytTmp(6)) And &HFF0000) Xor _
                  (TE4(mabytTmp(11)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 4)
        
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&

    Next lngRound
        
End Sub

Private Sub BlockDecrypt_160(ByRef abytData() As Byte)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        
        lngTmp0 = lngTmp0 Xor malngDecKey(0)
        lngTmp1 = lngTmp1 Xor malngDecKey(1)
        lngTmp2 = lngTmp2 Xor malngDecKey(2)
        lngTmp3 = lngTmp3 Xor malngDecKey(3)
        lngTmp4 = lngTmp4 Xor malngDecKey(4)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            
            lngTmp0 = TD0(mabytTmp(0)) Xor _
                      TD1(mabytTmp(17)) Xor _
                      TD2(mabytTmp(14)) Xor _
                      TD3(mabytTmp(11)) Xor _
                      malngDecKey(lngIdx)
    
            lngTmp1 = TD0(mabytTmp(4)) Xor _
                      TD1(mabytTmp(1)) Xor _
                      TD2(mabytTmp(18)) Xor _
                      TD3(mabytTmp(15)) Xor _
                      malngDecKey(lngIdx + 1)
    
            lngTmp2 = TD0(mabytTmp(8)) Xor _
                      TD1(mabytTmp(5)) Xor _
                      TD2(mabytTmp(2)) Xor _
                      TD3(mabytTmp(19)) Xor _
                      malngDecKey(lngIdx + 2)
    
            lngTmp3 = TD0(mabytTmp(12)) Xor _
                      TD1(mabytTmp(9)) Xor _
                      TD2(mabytTmp(6)) Xor _
                      TD3(mabytTmp(3)) Xor _
                      malngDecKey(lngIdx + 3)
    
            lngTmp4 = TD0(mabytTmp(16)) Xor _
                      TD1(mabytTmp(13)) Xor _
                      TD2(mabytTmp(10)) Xor _
                      TD3(mabytTmp(7)) Xor _
                      malngDecKey(lngIdx + 4)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
    
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        
        lngTmp0 = (TD4(mabytTmp(0)) And &HFF&) Xor _
                  (TD4(mabytTmp(17)) And &HFF00&) Xor _
                  (TD4(mabytTmp(14)) And &HFF0000) Xor _
                  (TD4(mabytTmp(11)) And &HFF000000) Xor _
                  malngDecKey(lngIdx)
    
        lngTmp1 = (TD4(mabytTmp(4)) And &HFF&) Xor _
                  (TD4(mabytTmp(1)) And &HFF00&) Xor _
                  (TD4(mabytTmp(18)) And &HFF0000) Xor _
                  (TD4(mabytTmp(15)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 1)
    
        lngTmp2 = (TD4(mabytTmp(8)) And &HFF&) Xor _
                  (TD4(mabytTmp(5)) And &HFF00&) Xor _
                  (TD4(mabytTmp(2)) And &HFF0000) Xor _
                  (TD4(mabytTmp(19)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 2)
    
        lngTmp3 = (TD4(mabytTmp(12)) And &HFF&) Xor _
                  (TD4(mabytTmp(9)) And &HFF00&) Xor _
                  (TD4(mabytTmp(6)) And &HFF0000) Xor _
                  (TD4(mabytTmp(3)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 3)
    
        lngTmp4 = (TD4(mabytTmp(16)) And &HFF&) Xor _
                  (TD4(mabytTmp(13)) And &HFF00&) Xor _
                  (TD4(mabytTmp(10)) And &HFF0000) Xor _
                  (TD4(mabytTmp(7)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 4)
    
        ' Copy Decrytped back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&

    Next lngRound
        
End Sub

Private Sub BlockEncrypt_192(ByRef abytData() As Byte)
    
    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        
        lngTmp0 = lngTmp0 Xor malngEncKey(0)
        lngTmp1 = lngTmp1 Xor malngEncKey(1)
        lngTmp2 = lngTmp2 Xor malngEncKey(2)
        lngTmp3 = lngTmp3 Xor malngEncKey(3)
        lngTmp4 = lngTmp4 Xor malngEncKey(4)
        lngTmp5 = lngTmp5 Xor malngEncKey(5)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            
            lngTmp0 = TE0(mabytTmp(0)) Xor _
                      TE1(mabytTmp(5)) Xor _
                      TE2(mabytTmp(10)) Xor _
                      TE3(mabytTmp(15)) Xor _
                      malngEncKey(lngIdx)
            
            lngTmp1 = TE0(mabytTmp(4)) Xor _
                      TE1(mabytTmp(9)) Xor _
                      TE2(mabytTmp(14)) Xor _
                      TE3(mabytTmp(19)) Xor _
                      malngEncKey(lngIdx + 1)
            
            lngTmp2 = TE0(mabytTmp(8)) Xor _
                      TE1(mabytTmp(13)) Xor _
                      TE2(mabytTmp(18)) Xor _
                      TE3(mabytTmp(23)) Xor _
                      malngEncKey(lngIdx + 2)
            
            lngTmp3 = TE0(mabytTmp(12)) Xor _
                      TE1(mabytTmp(17)) Xor _
                      TE2(mabytTmp(22)) Xor _
                      TE3(mabytTmp(3)) Xor _
                      malngEncKey(lngIdx + 3)
            
            lngTmp4 = TE0(mabytTmp(16)) Xor _
                      TE1(mabytTmp(21)) Xor _
                      TE2(mabytTmp(2)) Xor _
                      TE3(mabytTmp(7)) Xor _
                      malngEncKey(lngIdx + 4)
            
            lngTmp5 = TE0(mabytTmp(20)) Xor _
                      TE1(mabytTmp(1)) Xor _
                      TE2(mabytTmp(6)) Xor _
                      TE3(mabytTmp(11)) Xor _
                      malngEncKey(lngIdx + 5)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
        
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        
        lngTmp0 = (TE4(mabytTmp(0)) And &HFF&) Xor _
                  (TE4(mabytTmp(5)) And &HFF00&) Xor _
                  (TE4(mabytTmp(10)) And &HFF0000) Xor _
                  (TE4(mabytTmp(15)) And &HFF000000) Xor _
                  malngEncKey(lngIdx)
        
        lngTmp1 = (TE4(mabytTmp(4)) And &HFF&) Xor _
                  (TE4(mabytTmp(9)) And &HFF00&) Xor _
                  (TE4(mabytTmp(14)) And &HFF0000) Xor _
                  (TE4(mabytTmp(19)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 1)
        
        lngTmp2 = (TE4(mabytTmp(8)) And &HFF&) Xor _
                  (TE4(mabytTmp(13)) And &HFF00&) Xor _
                  (TE4(mabytTmp(18)) And &HFF0000) Xor _
                  (TE4(mabytTmp(23)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 2)
        
        lngTmp3 = (TE4(mabytTmp(12)) And &HFF&) Xor _
                  (TE4(mabytTmp(17)) And &HFF00&) Xor _
                  (TE4(mabytTmp(22)) And &HFF0000) Xor _
                  (TE4(mabytTmp(3)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 3)
        
        lngTmp4 = (TE4(mabytTmp(16)) And &HFF&) Xor _
                  (TE4(mabytTmp(21)) And &HFF00&) Xor _
                  (TE4(mabytTmp(2)) And &HFF0000) Xor _
                  (TE4(mabytTmp(7)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 4)
        
        lngTmp5 = (TE4(mabytTmp(20)) And &HFF&) Xor _
                  (TE4(mabytTmp(1)) And &HFF00&) Xor _
                  (TE4(mabytTmp(6)) And &HFF0000) Xor _
                  (TE4(mabytTmp(11)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 5)
    
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&
    
    Next lngRound
        
End Sub

Private Sub BlockDecrypt_192(ByRef abytData() As Byte)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        
        lngTmp0 = lngTmp0 Xor malngDecKey(0)
        lngTmp1 = lngTmp1 Xor malngDecKey(1)
        lngTmp2 = lngTmp2 Xor malngDecKey(2)
        lngTmp3 = lngTmp3 Xor malngDecKey(3)
        lngTmp4 = lngTmp4 Xor malngDecKey(4)
        lngTmp5 = lngTmp5 Xor malngDecKey(5)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            
            lngTmp0 = TD0(mabytTmp(0)) Xor _
                      TD1(mabytTmp(21)) Xor _
                      TD2(mabytTmp(18)) Xor _
                      TD3(mabytTmp(15)) Xor _
                      malngDecKey(lngIdx)
            
            lngTmp1 = TD0(mabytTmp(4)) Xor _
                      TD1(mabytTmp(1)) Xor _
                      TD2(mabytTmp(22)) Xor _
                      TD3(mabytTmp(19)) Xor _
                      malngDecKey(lngIdx + 1)
            
            lngTmp2 = TD0(mabytTmp(8)) Xor _
                      TD1(mabytTmp(5)) Xor _
                      TD2(mabytTmp(2)) Xor _
                      TD3(mabytTmp(23)) Xor _
                      malngDecKey(lngIdx + 2)
                   
            lngTmp3 = TD0(mabytTmp(12)) Xor _
                      TD1(mabytTmp(9)) Xor _
                      TD2(mabytTmp(6)) Xor _
                      TD3(mabytTmp(3)) Xor _
                      malngDecKey(lngIdx + 3)
            
            lngTmp4 = TD0(mabytTmp(16)) Xor _
                      TD1(mabytTmp(13)) Xor _
                      TD2(mabytTmp(10)) Xor _
                      TD3(mabytTmp(7)) Xor _
                      malngDecKey(lngIdx + 4)
            
            lngTmp5 = TD0(mabytTmp(20)) Xor _
                      TD1(mabytTmp(17)) Xor _
                      TD2(mabytTmp(14)) Xor _
                      TD3(mabytTmp(11)) Xor _
                      malngDecKey(lngIdx + 5)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
            
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        
        lngTmp0 = (TD4(mabytTmp(0)) And &HFF&) Xor _
                  (TD4(mabytTmp(21)) And &HFF00&) Xor _
                  (TD4(mabytTmp(18)) And &HFF0000) Xor _
                  (TD4(mabytTmp(15)) And &HFF000000) Xor _
                  malngDecKey(lngIdx)
        
        lngTmp1 = (TD4(mabytTmp(4)) And &HFF&) Xor _
                  (TD4(mabytTmp(1)) And &HFF00&) Xor _
                  (TD4(mabytTmp(22)) And &HFF0000) Xor _
                  (TD4(mabytTmp(19)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 1)
        
        lngTmp2 = (TD4(mabytTmp(8)) And &HFF&) Xor _
                  (TD4(mabytTmp(5)) And &HFF00&) Xor _
                  (TD4(mabytTmp(2)) And &HFF0000) Xor _
                  (TD4(mabytTmp(23)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 2)
        
        lngTmp3 = (TD4(mabytTmp(12)) And &HFF&) Xor _
                  (TD4(mabytTmp(9)) And &HFF00&) Xor _
                  (TD4(mabytTmp(6)) And &HFF0000) Xor _
                  (TD4(mabytTmp(3)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 3)
        
        lngTmp4 = (TD4(mabytTmp(16)) And &HFF&) Xor _
                  (TD4(mabytTmp(13)) And &HFF00&) Xor _
                  (TD4(mabytTmp(10)) And &HFF0000) Xor _
                  (TD4(mabytTmp(7)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 4)
        
        lngTmp5 = (TD4(mabytTmp(20)) And &HFF&) Xor _
                  (TD4(mabytTmp(17)) And &HFF00&) Xor _
                  (TD4(mabytTmp(14)) And &HFF0000) Xor _
                  (TD4(mabytTmp(11)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 5)
    
        ' Copy decrytped back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&

    Next lngRound
        
End Sub

Private Sub BlockEncrypt_224(ByRef abytData() As Byte)
    
    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngTmp6  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        CopyMemory lngTmp6, abytData(24), 4&
        
        lngTmp0 = lngTmp0 Xor malngEncKey(0)
        lngTmp1 = lngTmp1 Xor malngEncKey(1)
        lngTmp2 = lngTmp2 Xor malngEncKey(2)
        lngTmp3 = lngTmp3 Xor malngEncKey(3)
        lngTmp4 = lngTmp4 Xor malngEncKey(4)
        lngTmp5 = lngTmp5 Xor malngEncKey(5)
        lngTmp6 = lngTmp6 Xor malngEncKey(6)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            CopyMemory mabytTmp(24), lngTmp6, 4&
            
            lngTmp0 = TE0(mabytTmp(0)) Xor _
                      TE1(mabytTmp(5)) Xor _
                      TE2(mabytTmp(10)) Xor _
                      TE3(mabytTmp(19)) Xor _
                      malngEncKey(lngIdx)
            
            lngTmp1 = TE0(mabytTmp(4)) Xor _
                      TE1(mabytTmp(9)) Xor _
                      TE2(mabytTmp(14)) Xor _
                      TE3(mabytTmp(23)) Xor _
                      malngEncKey(lngIdx + 1)
            
            lngTmp2 = TE0(mabytTmp(8)) Xor _
                      TE1(mabytTmp(13)) Xor _
                      TE2(mabytTmp(18)) Xor _
                      TE3(mabytTmp(27)) Xor _
                      malngEncKey(lngIdx + 2)
            
            lngTmp3 = TE0(mabytTmp(12)) Xor _
                      TE1(mabytTmp(17)) Xor _
                      TE2(mabytTmp(22)) Xor _
                      TE3(mabytTmp(3)) Xor _
                      malngEncKey(lngIdx + 3)
            
            lngTmp4 = TE0(mabytTmp(16)) Xor _
                      TE1(mabytTmp(21)) Xor _
                      TE2(mabytTmp(26)) Xor _
                      TE3(mabytTmp(7)) Xor _
                      malngEncKey(lngIdx + 4)
            
            lngTmp5 = TE0(mabytTmp(20)) Xor _
                      TE1(mabytTmp(25)) Xor _
                      TE2(mabytTmp(2)) Xor _
                      TE3(mabytTmp(11)) Xor _
                      malngEncKey(lngIdx + 5)
            
            lngTmp6 = TE0(mabytTmp(24)) Xor _
                      TE1(mabytTmp(1)) Xor _
                      TE2(mabytTmp(6)) Xor _
                      TE3(mabytTmp(15)) Xor _
                      malngEncKey(lngIdx + 6)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
        
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        CopyMemory mabytTmp(24), lngTmp6, 4&
        
        lngTmp0 = (TE4(mabytTmp(0)) And &HFF&) Xor _
                  (TE4(mabytTmp(5)) And &HFF00&) Xor _
                  (TE4(mabytTmp(10)) And &HFF0000) Xor _
                  (TE4(mabytTmp(19)) And &HFF000000) Xor _
                  malngEncKey(lngIdx)
        
        lngTmp1 = (TE4(mabytTmp(4)) And &HFF&) Xor _
                  (TE4(mabytTmp(9)) And &HFF00&) Xor _
                  (TE4(mabytTmp(14)) And &HFF0000) Xor _
                  (TE4(mabytTmp(23)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 1)
        
        lngTmp2 = (TE4(mabytTmp(8)) And &HFF&) Xor _
                  (TE4(mabytTmp(13)) And &HFF00&) Xor _
                  (TE4(mabytTmp(18)) And &HFF0000) Xor _
                  (TE4(mabytTmp(27)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 2)
        
        lngTmp3 = (TE4(mabytTmp(12)) And &HFF&) Xor _
                  (TE4(mabytTmp(17)) And &HFF00&) Xor _
                  (TE4(mabytTmp(22)) And &HFF0000) Xor _
                  (TE4(mabytTmp(3)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 3)
        
        lngTmp4 = (TE4(mabytTmp(16)) And &HFF&) Xor _
                  (TE4(mabytTmp(21)) And &HFF00&) Xor _
                  (TE4(mabytTmp(26)) And &HFF0000) Xor _
                  (TE4(mabytTmp(7)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 4)
        
        lngTmp5 = (TE4(mabytTmp(20)) And &HFF&) Xor _
                  (TE4(mabytTmp(25)) And &HFF00&) Xor _
                  (TE4(mabytTmp(2)) And &HFF0000) Xor _
                  (TE4(mabytTmp(11)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 5)
        
        lngTmp6 = (TE4(mabytTmp(24)) And &HFF&) Xor _
                  (TE4(mabytTmp(1)) And &HFF00&) Xor _
                  (TE4(mabytTmp(6)) And &HFF0000) Xor _
                  (TE4(mabytTmp(15)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 6)
    
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&
        CopyMemory abytData(24), lngTmp6, 4&

    Next lngRound
        
End Sub

Private Sub BlockDecrypt_224(ByRef abytData() As Byte)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngTmp6  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        CopyMemory lngTmp6, abytData(24), 4&
        
        lngTmp0 = lngTmp0 Xor malngDecKey(0)
        lngTmp1 = lngTmp1 Xor malngDecKey(1)
        lngTmp2 = lngTmp2 Xor malngDecKey(2)
        lngTmp3 = lngTmp3 Xor malngDecKey(3)
        lngTmp4 = lngTmp4 Xor malngDecKey(4)
        lngTmp5 = lngTmp5 Xor malngDecKey(5)
        lngTmp6 = lngTmp6 Xor malngDecKey(6)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            CopyMemory mabytTmp(24), lngTmp6, 4&
            
            lngTmp0 = TD0(mabytTmp(0)) Xor _
                      TD1(mabytTmp(25)) Xor _
                      TD2(mabytTmp(22)) Xor _
                      TD3(mabytTmp(15)) Xor _
                      malngDecKey(lngIdx)
            
            lngTmp1 = TD0(mabytTmp(4)) Xor _
                      TD1(mabytTmp(1)) Xor _
                      TD2(mabytTmp(26)) Xor _
                      TD3(mabytTmp(19)) Xor _
                      malngDecKey(lngIdx + 1)
            
            lngTmp2 = TD0(mabytTmp(8)) Xor _
                      TD1(mabytTmp(5)) Xor _
                      TD2(mabytTmp(2)) Xor _
                      TD3(mabytTmp(23)) Xor _
                      malngDecKey(lngIdx + 2)
            
            lngTmp3 = TD0(mabytTmp(12)) Xor _
                      TD1(mabytTmp(9)) Xor _
                      TD2(mabytTmp(6)) Xor _
                      TD3(mabytTmp(27)) Xor _
                      malngDecKey(lngIdx + 3)
            
            lngTmp4 = TD0(mabytTmp(16)) Xor _
                      TD1(mabytTmp(13)) Xor _
                      TD2(mabytTmp(10)) Xor _
                      TD3(mabytTmp(3)) Xor _
                      malngDecKey(lngIdx + 4)
            
            lngTmp5 = TD0(mabytTmp(20)) Xor _
                      TD1(mabytTmp(17)) Xor _
                      TD2(mabytTmp(14)) Xor _
                      TD3(mabytTmp(7)) Xor _
                      malngDecKey(lngIdx + 5)
            
            lngTmp6 = TD0(mabytTmp(24)) Xor _
                      TD1(mabytTmp(21)) Xor _
                      TD2(mabytTmp(18)) Xor _
                      TD3(mabytTmp(11)) Xor _
                      malngDecKey(lngIdx + 6)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
        
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        CopyMemory mabytTmp(24), lngTmp6, 4&
        
        lngTmp0 = (TD4(mabytTmp(0)) And &HFF&) Xor _
                  (TD4(mabytTmp(25)) And &HFF00&) Xor _
                  (TD4(mabytTmp(22)) And &HFF0000) Xor _
                  (TD4(mabytTmp(15)) And &HFF000000) Xor _
                  malngDecKey(lngIdx)
        
        lngTmp1 = (TD4(mabytTmp(4)) And &HFF&) Xor _
                  (TD4(mabytTmp(1)) And &HFF00&) Xor _
                  (TD4(mabytTmp(26)) And &HFF0000) Xor _
                  (TD4(mabytTmp(19)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 1)
        
        lngTmp2 = (TD4(mabytTmp(8)) And &HFF&) Xor _
                  (TD4(mabytTmp(5)) And &HFF00&) Xor _
                  (TD4(mabytTmp(2)) And &HFF0000) Xor _
                  (TD4(mabytTmp(23)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 2)
        
        lngTmp3 = (TD4(mabytTmp(12)) And &HFF&) Xor _
                  (TD4(mabytTmp(9)) And &HFF00&) Xor _
                  (TD4(mabytTmp(6)) And &HFF0000) Xor _
                  (TD4(mabytTmp(27)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 3)
        
        lngTmp4 = (TD4(mabytTmp(16)) And &HFF&) Xor _
                  (TD4(mabytTmp(13)) And &HFF00&) Xor _
                  (TD4(mabytTmp(10)) And &HFF0000) Xor _
                  (TD4(mabytTmp(3)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 4)
        
        lngTmp5 = (TD4(mabytTmp(20)) And &HFF&) Xor _
                  (TD4(mabytTmp(17)) And &HFF00&) Xor _
                  (TD4(mabytTmp(14)) And &HFF0000) Xor _
                  (TD4(mabytTmp(7)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 5)
        
        lngTmp6 = (TD4(mabytTmp(24)) And &HFF&) Xor _
                  (TD4(mabytTmp(21)) And &HFF00&) Xor _
                  (TD4(mabytTmp(18)) And &HFF0000) Xor _
                  (TD4(mabytTmp(11)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 6)
    
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&
        CopyMemory abytData(24), lngTmp6, 4&

    Next lngRound
        
End Sub

Private Sub BlockEncrypt_256(ByRef abytData() As Byte)
    
    ' Called by EncryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngTmp6  As Long
    Dim lngTmp7  As Long
    Dim lngRound As Long
        
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        CopyMemory lngTmp6, abytData(24), 4&
        CopyMemory lngTmp7, abytData(28), 4&
        
        lngTmp0 = lngTmp0 Xor malngEncKey(0)
        lngTmp1 = lngTmp1 Xor malngEncKey(1)
        lngTmp2 = lngTmp2 Xor malngEncKey(2)
        lngTmp3 = lngTmp3 Xor malngEncKey(3)
        lngTmp4 = lngTmp4 Xor malngEncKey(4)
        lngTmp5 = lngTmp5 Xor malngEncKey(5)
        lngTmp6 = lngTmp6 Xor malngEncKey(6)
        lngTmp7 = lngTmp7 Xor malngEncKey(7)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            CopyMemory mabytTmp(24), lngTmp6, 4&
            CopyMemory mabytTmp(28), lngTmp7, 4&
            
            lngTmp0 = TE0(mabytTmp(0)) Xor _
                      TE1(mabytTmp(5)) Xor _
                      TE2(mabytTmp(14)) Xor _
                      TE3(mabytTmp(19)) Xor _
                      malngEncKey(lngIdx)
            
            lngTmp1 = TE0(mabytTmp(4)) Xor _
                      TE1(mabytTmp(9)) Xor _
                      TE2(mabytTmp(18)) Xor _
                      TE3(mabytTmp(23)) Xor _
                      malngEncKey(lngIdx + 1)
            
            lngTmp2 = TE0(mabytTmp(8)) Xor _
                      TE1(mabytTmp(13)) Xor _
                      TE2(mabytTmp(22)) Xor _
                      TE3(mabytTmp(27)) Xor _
                      malngEncKey(lngIdx + 2)
            
            lngTmp3 = TE0(mabytTmp(12)) Xor _
                      TE1(mabytTmp(17)) Xor _
                      TE2(mabytTmp(26)) Xor _
                      TE3(mabytTmp(31)) Xor _
                      malngEncKey(lngIdx + 3)
            
            lngTmp4 = TE0(mabytTmp(16)) Xor _
                      TE1(mabytTmp(21)) Xor _
                      TE2(mabytTmp(30)) Xor _
                      TE3(mabytTmp(3)) Xor _
                      malngEncKey(lngIdx + 4)
            
            lngTmp5 = TE0(mabytTmp(20)) Xor _
                      TE1(mabytTmp(25)) Xor _
                      TE2(mabytTmp(2)) Xor _
                      TE3(mabytTmp(7)) Xor _
                      malngEncKey(lngIdx + 5)
            
            lngTmp6 = TE0(mabytTmp(24)) Xor _
                      TE1(mabytTmp(29)) Xor _
                      TE2(mabytTmp(6)) Xor _
                      TE3(mabytTmp(11)) Xor _
                      malngEncKey(lngIdx + 6)
            
            lngTmp7 = TE0(mabytTmp(28)) Xor _
                      TE1(mabytTmp(1)) Xor _
                      TE2(mabytTmp(10)) Xor _
                      TE3(mabytTmp(15)) Xor _
                      malngEncKey(lngIdx + 7)
            
            lngIdx = lngIdx + mlngStep    ' Increment index
        
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        CopyMemory mabytTmp(24), lngTmp6, 4&
        CopyMemory mabytTmp(28), lngTmp7, 4&
        
        lngTmp0 = (TE4(mabytTmp(0)) And &HFF&) Xor _
                  (TE4(mabytTmp(5)) And &HFF00&) Xor _
                  (TE4(mabytTmp(14)) And &HFF0000) Xor _
                  (TE4(mabytTmp(19)) And &HFF000000) Xor _
                  malngEncKey(lngIdx)
        
        lngTmp1 = (TE4(mabytTmp(4)) And &HFF&) Xor _
                  (TE4(mabytTmp(9)) And &HFF00&) Xor _
                  (TE4(mabytTmp(18)) And &HFF0000) Xor _
                  (TE4(mabytTmp(23)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 1)
        
        lngTmp2 = (TE4(mabytTmp(8)) And &HFF&) Xor _
                  (TE4(mabytTmp(13)) And &HFF00&) Xor _
                  (TE4(mabytTmp(22)) And &HFF0000) Xor _
                  (TE4(mabytTmp(27)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 2)
        
        lngTmp3 = (TE4(mabytTmp(12)) And &HFF&) Xor _
                  (TE4(mabytTmp(17)) And &HFF00&) Xor _
                  (TE4(mabytTmp(26)) And &HFF0000) Xor _
                  (TE4(mabytTmp(31)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 3)
        
        lngTmp4 = (TE4(mabytTmp(16)) And &HFF&) Xor _
                  (TE4(mabytTmp(21)) And &HFF00&) Xor _
                  (TE4(mabytTmp(30)) And &HFF0000) Xor _
                  (TE4(mabytTmp(3)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 4)
        
        lngTmp5 = (TE4(mabytTmp(20)) And &HFF&) Xor _
                  (TE4(mabytTmp(25)) And &HFF00&) Xor _
                  (TE4(mabytTmp(2)) And &HFF0000) Xor _
                  (TE4(mabytTmp(7)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 5)
        
        lngTmp6 = (TE4(mabytTmp(24)) And &HFF&) Xor _
                  (TE4(mabytTmp(29)) And &HFF00&) Xor _
                  (TE4(mabytTmp(6)) And &HFF0000) Xor _
                  (TE4(mabytTmp(11)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 6)
        
        lngTmp7 = (TE4(mabytTmp(28)) And &HFF&) Xor _
                  (TE4(mabytTmp(1)) And &HFF00&) Xor _
                  (TE4(mabytTmp(10)) And &HFF0000) Xor _
                  (TE4(mabytTmp(15)) And &HFF000000) Xor _
                  malngEncKey(lngIdx + 7)
    
        ' Copy encrypted back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&
        CopyMemory abytData(24), lngTmp6, 4&
        CopyMemory abytData(28), lngTmp7, 4&

    Next lngRound
        
End Sub

Private Sub BlockDecrypt_256(ByRef abytData() As Byte)

    ' Called by DecryptData()
    
    Dim lngIdx   As Long
    Dim lngLoop  As Long
    Dim lngTmp0  As Long
    Dim lngTmp1  As Long
    Dim lngTmp2  As Long
    Dim lngTmp3  As Long
    Dim lngTmp4  As Long
    Dim lngTmp5  As Long
    Dim lngTmp6  As Long
    Dim lngTmp7  As Long
    Dim lngRound As Long
    
    ' 15-Dec-2009  Perform encryption for
    '              required number of rounds
    For lngRound = 1 To mlngRounds
        
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
        ' Copy incoming data into long integers
        CopyMemory lngTmp0, abytData(0), 4&
        CopyMemory lngTmp1, abytData(4), 4&
        CopyMemory lngTmp2, abytData(8), 4&
        CopyMemory lngTmp3, abytData(12), 4&
        CopyMemory lngTmp4, abytData(16), 4&
        CopyMemory lngTmp5, abytData(20), 4&
        CopyMemory lngTmp6, abytData(24), 4&
        CopyMemory lngTmp7, abytData(28), 4&
        
        lngTmp0 = lngTmp0 Xor malngDecKey(0)
        lngTmp1 = lngTmp1 Xor malngDecKey(1)
        lngTmp2 = lngTmp2 Xor malngDecKey(2)
        lngTmp3 = lngTmp3 Xor malngDecKey(3)
        lngTmp4 = lngTmp4 Xor malngDecKey(4)
        lngTmp5 = lngTmp5 Xor malngDecKey(5)
        lngTmp6 = lngTmp6 Xor malngDecKey(6)
        lngTmp7 = lngTmp7 Xor malngDecKey(7)
        
        lngIdx = mlngStep   ' Set starting index
        
        For lngLoop = 1 To (mlngEncryptRnds - 1)
            
            CopyMemory mabytTmp(0), lngTmp0, 4&
            CopyMemory mabytTmp(4), lngTmp1, 4&
            CopyMemory mabytTmp(8), lngTmp2, 4&
            CopyMemory mabytTmp(12), lngTmp3, 4&
            CopyMemory mabytTmp(16), lngTmp4, 4&
            CopyMemory mabytTmp(20), lngTmp5, 4&
            CopyMemory mabytTmp(24), lngTmp6, 4&
            CopyMemory mabytTmp(28), lngTmp7, 4&
            
            lngTmp0 = TD0(mabytTmp(0)) Xor _
                      TD1(mabytTmp(29)) Xor _
                      TD2(mabytTmp(22)) Xor _
                      TD3(mabytTmp(19)) Xor _
                      malngDecKey(lngIdx)
     
            lngTmp1 = TD0(mabytTmp(4)) Xor _
                      TD1(mabytTmp(1)) Xor _
                      TD2(mabytTmp(26)) Xor _
                      TD3(mabytTmp(23)) Xor _
                      malngDecKey(lngIdx + 1)
     
            lngTmp2 = TD0(mabytTmp(8)) Xor _
                      TD1(mabytTmp(5)) Xor _
                      TD2(mabytTmp(30)) Xor _
                      TD3(mabytTmp(27)) Xor _
                      malngDecKey(lngIdx + 2)
     
            lngTmp3 = TD0(mabytTmp(12)) Xor _
                      TD1(mabytTmp(9)) Xor _
                      TD2(mabytTmp(2)) Xor _
                      TD3(mabytTmp(31)) Xor _
                      malngDecKey(lngIdx + 3)
     
            lngTmp4 = TD0(mabytTmp(16)) Xor _
                      TD1(mabytTmp(13)) Xor _
                      TD2(mabytTmp(6)) Xor _
                      TD3(mabytTmp(3)) Xor _
                      malngDecKey(lngIdx + 4)
     
            lngTmp5 = TD0(mabytTmp(20)) Xor _
                      TD1(mabytTmp(17)) Xor _
                      TD2(mabytTmp(10)) Xor _
                      TD3(mabytTmp(7)) Xor _
                      malngDecKey(lngIdx + 5)
     
            lngTmp6 = TD0(mabytTmp(24)) Xor _
                      TD1(mabytTmp(21)) Xor _
                      TD2(mabytTmp(14)) Xor _
                      TD3(mabytTmp(11)) Xor _
                      malngDecKey(lngIdx + 6)
     
            lngTmp7 = TD0(mabytTmp(28)) Xor _
                      TD1(mabytTmp(25)) Xor _
                      TD2(mabytTmp(18)) Xor _
                      TD3(mabytTmp(15)) Xor _
                      malngDecKey(lngIdx + 7)
        
            lngIdx = lngIdx + mlngStep    ' Increment index
        
        Next lngLoop
    
        ' Final round
        CopyMemory mabytTmp(0), lngTmp0, 4&
        CopyMemory mabytTmp(4), lngTmp1, 4&
        CopyMemory mabytTmp(8), lngTmp2, 4&
        CopyMemory mabytTmp(12), lngTmp3, 4&
        CopyMemory mabytTmp(16), lngTmp4, 4&
        CopyMemory mabytTmp(20), lngTmp5, 4&
        CopyMemory mabytTmp(24), lngTmp6, 4&
        CopyMemory mabytTmp(28), lngTmp7, 4&
            
        lngTmp0 = (TD4(mabytTmp(0)) And &HFF&) Xor _
                  (TD4(mabytTmp(29)) And &HFF00&) Xor _
                  (TD4(mabytTmp(22)) And &HFF0000) Xor _
                  (TD4(mabytTmp(19)) And &HFF000000) Xor _
                  malngDecKey(lngIdx)
    
        lngTmp1 = (TD4(mabytTmp(4)) And &HFF&) Xor _
                  (TD4(mabytTmp(1)) And &HFF00&) Xor _
                  (TD4(mabytTmp(26)) And &HFF0000) Xor _
                  (TD4(mabytTmp(23)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 1)
    
        lngTmp2 = (TD4(mabytTmp(8)) And &HFF&) Xor _
                  (TD4(mabytTmp(5)) And &HFF00&) Xor _
                  (TD4(mabytTmp(30)) And &HFF0000) Xor _
                  (TD4(mabytTmp(27)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 2)
    
        lngTmp3 = (TD4(mabytTmp(12)) And &HFF&) Xor _
                  (TD4(mabytTmp(9)) And &HFF00&) Xor _
                  (TD4(mabytTmp(2)) And &HFF0000) Xor _
                  (TD4(mabytTmp(31)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 3)
    
        lngTmp4 = (TD4(mabytTmp(16)) And &HFF&) Xor _
                  (TD4(mabytTmp(13)) And &HFF00&) Xor _
                  (TD4(mabytTmp(6)) And &HFF0000) Xor _
                  (TD4(mabytTmp(3)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 4)
    
        lngTmp5 = (TD4(mabytTmp(20)) And &HFF&) Xor _
                  (TD4(mabytTmp(17)) And &HFF00&) Xor _
                  (TD4(mabytTmp(10)) And &HFF0000) Xor _
                  (TD4(mabytTmp(7)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 5)
    
        lngTmp6 = (TD4(mabytTmp(24)) And &HFF&) Xor _
                  (TD4(mabytTmp(21)) And &HFF00&) Xor _
                  (TD4(mabytTmp(14)) And &HFF0000) Xor _
                  (TD4(mabytTmp(11)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 6)
    
        lngTmp7 = (TD4(mabytTmp(28)) And &HFF&) Xor _
                  (TD4(mabytTmp(25)) And &HFF00&) Xor _
                  (TD4(mabytTmp(18)) And &HFF0000) Xor _
                  (TD4(mabytTmp(15)) And &HFF000000) Xor _
                  malngDecKey(lngIdx + 7)
            
        ' Copy decrytped back to input array
        CopyMemory abytData(0), lngTmp0, 4&
        CopyMemory abytData(4), lngTmp1, 4&
        CopyMemory abytData(8), lngTmp2, 4&
        CopyMemory abytData(12), lngTmp3, 4&
        CopyMemory abytData(16), lngTmp4, 4&
        CopyMemory abytData(20), lngTmp5, 4&
        CopyMemory abytData(24), lngTmp6, 4&
        CopyMemory abytData(28), lngTmp7, 4&

    Next lngRound
        
End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the record to be written.  The
'                write process has been speeded up by 50% or more by
'                adjusting the record length based on amount of data left
'                to write.
'
' Parameters:    curAmtLeft - Amount of data left to be written
'
' Returns:       New record size as a long integer
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 18-Jun-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote logical selections
' ***************************************************************************
Private Function GetBlockSize(ByVal curAmtLeft As Currency) As Long

    ' Called by Rijndael_File()
    '           Rijndael_String()
    
    ' Determine record size to process.
    Select Case mlngBlockSize
           
           Case 128, 256
                Select Case curAmtLeft
                       Case Is >= MB_5: GetBlockSize = MB_5
                       Case Else:       GetBlockSize = CLng(curAmtLeft)
                End Select
                
           Case 160, 192, 224
                Select Case curAmtLeft
                       Case Is >= &HFFF0&: GetBlockSize = &HFFF0&   ' 65520
                       Case Else:          GetBlockSize = CLng(curAmtLeft)
                End Select
    End Select
    
    RaiseEvent CurrentBlockSize(GetBlockSize)
    
End Function

Private Sub LoadEncKeyArray(ByRef abytPwd() As Byte, _
                            ByVal blnEncrypt As Boolean)
    
    ' Called by EvaluateKey()
    
    Dim lngStep As Long
    
    Const ROUTINE_NAME As String = "LoadEncKeyArray"
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    Erase malngEncKey()                  ' Always start with empty arrays
    lngStep = ((mlngKeyLength \ 8) \ 4)  ' Calc starting index and incremental step
    
    ' If performing encryption then save memory
    ' resources by emptying these arrays since
    ' they are no longer required.
    If blnEncrypt Then
        Erase TD0()
        Erase TD1()
        Erase TD2()
        Erase TD3()
        Erase TD4()
        Erase malngDecKey()
    End If
    
    ReDim malngEncKey(0 To 119)   ' Size encryption key array
    
    Select Case mlngBlockSize
           Case 128: LoadEncKey_128 abytPwd(), lngStep   ' 128 bit block size (16 bytes)
           Case 160: LoadEncKey_160 abytPwd(), lngStep   ' 160 bit block size (20 bytes)
           Case 192: LoadEncKey_192 abytPwd(), lngStep   ' 192 bit block size (24 bytes)
           Case 224: LoadEncKey_224 abytPwd(), lngStep   ' 224 bit block size (28 bytes)
           Case 256: LoadEncKey_256 abytPwd(), lngStep   ' 256 bit block size (32 bytes)
           Case Else
                InfoMsg "Invalid block size  (Block size:  " & _
                        CStr(mlngBlockSize) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    ' Perform decryption
    If Not blnEncrypt Then
        
        Erase TE0()   ' Always empty arrays when not needed
        Erase TE1()
        Erase TE2()
        Erase TE3()
    
        ' Load decryption key array
        ' 24-Jan-2011 Removed passed parameter
        LoadDecKeyArray
    
        Erase TE4()          ' Always empty arrays when not needed
        Erase malngEncKey()
    
    End If
    
End Sub

Private Sub LoadEncKey_128(ByRef abytPwd() As Byte, _
                           ByVal lngStep As Long)
                              
    ' Called by LoadEncKeyArray()
    
    Dim lngCnt     As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    Dim abytTmp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadEncKey_128"

    On Error GoTo LoadEncKey_128_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadEncKey_128_CleanUp
    End If

    Erase alngRcon()              ' Always start with an empty array
    alngRcon() = LoadRconArray()  ' Load Rcon array
    lngIndex = lngStep
                
    ' Use just the amount of key
    ' length requested by user
    Select Case mlngKeyLength
           
           Case 128  ' 128 bit key Length (16 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 9
                
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 2)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
           
           Case 160  ' 160 bit key length (20 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 8
                
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 1)
                    
                    If lngCnt = 8 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 3)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case 192  ' 192 bit key length (24 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 7
                
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 2)
                    
                    If lngCnt = 7 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 3)
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 4)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case 224  ' 224 bit key length (28 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 6
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 3) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                                              
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 5)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case 256  ' 256 bit key length (32 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 6
                
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 7) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 2)
                    
                    If lngCnt = 6 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 4) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                                              
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 5)
                    malngEncKey(lngIndex + 7) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 6)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case Else
                InfoMsg "Invalid key length  (Key length:  " & _
                        CStr(mlngKeyLength) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
    
LoadEncKey_128_CleanUp:
    Erase alngRcon()  ' Always empty array when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

LoadEncKey_128_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadEncKey_128_CleanUp

End Sub

Private Sub LoadEncKey_160(ByRef abytPwd() As Byte, _
                           ByVal lngStep As Long)
                              
    ' Called by LoadEncKeyArray()
    
    Dim lngCnt     As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    Dim abytTmp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadEncKey_160"

    On Error GoTo LoadEncKey_160_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadEncKey_160_CleanUp
    End If

    Erase alngRcon()              ' Always start with an empty array
    alngRcon() = LoadRconArray()  ' Load Rcon array
    lngIndex = lngStep
                
    ' Use just the amount of key
    ' length requested by user
    Select Case mlngKeyLength
           
           Case 128  ' 128 bit key length (16 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 13
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 2)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 160  ' 160 bit key length (20 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 10
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 3)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 192  ' 192 bit key length (24 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 9
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 3)
                    
                    If lngCnt = 9 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 4)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 224  ' 224 bit key length (28 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 8
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 2)
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 3) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 5)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 256  ' 256 bit key length (32 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 8
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 7) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex + 1)
                    
                    If lngCnt = 8 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 4) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                                              
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 5)
                    malngEncKey(lngIndex + 7) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 6)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case Else
                InfoMsg "Invalid key length  (Key length:  " & _
                        CStr(mlngKeyLength) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
            
LoadEncKey_160_CleanUp:
    Erase alngRcon()  ' Always empty array when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

LoadEncKey_160_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadEncKey_160_CleanUp

End Sub

Private Sub LoadEncKey_192(ByRef abytPwd() As Byte, _
                           ByVal lngStep As Long)
                              
    ' Called by LoadEncKeyArray()
    
    Dim lngCnt     As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    Dim abytTmp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadEncKey_192"

    On Error GoTo LoadEncKey_192_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadEncKey_192_CleanUp
    End If

    Erase alngRcon()              ' Always start with an empty array
    alngRcon() = LoadRconArray()  ' Load Rcon array
    lngIndex = lngStep
                
    ' Use just the amount of key
    ' length requested by user
    Select Case mlngKeyLength
           
           Case 128  ' 128 bit key length (16 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 18
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex)
                    
                    If lngCnt = 18 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 2)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 160  ' 160 bit key length (20 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 14
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 1)
                    
                    If lngCnt = 14 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 3)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 192  ' 192 bit key length (24 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 11
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 3)
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 4)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case 224  ' 224 bit key length (28 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 10
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 3) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 5)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 256  ' 256 bit key length (32 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 10
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                                          
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 7) Xor malngEncKey(lngIndex)
                    
                    If lngCnt = 10 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 4) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 5)
                    malngEncKey(lngIndex + 7) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 6)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case Else
                InfoMsg "Invalid key length  (Key length:  " & _
                        CStr(mlngKeyLength) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
    
LoadEncKey_192_CleanUp:
    Erase alngRcon()  ' Always empty array when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

LoadEncKey_192_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadEncKey_192_CleanUp

End Sub

Private Sub LoadEncKey_224(ByRef abytPwd() As Byte, _
                           ByVal lngStep As Long)
                              
    ' Called by LoadEncKeyArray()
    
    Dim lngCnt     As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    Dim abytTmp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadEncKey_224"

    On Error GoTo LoadEncKey_224_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadEncKey_224_CleanUp
    End If

    Erase alngRcon()              ' Always start with an empty array
    alngRcon() = LoadRconArray()  ' Load Rcon array
    lngIndex = lngStep
                
    ' Use just the amount of key
    ' length requested by user
    Select Case mlngKeyLength
           
           Case 128  ' 128 bit key length (16 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 23
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex)
                    
                    If lngCnt = 23 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 2)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 160  ' 160 bit key length (20 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 18
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 1)
                    
                    If lngCnt = 18 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 3)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 192  ' 192 bit key length (24 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 15
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex)
                    
                    If lngCnt = 15 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 3)
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 4)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 224  ' 224 bit key length (28 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 12
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 3) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 5)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 256  ' 256 bit key length (32 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 12
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    If lngCnt = 12 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 7) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 4) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                    
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 5)
                    malngEncKey(lngIndex + 7) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 6)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case Else
                InfoMsg "Invalid key length  (Key length:  " & _
                        CStr(mlngKeyLength) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
                
LoadEncKey_224_CleanUp:
    Erase alngRcon()  ' Always empty array when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

LoadEncKey_224_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadEncKey_224_CleanUp

End Sub

Private Sub LoadEncKey_256(ByRef abytPwd() As Byte, _
                           ByVal lngStep As Long)
                              
    ' Called by LoadEncKeyArray()
    
    Dim lngCnt     As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    Dim abytTmp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadEncKey_256"

    On Error GoTo LoadEncKey_256_Error

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadEncKey_256_CleanUp
    End If

    Erase alngRcon()              ' Always start with an empty array
    alngRcon() = LoadRconArray()  ' Load Rcon array
    lngIndex = lngStep
    
    ' Use just the amount of key
    ' length requested by user
    Select Case mlngKeyLength
           
           Case 128  ' 128 bit key length (16 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 28
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 2)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
           
           Case 160  ' 160 bit key length (20 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 22
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 3)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 192  ' 192 bit key length (24 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 18
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 2)
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 3)
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 4)
                    lngIndex = lngIndex + lngStep    ' Increment index
                
                Next lngCnt
                
           Case 224  ' 224 bit key length (28 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 16
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    If lngCnt = 16 Then
                        Exit For  ' exit For..Next loop
                    End If
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 4) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 3) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                                              
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 5)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case 256  ' 256 bit key length (32 bytes)
                CopyMemory malngEncKey(0), abytPwd(0), 4& * lngIndex
                
                For lngCnt = 0 To 13
                    
                    CopyMemory abytTmp(0), malngEncKey(lngIndex - 1), 4&
                    
                    malngEncKey(lngIndex) = malngEncKey(lngIndex - lngStep) Xor _
                                            (TE4(abytTmp(0)) And &HFF000000) Xor _
                                            (TE4(abytTmp(3)) And &HFF0000) Xor _
                                            (TE4(abytTmp(2)) And &HFF00&) Xor _
                                            (TE4(abytTmp(1)) And &HFF&) Xor _
                                            alngRcon(lngCnt)
                    
                    malngEncKey(lngIndex + 1) = malngEncKey(lngIndex - 7) Xor malngEncKey(lngIndex)
                    malngEncKey(lngIndex + 2) = malngEncKey(lngIndex - 6) Xor malngEncKey(lngIndex + 1)
                    malngEncKey(lngIndex + 3) = malngEncKey(lngIndex - 5) Xor malngEncKey(lngIndex + 2)
                    CopyMemory abytTmp(0), malngEncKey(lngIndex + 3), 4&
                    
                    malngEncKey(lngIndex + 4) = malngEncKey(lngIndex - 4) Xor _
                                                (TE4(abytTmp(3)) And &HFF000000) Xor _
                                                (TE4(abytTmp(2)) And &HFF0000) Xor _
                                                (TE4(abytTmp(1)) And &HFF00&) Xor _
                                                (TE4(abytTmp(0)) And &HFF&)
                                              
                    malngEncKey(lngIndex + 5) = malngEncKey(lngIndex - 3) Xor malngEncKey(lngIndex + 4)
                    malngEncKey(lngIndex + 6) = malngEncKey(lngIndex - 2) Xor malngEncKey(lngIndex + 5)
                    malngEncKey(lngIndex + 7) = malngEncKey(lngIndex - 1) Xor malngEncKey(lngIndex + 6)
                    lngIndex = lngIndex + lngStep    ' Increment index
                    
                Next lngCnt
                
           Case Else
                InfoMsg "Invalid key length  (Key length:  " & _
                        CStr(mlngKeyLength) & ")" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                gblnStopProcessing = True
    End Select
                
LoadEncKey_256_CleanUp:
    Erase alngRcon()  ' Always empty array when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

LoadEncKey_256_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadEncKey_256_CleanUp

End Sub

Private Sub LoadDecKeyArray()
    
    ' Called by LoadEncKeyArray()
    
    Dim lngIdx      As Long
    Dim lngLoop     As Long
    Dim lngIndex    As Long
    Dim lngDataLen  As Long   ' 24-Jan-2011 No longer a passed parameter
    Dim abytTemp(4) As Byte
    
    Const ROUTINE_NAME As String = "LoadDecKeyArray"

    On Error GoTo LoadDecKeyArray_Error
    
    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadDecKeyArray_CleanUp
    End If
    
    ReDim malngDecKey(0 To 119)   ' Size decryption key array
    
    lngIndex = 0
    lngDataLen = mlngDataChunk \ 4
    lngIdx = lngDataLen * mlngEncryptRnds
    
    ' Invert the order of the round keys
    ' based on number of elements used
    For lngLoop = 0 To mlngEncryptRnds
        
        CopyMemory malngDecKey(lngIndex), malngEncKey(lngIdx), (lngDataLen * 4&)
        lngIndex = (lngIndex + lngDataLen)
        lngIdx = (lngIdx - lngDataLen)
        
    Next lngLoop

    ' An error occurred or user opted to STOP processing
    If gblnStopProcessing Then
        GoTo LoadDecKeyArray_CleanUp
    End If
    
    ' Apply the inverse MixColumn transform to
    ' all round keys except the first and the last
    For lngIndex = lngDataLen To (lngDataLen * mlngEncryptRnds - 1)
            
        CopyMemory abytTemp(0), malngDecKey(lngIndex), 4&
        
        malngDecKey(lngIndex) = TD0(TE4(abytTemp(0)) And &HFF&) Xor _
                                TD1(TE4(abytTemp(1)) And &HFF&) Xor _
                                TD2(TE4(abytTemp(2)) And &HFF&) Xor _
                                TD3(TE4(abytTemp(3)) And &HFF&)
    Next lngIndex

LoadDecKeyArray_CleanUp:
    On Error GoTo 0    ' Nullify this error trap
    Exit Sub

LoadDecKeyArray_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume LoadDecKeyArray_CleanUp
    
End Sub

Private Sub LoadConstArrays()

    ' Called by EvaluateKey()
    
    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim bytHold     As Byte
    Dim bytValue    As Byte
    Dim abytTemp()  As Byte
    Dim abytTbl_P() As Byte
    Dim abytTbl_L() As Byte

    Const MAX_SIZE As Long = 255

    ' Size temp work arrays
    ReDim abytTemp(8)
    ReDim abytTbl_P(MAX_SIZE)
    ReDim abytTbl_L(MAX_SIZE)
    
    ' Size constant encryption arrays
    ReDim TE0(MAX_SIZE)
    ReDim TE1(MAX_SIZE)
    ReDim TE2(MAX_SIZE)
    ReDim TE3(MAX_SIZE)
    ReDim TE4(MAX_SIZE)
    
    ' Size constant decryption arrays
    ReDim TD0(MAX_SIZE)
    ReDim TD1(MAX_SIZE)
    ReDim TD2(MAX_SIZE)
    ReDim TD3(MAX_SIZE)
    ReDim TD4(MAX_SIZE)
    
    ' Use 3 as primitive root to
    ' generate power and log tables
    abytTbl_L(0) = 0
    abytTbl_L(1) = 0
    abytTbl_L(3) = 1
    abytTbl_P(0) = 1
    abytTbl_P(1) = 3
    
    For lngIndex = 2 To MAX_SIZE
        
        If (abytTbl_P(lngIndex - 1) And &H80) Then
            abytTbl_P(lngIndex) = abytTbl_P(lngIndex - 1) Xor ((abytTbl_P(lngIndex - 1) And 127) * 2) Xor &H1B
        Else
            abytTbl_P(lngIndex) = abytTbl_P(lngIndex - 1) Xor (abytTbl_P(lngIndex - 1) * 2)
        End If
        
        abytTbl_L(abytTbl_P(lngIndex)) = lngIndex
        
    Next lngIndex

    ' Each bit is Xor with itself shifted one bit
    TE4(0) = &H63636363
    TD4(&H63) = 0
    
    For lngIndex = 1 To MAX_SIZE
        
        ' multiplicative inverse
        bytHold = abytTbl_P(MAX_SIZE - abytTbl_L(lngIndex))
        bytValue = bytHold
        
        ' 08-Mar-2010 Replaced multiple "If"
        '             statements with For..Next loop
        ' Rotate Left Byte
        For lngIdx = 1 To 4
            bytValue = IIf(bytValue And &H80, (bytValue And 127) * 2 Or 1, bytValue * 2)
            bytHold = bytHold Xor bytValue
        Next lngIdx
        
        bytHold = bytHold Xor &H63
        
        ' Load forward table
        abytTemp(0) = bytHold
        abytTemp(1) = abytTemp(0)
        abytTemp(2) = abytTemp(0)
        abytTemp(3) = abytTemp(0)
        CopyMemory TE4(lngIndex), abytTemp(0), 4&

        ' Load reverse table
        abytTemp(0) = lngIndex
        abytTemp(1) = abytTemp(0)
        abytTemp(2) = abytTemp(0)
        abytTemp(3) = abytTemp(0)
        CopyMemory TD4(bytHold), abytTemp(0), 4&

    Next lngIndex

    ' calculate forward and reverse tables
    For lngIndex = 0 To MAX_SIZE
        
        bytHold = TE4(lngIndex) And &HFF&
        
        If (bytHold And &H80) Then
            abytTemp(0) = ((bytHold And 127) * 2) Xor &H1B
            abytTemp(3) = bytHold Xor abytTemp(0)
        Else
            abytTemp(0) = bytHold * 2
            abytTemp(3) = bytHold Xor abytTemp(0)
        End If
        
        abytTemp(2) = bytHold
        abytTemp(1) = bytHold
        
        CopyMemory abytTemp(4), abytTemp(0), 4&
        CopyMemory TE0(lngIndex), abytTemp(0), 4&
        CopyMemory TE1(lngIndex), abytTemp(3), 4&
        CopyMemory TE2(lngIndex), abytTemp(2), 4&
        CopyMemory TE3(lngIndex), abytTemp(1), 4&
        bytHold = TD4(lngIndex) And &HFF&
        
        ' x.y= AntiLog(Log(x) + Log(bytHold))
        If bytHold = 0 Then
            abytTemp(3) = 0
            abytTemp(2) = 0
            abytTemp(1) = 0
            abytTemp(0) = 0
        Else
            abytTemp(3) = abytTbl_P((CLng(abytTbl_L(&HB)) + CLng(abytTbl_L(bytHold))) Mod MAX_SIZE)
            abytTemp(2) = abytTbl_P((CLng(abytTbl_L(&HD)) + CLng(abytTbl_L(bytHold))) Mod MAX_SIZE)
            abytTemp(1) = abytTbl_P((CLng(abytTbl_L(&H9)) + CLng(abytTbl_L(bytHold))) Mod MAX_SIZE)
            abytTemp(0) = abytTbl_P((CLng(abytTbl_L(&HE)) + CLng(abytTbl_L(bytHold))) Mod MAX_SIZE)
        End If
        
        CopyMemory abytTemp(4), abytTemp(0), 4&
        CopyMemory TD0(lngIndex), abytTemp(0), 4&
        CopyMemory TD1(lngIndex), abytTemp(3), 4&
        CopyMemory TD2(lngIndex), abytTemp(2), 4&
        CopyMemory TD3(lngIndex), abytTemp(1), 4&

    Next lngIndex

End Sub

Private Function LoadRconArray() As Long()

    ' Called by LoadEncKey_128()
    '           LoadEncKey_160()
    '           LoadEncKey_192()
    '           LoadEncKey_224()
    '           LoadEncKey_256()
    
    Dim lngTemp    As Long
    Dim lngIndex   As Long
    Dim alngRcon() As Long
    
    ReDim alngRcon(0 To 28)
    lngTemp = 1
    
    ' Load the Rcon table
    For lngIndex = 0 To 28
        
        alngRcon(lngIndex) = lngTemp
        
        If (lngTemp And &H80) Then
            lngTemp = ((lngTemp And 127) * 2) Xor &H1B
        Else
            lngTemp = lngTemp * 2
        End If
        
    Next lngIndex

    LoadRconArray = alngRcon()
    
End Function

Private Sub ResetVariables()

    Erase mabytTmp()
    Erase malngEncKey()
    Erase malngDecKey()
    
    Erase TE0()
    Erase TE1()
    Erase TE2()
    Erase TE3()
    Erase TE4()
    
    Erase TD0()
    Erase TD1()
    Erase TD2()
    Erase TD3()
    Erase TD4()
    
    mcurOrigLength = 0@
        
End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    ResetVariables
    Erase mabytPassword()  ' Always start with empty arrays
    
    ' Preset property values
    KeyLength = 128        ' Minimum key length
    BlockSize = 128        ' Minimum block size
    EncryptRounds = 1      ' Number of rounds of encryption
    CreateNewFile = True   ' Default is to create a new output file

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    ResetVariables
    Erase mabytPassword()  ' Always empty arrays when not needed

End Sub


