VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArcFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cArcFour  (clsArcFour.cls)
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
' Description:   ArcFour is a stream cipher symmetric key algorithm.  Very
'                similar to RC4 that was developed in 1987 by Ronald Rivest
'                and kept as a trade secret by RSA Data Security.  On
'                September 9, 1994, the RC4 algorithm was anonymously posted
'                on the Internet on the Cyperpunks’ anonymous remailers list.
'
'                The name "RC4" is trademarked.  The current status seems to
'                be that "unofficial" implementations are legal, but cannot
'                use the RC4 name.  RC4 is often referred to as "ARCFOUR" or
'                "ARC4" (meaning Alleged RC4, because RSA has never officially
'                released the algorithm).  This is to avoid possible trademark
'                problems.
'
'                ArcFour uses a variable length key from 1 to 256 bytes to
'                initialize a 256-byte state table.  The state table is
'                used for subsequent generation of pseudo-random bytes
'                and then to generate a pseudo-random stream which is
'                XORed with the plaintext to give the ciphertext.  Each
'                element in the state table is Swapped at least once.
'
'                The ArcFour key is often limited to 40 bits, because of export
'                restrictions but it is sometimes used as a 128 bit key.
'                It has the capability of using keys between 1 and 2048 bits.
'                ArcFour is used in many commercial software packages such as
'                Lotus Notes and Oracle Secure SQL.  It is also part of the
'                Cellular Specification.
'
'                The ArcFour algorithm works in two phases, key setup and
'                ciphering. Key setup is the first and most difficult phase of
'                this algorithm. During a N-bit key setup (N being your key
'                length), the encryption key is used to generate an encrypting
'                variable using two arrays, state and key, and N-number of
'                mixing operations. These mixing operations consist of swapping
'                bytes, modulo operations, and other formulas. A modulo
'                operation is the process of yielding a remainder from division.
'
'                Once the encrypting variable is produced from the key setup,
'                it enters the ciphering phase, where it is XORed with the plain
'                text message to create and encrypted message. XOR is the
'                logical operation of comparing two binary bits. If the bits are
'                different, the result is 1.  If the bits are the same, the
'                result is 0. Once the receiver gets the encrypted message, he
'                decrypts it by XORing the encrypted message with the same
'                encrypting variable.
'
'                **********************************************************
' SPECIAL NOTE:  The encryption process has been enhanced.
'
'                - Encryption process can be performed multiple times.
'                  See property EncryptRounds().
'
'                - Password key length is expanded to meet and exceed AES
'                  standards.  Key lengths are 128 to 448 in 32 bit
'                  increments and 448 to 1024 in 64 bit increments.  See
'                  property KeyLength().
'
'                - S-Box data will be mixed multiple times.
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
' Reference:     RC4 Encryption Algorithm
'                www.vocal.com/RC4.pdf
'
'                RC4 from Wikipedia, the free encyclopedia
'                http://en.wikipedia.org/wiki/RC4_(cipher)
'
'                Ronald Rivest Home Page
'                http://theory.lcs.mit.edu/~rivest/
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 28-NOV-2003  Kenneth Ives  kenaso@tx.rr.com
'              Documented, modified, updated variable names
' 04-AUG-2007  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote LoadXBoxArray() routine for mixing values 0-255
'              without creating duplicates
' 12-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified to handle files larger than 2gb
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified ArcFour_File() and ArcFour_String() routines to
'              encrypt appended data information for additional security.
'              Moved ExpandData() routine to basCommon.bas module.
' 17-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated property EncryptRounds().  Number of encryption rounds
'              is now user-defined.  See constant MAX_ROUNDS.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Moved GetBlockSize() routine to common module.
'              Set key length to increments of 32 bits for 128 to 448 and
'              increments of 64 bits for 448 to 1024.
'              Modified LoadXBoxArray() routine to use utilize the new key
'              lengths.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in ArcFour_String().
' 29-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed LoadWorkArrays() routine.  Now makes a call to a common
'              routine named LoadXBoxArray().
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated EvaluateKey() and Encrypt_Decrypt() routines to use
'              only byte arrays.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while encrypting or
'                decrypting data (Encrypt_Decrypt() routine).
'              - Removed obsolete code.
'              - Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within ArcFour_File() routine.
'                4. Split size testing in ArcFour_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated ArcFour_File() routine to reference new property.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "clsArcFour"
  Private Const MAX_FILE_SIZE As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MAX_BYTE      As Long = 256
  Private Const MAX_ROUNDS    As Long = 10
  Private Const MB_5          As Long = &H500000           ' 5242880 bytes
  
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
  Private mlngKeyLength     As Long
  Private mabytSBox()       As Byte
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
        
    ' 03-Jun-2014 Added key length 128
    ' 01-Jan-2009 Set increments of 32 bits for 128 to 448
    '             Set increments of 64 bits for 448 to 1024
    Select Case lngKeyLength
           Case 1024: mlngKeyLength = 1024
           Case 960:  mlngKeyLength = 960
           Case 896:  mlngKeyLength = 896
           Case 832:  mlngKeyLength = 832
           Case 768:  mlngKeyLength = 768
           Case 704:  mlngKeyLength = 704
           Case 640:  mlngKeyLength = 640
           Case 576:  mlngKeyLength = 576
           Case 512:  mlngKeyLength = 512
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

Public Property Let Password(ByVal vntData As Variant)

    ' Used when assigning a value to the password property, on the
    ' left side of an assignment.
    '
    ' Syntax:     X.Password = "User Supplied Password"
    '                            or
    '             X.Password = alngPwd()  <-- Highly recommend using a byte array
    
    Dim lngIndex As Long
    
    Const ROUTINE_NAME As String = "Password Property"
    
    On Error GoTo Password_Error

    Erase mabytPassword()
    
    ' Is password data in an array
    If CBool(IsArrayInitialized(vntData)) Then
    
        ReDim mabytPassword(UBound(vntData))
        
        For lngIndex = 0 To UBound(vntData) - 1
            mabytPassword(lngIndex) = CByte(vntData(lngIndex))
        Next lngIndex
    
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
' Routine:       ArcFour_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objArcFour.ArcFour_File abytFile(), blnEncrypt               ' Encrypt/Decrypt data
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
'              Modified to encrypt appended data information for additional
'              security.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function ArcFour_File(ByRef abytFile() As Byte, _
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
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curAmtLeft   As Currency
    Dim abytData()   As Byte
    Dim abytSize(7)  As Byte
    Dim objBigFiles  As cBigFiles
    
    Const ROUTINE_NAME As String = "ArcFour_File"
    
    On Error GoTo ArcFour_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles  ' Instantiate class object
    Erase abytData()                 ' Always start with empty arrays
    Erase abytSize()
    ResetVariables                   ' Reset class variables
    
    ' Expand password key and load S-Boxes
    If Not EvaluateKey() Then
        GoTo ArcFour_File_CleanUp
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
        GoTo ArcFour_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo ArcFour_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)  ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                          ' Set pointer to capture the last 8 bytes
            
            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If
            
            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 8 bytes of source file
            .API_CloseFile hFile1                             ' close source file
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo ArcFour_File_CleanUp
            End If
    
            MixAppendedData abytSize()                                      ' Decrypt appended data
            CopyMemory mcurOrigLength, abytSize(0), (UBound(abytSize) + 1)  ' Capture original file size [0-7]
            
            ' Decryption error test
            If (curAmtLeft < 1) Or _
               (mcurOrigLength < 1) Or _
               (mcurOrigLength > MAX_FILE_SIZE) Then
            
                InfoMsg "This file cannot be Decrypted, it is" & vbNewLine & _
                        "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                        "Current size:  " & Format$(curAmtLeft, "#,##0") & " bytes" & vbNewLine & _
                        "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                        "Max allowed size:  " & Format$(MAX_FILE_SIZE, "#,##0") & " bytes" & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
                
                gblnStopProcessing = True
                GoTo ArcFour_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds  ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo ArcFour_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
    
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo ArcFour_File_CleanUp
        End If
        
        Do
            Erase abytData()                               ' empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)        ' Process the source file in chunks
            ReDim abytData(lngBlockSize - 1)               ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)   ' Adjust amount of file left to process
            
            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If
                       
            curFilePos1 = curFilePos1 + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
            
            If blnEncrypt Then
        
                ' Encrypt data
                If Not Encrypt_Decrypt(abytData()) Then
                    gblnStopProcessing = True
                    Exit Do  ' exit Do..Loop
                End If
            
            Else
                
                ' Decrypt data
                If Not Encrypt_Decrypt(abytData()) Then
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
            GoTo ArcFour_File_CleanUp
        End If
    
        .API_CloseFile hFile1    ' Close all files opened by this routine
        .API_CloseFile hFile2
        
        ' Add important data to end of encrypted file
        If blnEncrypt Then
            
            CopyMemory abytSize(0), mcurOrigLength, 8&       ' Save original file size [0-7]
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
                GoTo ArcFour_File_CleanUp
                
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

ArcFour_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Verify all open files have been closed
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory
    
    ResetVariables    ' Reset class variables
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ArcFour_File = False
    Else
        ArcFour_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
ArcFour_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ArcFour_File_CleanUp

End Function

' ***************************************************************************
' Routine:       ArcFour_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)        ' convert string data to byte array
'                objArcFour.ArcFour_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
'              Modified to encrypt appended data information for additional
'              security.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Correct data size capture for decryption.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Split size testing as per Joe Sova's suggestion.
' ***************************************************************************
Public Function ArcFour_String(ByRef abytData() As Byte, _
                      Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "ArcFour_String"

    On Error GoTo ArcFour_String_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    ' Is there any data to process?
    If Not CBool(IsArrayInitialized(abytData())) Then
        gblnStopProcessing = True     ' Set flag to stop processing
        Exit Function
    End If

    Erase abytSize()                                    ' Always start with empty arrays
    ResetVariables                                      ' Reset class variables
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds   ' Save string length for progress calcs
    mcurAccrued = 0@                                    ' Init accrued progress
        
    ' Expand password key and load S-Boxes
    If Not EvaluateKey() Then
        GoTo ArcFour_String_CleanUp
    End If
    
    ' encrypt the data
    If blnEncrypt Then
        
        mcurOrigLength = CCur(UBound(abytData))         ' Get string length for progress display
        CopyMemory abytSize(0), mcurOrigLength, 8&      ' Copy original length to temp array
        MixAppendedData abytSize()                      ' Encrypt appended data
        
        If Encrypt_Decrypt(abytData()) Then
            
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
        If Encrypt_Decrypt(abytData()) Then
                
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
    
ArcFour_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        ArcFour_String = False
    Else
        ArcFour_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
ArcFour_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume ArcFour_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by ArcFour_File()
    '           ArcFour_String()
    
    Dim lngKeyLen As Long
    Dim abytPwd() As Byte
    
    Const ROUTINE_NAME As String = "EvaluateKey"

    On Error GoTo EvaluateKey_Error

    EvaluateKey = False  ' Preset flag to FALSE
    
    If Not CBool(IsArrayInitialized(mabytPassword)) Then
        InfoMsg "Need a password before data can be encrypted/decrypted." & _
                vbNewLine & vbNewLine & _
                "Min length - " & MIN_PWD_LENGTH & vbNewLine & _
                "Max length - " & MAX_PWD_LENGTH & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        Exit Function
    End If
    
    Erase abytPwd()   ' Always start with empty arrays
    Erase mabytSBox()
    
    ReDim mabytSBox(MAX_BYTE)  ' Size S-Box array
        
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    mabytSBox() = LoadXBoxArray(abytPwd())               ' Create a mixed byte array using ASCII
                                                         ' values 0-255 with no duplicate data
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo EvaluateKey_CleanUp
    End If
    
    EvaluateKey = True  ' Good finish
    
EvaluateKey_CleanUp:
    Erase abytPwd()   ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function Encrypt_Decrypt(ByRef abytData() As Byte) As Boolean

    ' Called by ArcFour_File()
    '           ArcFour_String()
    
    Dim lngRound      As Long
    Dim lngIndex1     As Long
    Dim lngIndex2     As Long
    Dim lngTmpIdx     As Long
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim abytSBox()    As Byte
    Dim abytInput()   As Byte
    Dim abytOutput()  As Byte

    Const ROUTINE_NAME As String = "Encrypt_Decrypt"

    On Error GoTo Encrypt_Decrypt_Error

    lngIndex1 = 0  ' Initialize outside the loop
    lngIndex2 = 0
    
    lngDataLength = UBound(abytData) + 1  ' Get size of incoming array
    
    ReDim abytInput(lngDataLength)   ' Size input work array
    ReDim abytOutput(lngDataLength)  ' Size output array
    ReDim abytSBox(MAX_BYTE)         ' Size temp S-Box array
    
    CopyMemory abytInput(0), abytData(0), lngDataLength   ' Load input array
            
    ' Make multiple passes for extra security
    For lngRound = 1 To mlngRounds
    
        ' Work with a copy of original S-Box array
        CopyMemory abytSBox(0), mabytSBox(0), MAX_BYTE   ' Copy data into temp s-box array
        
        ' Encrypt/Decrypt data
        For lngOffSet = 0 To lngDataLength - 1
        
            ' Increment S-Box indices
            lngIndex1 = (lngIndex1 + 1) Mod MAX_BYTE
            lngIndex2 = (lngIndex2 + abytSBox(lngIndex1)) Mod MAX_BYTE
            
            ' Swap data in S-Box array
            SwapBytes abytSBox(lngIndex1), abytSBox(lngIndex2)
            
            ' Create temp index
            lngTmpIdx = (CLng(abytSBox(lngIndex1)) + CLng(abytSBox(lngIndex2))) Mod MAX_BYTE
            
            ' Perform an XoR with data
            ' and store in output array
            abytOutput(lngOffSet) = abytInput(lngOffSet) Xor abytSBox(lngTmpIdx)
            
            ' 19-Feb-2012 Update progress bar. Separate
            '             steps for easier debugging
            mcurAccrued = mcurAccrued + 1
            lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
            RaiseEvent CipherProgress(lngProgress)
                    
        Next lngOffSet
    
        ' Copy manipulated data back into the input array
        ReDim abytInput(lngDataLength)
        CopyMemory abytInput(0), abytOutput(0), lngDataLength
    
        ReDim abytOutput(lngDataLength)  ' resize output array
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngRound
    
    ' Copy manipulated data back into the original array
    ReDim abytData(lngDataLength - 1)
    CopyMemory abytData(0), abytInput(0), lngDataLength
    
Encrypt_Decrypt_CleanUp:
    DoEvents
    If gblnStopProcessing Then
        Encrypt_Decrypt = False   ' An error occurred or user opted to STOP processing
    Else
        Encrypt_Decrypt = True    ' Good Finish
    End If

    On Error GoTo 0    ' Nullify this error trap
    Exit Function

Encrypt_Decrypt_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Encrypt_Decrypt_CleanUp

End Function

' ***************************************************************************
' Routine:       SwapBytes
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
Private Sub SwapBytes(ByRef bytValue1 As Byte, _
                      ByRef bytValue2 As Byte)

    ' Swap byte values (0 to 255)

    Dim bytHold As Byte
    
    bytHold = bytValue1
    bytValue1 = bytValue2
    bytValue2 = bytHold

End Sub

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines the size of the data to be processed.  Smaller
'                block sizes is more appealing on the progress display.
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

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select
    
    RaiseEvent CurrentBlockSize(GetBlockSize)
    
End Function

Private Sub ResetVariables()

    ' Always empty arrays and variants
    ' when not needed.  Saves resources.
    Erase mabytSBox()
    mcurOrigLength = 0@
    mcurCurrByteCnt = 0@

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    ResetVariables
    Erase mabytPassword()
    
    ' Preset property values
    KeyLength = 256        ' Default value
    EncryptRounds = 1      ' Number of rounds of encryption
    CreateNewFile = True   ' Default is to create a new output file
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    ResetVariables
    Erase mabytPassword()

End Sub
