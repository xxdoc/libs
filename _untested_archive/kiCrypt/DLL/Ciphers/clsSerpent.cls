VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSerpent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cSerpent  (clsSerpent.cls)
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
' Description:   Serpent was designed by Ross Anderson, Eli Biham and
'                Lars Knudsen as a candidate for the Advanced Encryption
'                Standard. It has been selected as one of the five
'                finalists in the AES competition. Serpent is faster
'                than DES and more secure than Triple DES. It provides
'                users with a very high level of assurance that no
'                shortcut attack will be found. To achieve this, the
'                algorithm's designers limited themselves to well
'                understood cryptography mechanisms, so that they could
'                rely on the wide experience and proven techniques of
'                block cipher cryptanalysis. The algorithm uses twice as
'                many rounds as are necessary to block all currently
'                known shortcut attacks. This means that Serpent should
'                be safe against as yet unknown attacks that may be
'                capable of breaking the standard 16 rounds used in many
'                types of encryption today. However, the fact that
'                Serpent uses so many rounds means that it is the
'                slowest of the five AES (Advanced Encryption Standard)
'                finalists. But this should not be an issue because it
'                still outperforms Triple DES.  The algorithm's designers
'                maintain that Serpent has a service life of at least a
'                century.
'
'                Serpent is a 128-bit block cipher, meaning that data is
'                encrypted and decrypted in 128-bit chunks. The key
'                length can vary, but for the purposes of the AES it is
'                defined to be either 128, 192, or 256 bits. This block
'                size and variable key length is standard among all AES
'                candidates and was one of the major design requirements
'                specified by NIST (National Institute of Standards and
'                Technology). The Serpent algorithm uses 32 rounds, or
'                iterations of the main algorithm.
'
'                Like DES, Serpent includes an initial and final
'                permutation of no cryptographic significance; these
'                permutations are used to optimize the data before
'                encryption. The detailed description of the actual
'                algorithm is contained in the official Serpent paper
'                submitted for the AES by the algorithm's designers.
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
'                - Added mixing routine to initial loading of password key
'                  array.  See EvaluateKey() routine.
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
' Note:          After encryption, data sizes will not match original
'                sizes.  This is due to internal padding and the storing
'                of information required to decrypt the data later.
'
' References:    Serpent Encryption Algorithm
'                http://www.cl.cam.ac.uk/~rja14/serpent.html
'
'                The algorithm specifications.
'                http://www.cl.cam.ac.uk/~rja14/Papers/serpent.pdf
'
'                Full submission package for AES
'                http://www.cl.cam.ac.uk/~rja14/Papers/serpent.tar.gz
'
'                Most of the description above was obtained from
'                http://www.tropsoft.com/strongenc/serpent.htm
'
'                An Introduction to Using Keys in Cryptography
'                By DI Management Services Pty Ltd, Sydney Australia
'                http://www.di-mgt.com.au/cryptokeys.html
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 08-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Converted from C to Visual Basic 6.0
' 17-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated property EncryptRounds().  Number of encryption rounds
'              is now user-defined.  See constant MAX_ROUNDS.
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet and exceed AES standards.
'              (128, 192, 256, 384, 512, 1024 bits)  Modified EvaluateKey()
'              routine to use utilize the new key lengths.
' 27-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed initial KK() array loading so that it now can access
'              up to 128 bytes of the password array.  See EvaluateKey()
'              routine.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed OffSet range in EncryptData() and DecryptData() routines.
'              Fixed an intermittent bug in EvaluateKey() routine.
'              Moved GetBlockSize() routine to common module.
'              Set key length to increments of 32 bits for 128 to 448 and
'              increments of 64 bits for 448 to 1024.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a potential bug in the looping logic in EvaluateKey().
'              Updated Documentation.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
'              Correct data size capture for decryption in Serpent_String().
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Enhanced initial loading of key array in EvaluateKey() routine.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to see progress display while encrypting or
'              decrypting data (EncryptData() and DecryptData() routines).
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
'                   within Serpent_File() routine.
'                4. Split size testing in Serpent_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Renamed EncryptBlock() to BlockEncrypt() and DecryptBlock()
'                to BlockDecrypt() routines to be consistant with other
'                cipher modules.
'              - Moved progress bar code from BlockEncrypt() and
'                BlockDecrypt() routines to EncryptData() and DecryptData()
'                routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Serpent_File() routine to reference new property.
'              - Corrected a flaw in the initial loading of temp key array
'                in EvaluateKey() routine.
'              - Increased speed by inserting bitwise routines into this
'                class and removing some IF..Then decisions.
' 05-Dec-2013  Kenneth Ives  kenaso@tx.rr.com
'              Modified EvaluateKey() routine to perform multiple rounds
'              of mixing the values prior to loading the primary key
'              [malngSubKey()] array.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME    As String = "clsSerpent"
  Private Const MAX_FILE_SIZE  As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MB_5           As Long = &H500000           ' 5242880 bytes
  Private Const BLOCK_SIZE     As Long = 16
  Private Const DATA_BLOCK     As Long = 32
  Private Const MAX_ROUNDS     As Long = 10
  Private Const PHI            As Long = &H9E3779B9  ' Constant used in key schedule

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
  Private mlngRounds        As Long
  Private mlngKeyRounds     As Long
  Private mlngKeyLength     As Long
  Private malngSubKey()     As Long
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
' Routine:       Serpent_File
'
' Description:   Function to encrypt/decrypt a file. 
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objSerpent.Serpent_File abytFile(), blnEncrypt               ' Encrypt/Decrypt data
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
' 08-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              Removed obsolete code. Updated documentation.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to earlier an position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function Serpent_File(ByRef abytFile() As Byte, _
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
    Dim curAmtLeft   As Currency
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim abytData()   As Byte
    Dim abytSize(11) As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "Serpent_File"
    
    On Error GoTo Serpent_File_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    Set objBigFiles = New cBigFiles   ' Instantiate class object
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    
    If Not EvaluateKey() Then
        GoTo Serpent_File_CleanUp
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
        GoTo Serpent_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Serpent_File_CleanUp
        End If
            
        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then
            mcurOrigLength = curAmtLeft   ' Capture total byte count
        Else
            curAmtLeft = curAmtLeft - (UBound(abytSize) + 1)   ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft                           ' Set pointer to capture the last 12 bytes
            
            If Not .OpenReadOnly(strSource, hFile1) Then       ' Open the source file
                gblnStopProcessing = True
            End If
            
            .API_ReadFile hFile1, curFilePos1, abytSize()      ' Capture last 12 bytes of source file
            .API_CloseFile hFile1                              ' close source file
            
            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo Serpent_File_CleanUp
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
                GoTo Serpent_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Serpent_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Serpent_File_CleanUp
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
                       
            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)  ' Adjust pointers accordingly
            
            ' If nothing else to read then leave
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
            GoTo Serpent_File_CleanUp
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
                GoTo Serpent_File_CleanUp
                
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

Serpent_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Close all files opened by this routine
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory
    
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Serpent_File = False
    Else
        Serpent_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Serpent_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Serpent_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Serpent_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)        ' convert string data to byte array
'                objSerpent.Serpent_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
' 08-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Split size testing as per Joe Sova's suggestion.
' ***************************************************************************
Public Function Serpent_String(ByRef abytData() As Byte, _
                      Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "Serpent_String"

    On Error GoTo Serpent_String_Error
    
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
    
    Erase abytSize()                                   ' Always start with empty arrays
    mcurMaxSize = CCur(UBound(abytData)) * mlngRounds  ' Save file size for progress calcs
    mcurAccrued = 0@                                   ' Init accrued progress
    
    If Not EvaluateKey() Then
        GoTo Serpent_String_CleanUp
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
    
Serpent_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Serpent_String = False
    Else
        Serpent_String = True
    End If
    
    Erase abytSize()
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Serpent_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Serpent_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Serpent_File()
    '           Serpent_String()
    
    Dim lngIdx    As Long
    Dim lngLoop   As Long
    Dim lngTemp   As Long
    Dim lngIndex  As Long
    Dim lngKeyLen As Long
    Dim KK()      As Long   ' Temp key array
    Dim WW()      As Long   ' Temp work array
    Dim ZZ()      As Long   ' All zeroes
    Dim abytPwd() As Byte   ' password array

    Const ROUTINE_NAME As String = "EvaluateKey"
    Const WORK_SIZE    As Long = 132
    
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
    
    ' Size arrays
    ReDim malngSubKey(0 To 32, 0 To 3)
    ReDim KK(WORK_SIZE)
    ReDim WW(WORK_SIZE)
    ReDim ZZ(WORK_SIZE)
    
    ' Preload array with zeroes
    For lngIndex = 0 To WORK_SIZE - 1
        ZZ(lngIndex) = 0&
    Next lngIndex
    
    '           +--------------------- Target array
    '           |      +-------------- Source array
    '           |      |          |--- # of bytes to copy
    CopyMemory KK(0), ZZ(0), (WORK_SIZE * 4&)            ' Load temp key array with zeroes
    CopyMemory WW(0), ZZ(0), (WORK_SIZE * 4&)            ' Load temp work array with zeroes
    
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    Erase abytPwd()                                      ' Always start with an empty array
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    lngIdx = 0                                           ' Init temp password index
    
    '***********************************************************************
    ' 02-Nov-2009 Enhanced initial loading of temp key array so all eight
    '             elements will have a unique starting value
    '
    ' Load temp key array in positions 8-15. Multiple looping due to unknown
    ' key length.  Key lengths will vary 16-128 bytes (128-1024 bits).
    For lngLoop = 1 To 4
        For lngIndex = 0 To 7
        
            ' Create one long integer and insert into Key array
            CopyMemory KK(lngIndex), abytPwd(lngIdx Mod lngKeyLen), 4&
            lngIdx = lngIdx + 4   ' Increment password index
            
            ' Update temp storage area (Enhanced)
            KK(lngIndex + 8) = (KK(lngIndex) Xor KK(lngIndex + 1)) Xor KK(lngIndex + 8)
                
        Next lngIndex
    Next lngLoop
    
    CopyMemory KK(0), ZZ(0), 32&   ' Zero out key array elements 0-7
    
    ' Mix temp key data and load into positions 0-7 (Enhanced)
    '     |----------Input-----------|  |---------Output---------|
    RND01 KK(8), KK(9), KK(10), KK(11), KK(0), KK(1), KK(2), KK(3)
    RND03 KK(12), KK(13), KK(14), KK(15), KK(4), KK(5), KK(6), KK(7)
    '***********************************************************************
    
    ' Transfer temp key array data to first
    ' eight positions of temp work array
    '
    ' for(i=0; i<keyLen/32; i++)
    '   w[i]=key->key[i];
    CopyMemory WW(0), KK(0), 32&   ' Copy key array to work array elements 0-7
    CopyMemory KK(0), ZZ(0), 64&   ' Zero out key array elements 0-15
         
    With mobjMath32
        ' Calculate next eight positions
        ' in temp work array (8-15)
        '
        ' for(i=8; i<16; i++)
        '   w[i]=ROL(w[i-8]^w[i-5]^w[i-3]^w[i-1]^PHI^w(i-8),11);
        For lngIndex = 8 To 15
            
            lngTemp = WW(lngIndex - 8) Xor WW(lngIndex - 5)
            lngTemp = lngTemp Xor WW(lngIndex - 3)
            lngTemp = lngTemp Xor WW(lngIndex - 1)
            lngTemp = lngTemp Xor PHI
            lngTemp = lngTemp Xor WW(lngIndex - 8)
            WW(lngIndex) = .w32Rotate(lngTemp, 11)
            
        Next lngIndex
            
        ' Overlay first eight positions
        ' in temp work array (0-7)
        '
        ' for(i=0; i<8; i++)
        '   w[i]=w[i+8];
         CopyMemory WW(0), WW(8), 32&   ' Copy work array data to elements 0-7
         CopyMemory WW(8), ZZ(0), 32&   ' Zero out work array elements 8-15
        
        ' Calculate remainder of temp work array (8-131)
        '
        ' for(i=8; i<132; i++)
        '   w[i]=ROL(w[i-8]^w[i-5]^w[i-3]^w[i-1]^PHI^i,11);
        For lngIndex = 8 To WORK_SIZE - 1
            
            lngTemp = WW(lngIndex - 8) Xor WW(lngIndex - 5)
            lngTemp = lngTemp Xor WW(lngIndex - 3)
            lngTemp = lngTemp Xor WW(lngIndex - 1)
            lngTemp = lngTemp Xor PHI
            lngTemp = lngTemp Xor lngIndex
            WW(lngIndex) = .w32Rotate(lngTemp, 11)
            
        Next lngIndex
    End With
    
    ' Calculate complete temp key array
    '        |----------Input-------------|  |---------Output-------------|
    '  RND03(w[  0], w[  1], w[  2], w[  3], k[  0], k[  1], k[  2], k[  3]);
    '  RND02(w[  4], w[  5], w[  6], w[  7], k[  4], k[  5], k[  6], k[  7]);
    '  RND01(w[  8], w[  9], w[ 10], w[ 11], k[  8], k[  9], k[ 10], k[ 11]);
    '  RND00(w[ 12], w[ 13], w[ 14], w[ 15], k[ 12], k[ 13], k[ 14], k[ 15]);
    '  RND31(w[ 16], w[ 17], w[ 18], w[ 19], k[ 16], k[ 17], k[ 18], k[ 19]);
    '  RND30(w[ 20], w[ 21], w[ 22], w[ 23], k[ 20], k[ 21], k[ 22], k[ 23]);
    '  RND29(w[ 24], w[ 25], w[ 26], w[ 27], k[ 24], k[ 25], k[ 26], k[ 27]);
    '  RND28(w[ 28], w[ 29], w[ 30], w[ 31], k[ 28], k[ 29], k[ 30], k[ 31]);
    '  RND27(w[ 32], w[ 33], w[ 34], w[ 35], k[ 32], k[ 33], k[ 34], k[ 35]);
    '  RND26(w[ 36], w[ 37], w[ 38], w[ 39], k[ 36], k[ 37], k[ 38], k[ 39]);
    '  RND25(w[ 40], w[ 41], w[ 42], w[ 43], k[ 40], k[ 41], k[ 42], k[ 43]);
    '  RND24(w[ 44], w[ 45], w[ 46], w[ 47], k[ 44], k[ 45], k[ 46], k[ 47]);
    '  RND23(w[ 48], w[ 49], w[ 50], w[ 51], k[ 48], k[ 49], k[ 50], k[ 51]);
    '  RND22(w[ 52], w[ 53], w[ 54], w[ 55], k[ 52], k[ 53], k[ 54], k[ 55]);
    '  RND21(w[ 56], w[ 57], w[ 58], w[ 59], k[ 56], k[ 57], k[ 58], k[ 59]);
    '  RND20(w[ 60], w[ 61], w[ 62], w[ 63], k[ 60], k[ 61], k[ 62], k[ 63]);
    '  RND19(w[ 64], w[ 65], w[ 66], w[ 67], k[ 64], k[ 65], k[ 66], k[ 67]);
    '  RND18(w[ 68], w[ 69], w[ 70], w[ 71], k[ 68], k[ 69], k[ 70], k[ 71]);
    '  RND17(w[ 72], w[ 73], w[ 74], w[ 75], k[ 72], k[ 73], k[ 74], k[ 75]);
    '  RND16(w[ 76], w[ 77], w[ 78], w[ 79], k[ 76], k[ 77], k[ 78], k[ 79]);
    '  RND15(w[ 80], w[ 81], w[ 82], w[ 83], k[ 80], k[ 81], k[ 82], k[ 83]);
    '  RND14(w[ 84], w[ 85], w[ 86], w[ 87], k[ 84], k[ 85], k[ 86], k[ 87]);
    '  RND13(w[ 88], w[ 89], w[ 90], w[ 91], k[ 88], k[ 89], k[ 90], k[ 91]);
    '  RND12(w[ 92], w[ 93], w[ 94], w[ 95], k[ 92], k[ 93], k[ 94], k[ 95]);
    '  RND11(w[ 96], w[ 97], w[ 98], w[ 99], k[ 96], k[ 97], k[ 98], k[ 99]);
    '  RND10(w[100], w[101], w[102], w[103], k[100], k[101], k[102], k[103]);
    '  RND09(w[104], w[105], w[106], w[107], k[104], k[105], k[106], k[107]);
    '  RND08(w[108], w[109], w[110], w[111], k[108], k[109], k[110], k[111]);
    '  RND07(w[112], w[113], w[114], w[115], k[112], k[113], k[114], k[115]);
    '  RND06(w[116], w[117], w[118], w[119], k[116], k[117], k[118], k[119]);
    '  RND05(w[120], w[121], w[122], w[123], k[120], k[121], k[122], k[123]);
    '  RND04(w[124], w[125], w[126], w[127], k[124], k[125], k[126], k[127]);
    '  RND03(w[128], w[129], w[130], w[131], k[128], k[129], k[130], k[131]);
    '
    ' 05-Dec-2013 Modified to perform multiple rounds of mixing the
    '             key [KK()] array values prior to loading the primary
    '             key [malngSubKey()] array
    For lngLoop = 1 To mlngKeyRounds
        
        ' If mixing key [KK()] array more than once
        If lngLoop > 1 Then
            CopyMemory WW(0), KK(0), (WORK_SIZE * 4&)   ' Transfer data from KK() to WW() array
            CopyMemory KK(0), ZZ(0), (WORK_SIZE * 4&)   ' Zero out key [KK()] array
        End If
          
        ' Process elements 0-127
        For lngIndex = 0 To 96 Step 32
            
            RND03 WW(lngIndex), WW(lngIndex + 1), WW(lngIndex + 2), WW(lngIndex + 3), _
                  KK(lngIndex), KK(lngIndex + 1), KK(lngIndex + 2), KK(lngIndex + 3)
            
            RND02 WW(lngIndex + 4), WW(lngIndex + 5), WW(lngIndex + 6), WW(lngIndex + 7), _
                  KK(lngIndex + 4), KK(lngIndex + 5), KK(lngIndex + 6), KK(lngIndex + 7)
            
            RND01 WW(lngIndex + 8), WW(lngIndex + 9), WW(lngIndex + 10), WW(lngIndex + 11), _
                  KK(lngIndex + 8), KK(lngIndex + 9), KK(lngIndex + 10), KK(lngIndex + 11)
            
            RND00 WW(lngIndex + 12), WW(lngIndex + 13), WW(lngIndex + 14), WW(lngIndex + 15), _
                  KK(lngIndex + 12), KK(lngIndex + 13), KK(lngIndex + 14), KK(lngIndex + 15)
            
            RND07 WW(lngIndex + 16), WW(lngIndex + 17), WW(lngIndex + 18), WW(lngIndex + 19), _
                  KK(lngIndex + 16), KK(lngIndex + 17), KK(lngIndex + 18), KK(lngIndex + 19)
            
            RND06 WW(lngIndex + 20), WW(lngIndex + 21), WW(lngIndex + 22), WW(lngIndex + 23), _
                  KK(lngIndex + 20), KK(lngIndex + 21), KK(lngIndex + 22), KK(lngIndex + 23)
            
            RND05 WW(lngIndex + 24), WW(lngIndex + 25), WW(lngIndex + 26), WW(lngIndex + 27), _
                  KK(lngIndex + 24), KK(lngIndex + 25), KK(lngIndex + 26), KK(lngIndex + 27)
            
            RND04 WW(lngIndex + 28), WW(lngIndex + 29), WW(lngIndex + 30), WW(lngIndex + 31), _
                  KK(lngIndex + 28), KK(lngIndex + 29), KK(lngIndex + 30), KK(lngIndex + 31)
                          
        Next lngIndex
            
        ' Process last four elements (128-131)
        RND03 WW(lngIndex), WW(lngIndex + 1), WW(lngIndex + 2), WW(lngIndex + 3), _
              KK(lngIndex), KK(lngIndex + 1), KK(lngIndex + 2), KK(lngIndex + 3)
                    
    Next lngLoop
                
    ' Calculate components of SubKey array
    '
    ' for(i=0; i<=32; i++)
    '   for(j=0; j<4; j++)
    '     key->subkeys[i][j] = KK[4*i+j];
    For lngIndex = 0 To 32
        For lngIdx = 0 To 3
            malngSubKey(lngIndex, lngIdx) = KK((4& * lngIndex + lngIdx) Mod WORK_SIZE)
        Next lngIdx
    Next lngIndex

EvaluateKey_CleanUp:
    Erase abytPwd()   ' Always empty arrays when not needed
    Erase KK()
    Erase WW()
    Erase ZZ()
    
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
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Serpent_File()
    '           Serpent_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim alngInput(4)  As Long
    Dim alngOutput(4) As Long
    Dim abytTemp(16)  As Byte
    
    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' The encoded length may have to be padded out.
    ' This normally does not take place until the
    ' last chunk of data is to be processed.
    If (lngDataLength Mod DATA_BLOCK <> 0) Then
        lngDataLength = lngDataLength + DATA_BLOCK - (lngDataLength Mod DATA_BLOCK)
        ReDim Preserve abytData(lngDataLength - 1)
    End If

    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE
    
        ' Transfer 16 bytes to temp work array
        CopyMemory abytTemp(0), abytData(lngOffSet), 16&
                
        ' Create bytes to four long integers
        CopyMemory alngInput(0), abytTemp(0), 4&
        CopyMemory alngInput(1), abytTemp(4), 4&
        CopyMemory alngInput(2), abytTemp(8), 4&
        CopyMemory alngInput(3), abytTemp(12), 4&
                
        ' main encryption process
        BlockEncrypt alngInput(), alngOutput()
        
        ' Convert four long integers to bytes
        CopyMemory abytTemp(0), alngOutput(0), 4&
        CopyMemory abytTemp(4), alngOutput(1), 4&
        CopyMemory abytTemp(8), alngOutput(2), 4&
        CopyMemory abytTemp(12), alngOutput(3), 4&
                
        ' Copy encrypted bytes back into data array
        CopyMemory abytData(lngOffSet), abytTemp(0), 16&
        
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
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False
    Else
        EncryptData = True
    End If

    On Error GoTo 0    ' Nullify this error trap
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Public Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Serpent_File()
    '           Serpent_String()
    
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim alngInput(4)  As Long
    Dim alngOutput(4) As Long
    Dim abytTemp(16)  As Byte
    
    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' 1-Jan-2009 Fixed lngOffSet range
    ' Encrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE
    
        ' Transfer 16 bytes to temp work array
        CopyMemory abytTemp(0), abytData(lngOffSet), 16&
                
        ' Create bytes to four long integers
        CopyMemory alngInput(0), abytTemp(0), 4&
        CopyMemory alngInput(1), abytTemp(4), 4&
        CopyMemory alngInput(2), abytTemp(8), 4&
        CopyMemory alngInput(3), abytTemp(12), 4&
                
        ' main decryption process
        BlockDecrypt alngInput(), alngOutput()
        
        ' Convert four long integers to bytes
        CopyMemory abytTemp(0), alngOutput(0), 4&
        CopyMemory abytTemp(4), alngOutput(1), 4&
        CopyMemory abytTemp(8), alngOutput(2), 4&
        CopyMemory abytTemp(12), alngOutput(3), 4&
                
        ' Copy decrypted bytes back into data array
        CopyMemory abytData(lngOffSet), abytTemp(0), 16&
        
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
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False
    Else
        DecryptData = True
    End If

    On Error GoTo 0     ' Nullify this error trap
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

Private Sub BlockEncrypt(ByRef alngInput() As Long, _
                         ByRef alngOutput() As Long)

    ' Called by EncryptData()
    
    Dim x0 As Long   ' Temp hold for input\output
    Dim x1 As Long
    Dim x2 As Long
    Dim x3 As Long
    
    Dim y0 As Long   ' Temp work variables
    Dim y1 As Long
    Dim y2 As Long
    Dim y3 As Long
    
    Dim lngIdx   As Long  ' Subkey index pointer
    Dim lngLoop  As Long  ' Loop counter
    Dim lngRound As Long  ' Encrypt rounds counter
    
    x0 = alngInput(0)   ' x0=plaintext[0];
    x1 = alngInput(1)   ' x1=plaintext[1];
    x2 = alngInput(2)   ' x2=plaintext[2];
    x3 = alngInput(3)   ' x3=plaintext[3];
    
    ' Added feature.
    ' Perform requested number of rounds.
    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit For..Next loop
        End If
    
        lngIdx = 0     ' Initialize counter
        
        For lngLoop = 1 To 4
        
            ' Keying(x0, x1, x2, x3, subkeys[ 0 to 30]);
            ' RND00 to 30(x0, x1, x2, x3, y0, y1, y2, y3);
            ' transform(y0, y1, y2, y3, x0, x1, x2, x3);
            Keying x0, x1, x2, x3, lngIdx
            RND00 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 1
            RND01 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 2
            RND02 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 3
            RND03 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 4
            RND04 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 5
            RND05 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            Keying x0, x1, x2, x3, lngIdx + 6
            RND06 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            If (lngIdx + 6) = 30 Then
                Exit For    ' exit For..Next loop
            End If
            
            Keying x0, x1, x2, x3, lngIdx + 7
            RND07 x0, x1, x2, x3, y0, y1, y2, y3
            Transform y0, y1, y2, y3, x0, x1, x2, x3
        
            lngIdx = lngIdx + 8
            
        Next lngLoop
        
        Keying x0, x1, x2, x3, 31              ' Keying(x0, x1, x2, x3, subkeys[31]);
        RND07 x0, x1, x2, x3, y0, y1, y2, y3   ' RND31(x0, x1, x2, x3, y0, y1, y2, y3);
        
        ' x0 = y0; x1 = y1; x2 = y2; x3 = y3;
        x0 = y0
        x1 = y1
        x2 = y2
        x3 = y3
        
        ' Keying(x0, x1, x2, x3, subkeys[32]);
        Keying x0, x1, x2, x3, 32
    
    Next lngRound

    ' Transfer encrypted data to output array
    alngOutput(0) = x0   ' ciphertext[0]=x0
    alngOutput(1) = x1   ' ciphertext[1]=x1
    alngOutput(2) = x2   ' ciphertext[2]=x2
    alngOutput(3) = x3   ' ciphertext[3]=x3

End Sub

Private Sub BlockDecrypt(ByRef alngInput() As Long, _
                         ByRef alngOutput() As Long)
    
    ' Called by DecryptData()
    
    Dim x0 As Long   ' Temp hold for input\output
    Dim x1 As Long
    Dim x2 As Long
    Dim x3 As Long
    
    Dim y0 As Long   ' Temp work variables
    Dim y1 As Long
    Dim y2 As Long
    Dim y3 As Long
    
    Dim lngIdx   As Long  ' Subkey index pointer
    Dim lngLoop  As Long  ' Loop counter
    Dim lngRound As Long  ' Decrypt rounds counter
    
    x0 = alngInput(0)   ' x0=ciphertext[0];
    x1 = alngInput(1)   ' x1=ciphertext[1];
    x2 = alngInput(2)   ' x2=ciphertext[2];
    x3 = alngInput(3)   ' x3=ciphertext[3];
    
    ' Added feature.
    ' Perform requested number of rounds.
    For lngRound = 1 To mlngRounds
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For   ' exit For..Next loop
        End If
    
        lngIdx = 30    ' Initialize counter

        Keying x0, x1, x2, x3, 32                 ' keying(x0, x1, x2, x3, subkeys[32]);
        InvRND07 x0, x1, x2, x3, y0, y1, y2, y3   ' InvRND31(x0, x1, x2, x3, y0, y1, y2, y3);
        Keying y0, y1, y2, y3, 31                 ' keying(y0, y1, y2, y3, subkeys[31]);
        
        For lngLoop = 1 To 4
            
            ' inv_transform(y0, y1, y2, y3, x0, x1, x2, x3);
            ' InvRND30 to 01(x0, x1, x2, x3, y0, y1, y2, y3);
            ' keying(y0, y1, y2, y3, subkeys[30 to 1]);
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND06 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx
        
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND05 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 1
        
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND04 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 2
            
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND03 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 3
            
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND02 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 4
        
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND01 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 5
        
            If (lngIdx - 5) = 1 Then
                Exit For    ' exit For..Next loop
            End If
            
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND00 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 6
        
            InvTransform y0, y1, y2, y3, x0, x1, x2, x3
            InvRND07 x0, x1, x2, x3, y0, y1, y2, y3
            Keying y0, y1, y2, y3, lngIdx - 7
        
            lngIdx = lngIdx - 8
            
        Next lngLoop
        
        ' inv_transform(y0, y1, y2, y3, x0, x1, x2, x3);
        ' InvRND00(x0, x1, x2, x3, y0, y1, y2, y3);
        InvTransform y0, y1, y2, y3, x0, x1, x2, x3
        InvRND00 x0, x1, x2, x3, y0, y1, y2, y3
        
        ' x0 = y0; x1 = y1; x2 = y2; x3 = y3;
        x0 = y0
        x1 = y1
        x2 = y2
        x3 = y3
            
        ' keying(x0, x1, x2, x3, subkeys[ 0]);
        Keying x0, x1, x2, x3, 0
    
    Next lngRound

    ' Transfer decrypted data to output array
    alngOutput(0) = x0   ' plaintext[0]=x0
    alngOutput(1) = x1   ' plaintext[1]=x1
    alngOutput(2) = x2   ' plaintext[2]=x2
    alngOutput(3) = x3   ' plaintext[3]=x3

End Sub

' #define keying(x0, x1, x2, x3, subkey) \
Private Sub Keying(ByRef x0 As Long, _
                   ByRef x1 As Long, _
                   ByRef x2 As Long, _
                   ByRef x3 As Long, _
                   ByVal lngIdx As Long)

    ' Called by BlockEncrypt()
    '           BlockDecrypt()
    
    x0 = x0 Xor malngSubKey(lngIdx, 0)   ' x0^=subkey[0]
    x1 = x1 Xor malngSubKey(lngIdx, 1)   ' x1^=subkey[1]
    x2 = x2 Xor malngSubKey(lngIdx, 2)   ' x2^=subkey[2]
    x3 = x3 Xor malngSubKey(lngIdx, 3)   ' x3^=subkey[3]

End Sub
                   
' linear transformation
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub Transform(ByVal x0 As Long, _
                      ByVal x1 As Long, _
                      ByVal x2 As Long, _
                      ByVal x3 As Long, _
                      ByRef y0 As Long, _
                      ByRef y1 As Long, _
                      ByRef y2 As Long, _
                      ByRef y3 As Long)
    
    ' Called by BlockEncrypt()
    
    With mobjMath32
        y0 = .w32Rotate(x0, 13)                   ' y0 = ROL(x0, 13)
        y2 = .w32Rotate(x2, 3)                    ' y2 = ROL(x2, 3)
        y1 = x1 Xor y0 Xor y2                     ' y1 = x1 ^ y0 ^ y2
        y3 = x3 Xor y2 Xor (.w32Shift(y0, 3))     ' y3 = x3 ^ y2 ^ ((unsigned long)y0<<3)
        y1 = .w32Rotate(y1, 1)                    ' y1 = ROL(y1, 1)
        y3 = .w32Rotate(y3, 7)                    ' y3 = ROL(y3, 7)
        y0 = y0 Xor y1 Xor y3                     ' y0 = y0 ^ y1 ^ y3
        y2 = y2 Xor y3 Xor (.w32Shift(y1, 7))     ' y2 = y2 ^ y3 ^ ((unsigned long)y1<<7)
        y0 = .w32Rotate(y0, 5)                    ' y0 = ROL(y0, 5)
        y2 = .w32Rotate(y2, 22)                   ' y2 = ROL(y2, 22)
    End With
    
End Sub

' Inverse of Linear Transformation
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvTransform(ByVal x0 As Long, _
                         ByVal x1 As Long, _
                         ByVal x2 As Long, _
                         ByVal x3 As Long, _
                         ByRef y0 As Long, _
                         ByRef y1 As Long, _
                         ByRef y2 As Long, _
                         ByRef y3 As Long)

    ' Called by BlockDecrypt()
    
    With mobjMath32
        y2 = .w32Rotate(x2, -22)                  ' y2 = ROR(x2, 22);\
        y0 = .w32Rotate(x0, -5)                   ' y0 = ROR(x0, 5)
        y2 = y2 Xor x3 Xor (.w32Shift(x1, 7))     ' y2 = y2 ^ x3 ^ ((unsigned long)x1<<7)
        y0 = y0 Xor x1 Xor x3                     ' y0 = y0 ^ x1 ^ x3
        y3 = .w32Rotate(x3, -7)                   ' y3 = ROR(x3, 7)
        y1 = .w32Rotate(x1, -1)                   ' y1 = ROR(x1, 1)
        y3 = y3 Xor y2 Xor (.w32Shift(y0, 3))     ' y3 = y3 ^ y2 ^ ((unsigned long)y0<<3)
        y1 = y1 Xor y0 Xor y2                     ' y1 = y1 ^ y0 ^ y2
        y2 = .w32Rotate(y2, -3)                   ' y2 = ROR(y2, 3)
        y0 = .w32Rotate(y0, -13)                  ' y0 = ROR(y0, 13)
    End With
    
End Sub

' RND00 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND00(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                     
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(18) As Long

    alngTmp(1) = x1 Xor x2                        ' t01 = x1  ^ x2
    alngTmp(2) = x0 Or x3                         ' t02 = x0  | x3
    alngTmp(3) = x0 Xor x1                        ' t03 = x0  ^ x1
    y3 = alngTmp(2) Xor alngTmp(1)                ' y3  = t02 ^ t01
    alngTmp(5) = x2 Or y3                         ' t05 = x2  | y3
    alngTmp(6) = x0 Xor x3                        ' t06 = x0  ^ x3
    alngTmp(7) = x1 Or x2                         ' t07 = x1  | x2
    alngTmp(8) = x3 And alngTmp(5)                ' t08 = x3  & t05
    alngTmp(9) = alngTmp(3) And alngTmp(7)        ' t09 = t03 & t07
    y2 = alngTmp(9) Xor alngTmp(8)                ' y2  = t09 ^ t08
    alngTmp(11) = alngTmp(9) And y2               ' t11 = t09 & y2
    alngTmp(12) = x2 Xor x3                       ' t12 = x2  ^ x3
    alngTmp(13) = alngTmp(7) Xor alngTmp(11)      ' t13 = t07 ^ t11
    alngTmp(14) = x1 And alngTmp(6)               ' t14 = x1  & t06
    alngTmp(15) = alngTmp(6) Xor alngTmp(13)      ' t15 = t06 ^ t13
    y0 = Not alngTmp(15)                          ' y0  =     ~ t15
    alngTmp(17) = y0 Xor alngTmp(14)              ' t17 = y0  ^ t14
    y1 = alngTmp(12) Xor alngTmp(17)              ' y1  = t12 ^ t17
                                                 
End Sub
                                                 
' Inverse of RND00
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND00(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(19) As Long
                                                 
    alngTmp(1) = x2 Xor x3                        ' t01 = x2  ^ x3
    alngTmp(2) = x0 Or x1                         ' t02 = x0  | x1
    alngTmp(3) = x1 Or x2                         ' t03 = x1  | x2
    alngTmp(4) = x2 And alngTmp(1)                ' t04 = x2  & t01
    alngTmp(5) = alngTmp(2) Xor alngTmp(1)        ' t05 = t02 ^ t01
    alngTmp(6) = x0 Or alngTmp(4)                 ' t06 = x0  | t04
    y2 = Not alngTmp(5)                           ' y2  =     ~ t05
    alngTmp(8) = x1 Xor x3                        ' t08 = x1  ^ x3
    alngTmp(9) = alngTmp(3) And alngTmp(8)        ' t09 = t03 & t08
    alngTmp(10) = x3 Or y2                        ' t10 = x3  | y2
    y1 = alngTmp(9) Xor alngTmp(6)                ' y1  = t09 ^ t06
    alngTmp(12) = x0 Or alngTmp(5)                ' t12 = x0  | t05
    alngTmp(13) = y1 Xor alngTmp(12)              ' t13 = y1  ^ t12
    alngTmp(14) = alngTmp(3) Xor alngTmp(10)      ' t14 = t03 ^ t10
    alngTmp(15) = x0 Xor x2                       ' t15 = x0  ^ x2
    y3 = alngTmp(14) Xor alngTmp(13)              ' y3  = t14 ^ t13
    alngTmp(17) = alngTmp(5) And alngTmp(13)      ' t17 = t05 & t13
    alngTmp(18) = alngTmp(14) Or alngTmp(17)      ' t18 = t14 | t17
    y0 = alngTmp(15) Xor alngTmp(18)              ' y0  = t15 ^ t18
                                                 
End Sub
                                                 
' RND01 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND01(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(18) As Long
                                                 
    alngTmp(1) = x0 Or x3                         ' t01 = x0  | x3
    alngTmp(2) = x2 Xor x3                        ' t02 = x2  ^ x3
    alngTmp(3) = Not x1                           ' t03 =     ~ x1
    alngTmp(4) = x0 Xor x2                        ' t04 = x0  ^ x2
    alngTmp(5) = x0 Or alngTmp(3)                 ' t05 = x0  | t03
    alngTmp(6) = x3 And alngTmp(4)                ' t06 = x3  & t04
    alngTmp(7) = alngTmp(1) And alngTmp(2)        ' t07 = t01 & t02
    alngTmp(8) = x1 Or alngTmp(6)                 ' t08 = x1  | t06
    y2 = alngTmp(2) Xor alngTmp(5)                ' y2  = t02 ^ t05
    alngTmp(10) = alngTmp(7) Xor alngTmp(8)       ' t10 = t07 ^ t08
    alngTmp(11) = alngTmp(1) Xor alngTmp(10)      ' t11 = t01 ^ t10
    alngTmp(12) = y2 Xor alngTmp(11)              ' t12 = y2  ^ t11
    alngTmp(13) = x1 And x3                       ' t13 = x1  & x3
    y3 = Not alngTmp(10)                          ' y3  =     ~ t10
    y1 = alngTmp(13) Xor alngTmp(12)              ' y1  = t13 ^ t12
    alngTmp(16) = alngTmp(10) Or y1               ' t16 = t10 | y1
    alngTmp(17) = alngTmp(5) And alngTmp(16)      ' t17 = t05 & t16
    y0 = x2 Xor alngTmp(17)                       ' y0  = x2  ^ t17
                                                 
End Sub
                                                 
' Inverse of RND01
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND01(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(18) As Long
                                                 
    alngTmp(1) = x0 Xor x1                        ' t01 = x0  ^ x1
    alngTmp(2) = x1 Or x3                         ' t02 = x1  | x3
    alngTmp(3) = x0 And x2                        ' t03 = x0  & x2
    alngTmp(4) = x2 Xor alngTmp(2)                ' t04 = x2  ^ t02
    alngTmp(5) = x0 Or alngTmp(4)                 ' t05 = x0  | t04
    alngTmp(6) = alngTmp(1) And alngTmp(5)        ' t06 = t01 & t05
    alngTmp(7) = x3 Or alngTmp(3)                 ' t07 = x3  | t03
    alngTmp(8) = x1 Xor alngTmp(6)                ' t08 = x1  ^ t06
    alngTmp(9) = alngTmp(7) Xor alngTmp(6)        ' t09 = t07 ^ t06
    alngTmp(10) = alngTmp(4) Or alngTmp(3)        ' t10 = t04 | t03
    alngTmp(11) = x3 And alngTmp(8)               ' t11 = x3  & t08
    y2 = Not alngTmp(9)                           ' y2  =     ~ t09
    y1 = alngTmp(10) Xor alngTmp(11)              ' y1  = t10 ^ t11
    alngTmp(14) = x0 Or y2                        ' t14 = x0  | y2
    alngTmp(15) = alngTmp(6) Xor y1               ' t15 = t06 ^ y1
    y3 = alngTmp(1) Xor alngTmp(4)                ' y3  = t01 ^ t04
    alngTmp(17) = x2 Xor alngTmp(15)              ' t17 = x2  ^ t15
    y0 = alngTmp(14) Xor alngTmp(17)              ' y0  = t14 ^ t17
                                                 
End Sub
                                                 
' RND02 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND02(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(15) As Long
                                                 
    alngTmp(1) = x0 Or x2                         ' t01 = x0  | x2
    alngTmp(2) = x0 Xor x1                        ' t02 = x0  ^ x1
    alngTmp(3) = x3 Xor alngTmp(1)                ' t03 = x3  ^ t01
    y0 = alngTmp(2) Xor alngTmp(3)                ' y0  = t02 ^ t03
    alngTmp(5) = x2 Xor y0                        ' t05 = x2  ^ y0
    alngTmp(6) = x1 Xor alngTmp(5)                ' t06 = x1  ^ t05
    alngTmp(7) = x1 Or alngTmp(5)                 ' t07 = x1  | t05
    alngTmp(8) = alngTmp(1) And alngTmp(6)        ' t08 = t01 & t06
    alngTmp(9) = alngTmp(3) Xor alngTmp(7)        ' t09 = t03 ^ t07
    alngTmp(10) = alngTmp(2) Or alngTmp(9)        ' t10 = t02 | t09
    y1 = alngTmp(10) Xor alngTmp(8)               ' y1  = t10 ^ t08
    alngTmp(12) = x0 Or x3                        ' t12 = x0  | x3
    alngTmp(13) = alngTmp(9) Xor y1               ' t13 = t09 ^ y1
    alngTmp(14) = x1 Xor alngTmp(13)              ' t14 = x1  ^ t13
    y3 = Not alngTmp(9)                           ' y3  =     ~ t09
    y2 = alngTmp(12) Xor alngTmp(14)              ' y2  = t12 ^ t14
                                                 
End Sub
                                                 
' Inverse of RND02
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND02(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(18) As Long
                                                 
    alngTmp(1) = x0 Xor x3                        ' t01 = x0  ^ x3
    alngTmp(2) = x2 Xor x3                        ' t02 = x2  ^ x3
    alngTmp(3) = x0 And x2                        ' t03 = x0  & x2
    alngTmp(4) = x1 Or alngTmp(2)                 ' t04 = x1  | t02
    y0 = alngTmp(1) Xor alngTmp(4)                ' y0  = t01 ^ t04
    alngTmp(6) = x0 Or x2                         ' t06 = x0  | x2
    alngTmp(7) = x3 Or y0                         ' t07 = x3  | y0
    alngTmp(8) = Not x3                           ' t08 =     ~ x3
    alngTmp(9) = x1 And alngTmp(6)                ' t09 = x1  & t06
    alngTmp(10) = alngTmp(8) Or alngTmp(3)        ' t10 = t08 | t03
    alngTmp(11) = x1 And alngTmp(7)               ' t11 = x1  & t07
    alngTmp(12) = alngTmp(6) And alngTmp(2)       ' t12 = t06 & t02
    y3 = alngTmp(9) Xor alngTmp(10)               ' y3  = t09 ^ t10
    y1 = alngTmp(12) Xor alngTmp(11)              ' y1  = t12 ^ t11
    alngTmp(15) = x2 And y3                       ' t15 = x2  & y3
    alngTmp(16) = y0 Xor y1                       ' t16 = y0  ^ y1
    alngTmp(17) = alngTmp(10) Xor alngTmp(15)     ' t17 = t10 ^ t15
    y2 = alngTmp(16) Xor alngTmp(17)              ' y2  = t16 ^ t17
                                                 
End Sub
                                                 
' RND03 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND03(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(16) As Long
                                                 
    alngTmp(1) = x0 Xor x2                        ' t01 = x0  ^ x2
    alngTmp(2) = x0 Or x3                         ' t02 = x0  | x3
    alngTmp(3) = x0 And x3                        ' t03 = x0  & x3
    alngTmp(4) = alngTmp(1) And alngTmp(2)        ' t04 = t01 & t02
    alngTmp(5) = x1 Or alngTmp(3)                 ' t05 = x1  | t03
    alngTmp(6) = x0 And x1                        ' t06 = x0  & x1
    alngTmp(7) = x3 Xor alngTmp(4)                ' t07 = x3  ^ t04
    alngTmp(8) = x2 Or alngTmp(6)                 ' t08 = x2  | t06
    alngTmp(9) = x1 Xor alngTmp(7)                ' t09 = x1  ^ t07
    alngTmp(10) = x3 And alngTmp(5)               ' t10 = x3  & t05
    alngTmp(11) = alngTmp(2) Xor alngTmp(10)      ' t11 = t02 ^ t10
    y3 = alngTmp(8) Xor alngTmp(9)                ' y3  = t08 ^ t09
    alngTmp(13) = x3 Or y3                        ' t13 = x3  | y3
    alngTmp(14) = x0 Or alngTmp(7)                ' t14 = x0  | t07
    alngTmp(15) = x1 And alngTmp(13)              ' t15 = x1  & t13
    y2 = alngTmp(8) Xor alngTmp(11)               ' y2  = t08 ^ t11
    y0 = alngTmp(14) Xor alngTmp(15)              ' y0  = t14 ^ t15
    y1 = alngTmp(5) Xor alngTmp(4)                ' y1  = t05 ^ t04
                                                 
End Sub
                                                 
' Inverse of RND03
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND03(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(17) As Long
                                                 
    alngTmp(1) = x2 Or x3                         ' t01 = x2  | x3
    alngTmp(2) = x0 Or x3                         ' t02 = x0  | x3
    alngTmp(3) = x2 Xor alngTmp(2)                ' t03 = x2  ^ t02
    alngTmp(4) = x1 Xor alngTmp(2)                ' t04 = x1  ^ t02
    alngTmp(5) = x0 Xor x3                        ' t05 = x0  ^ x3
    alngTmp(6) = alngTmp(4) And alngTmp(3)        ' t06 = t04 & t03
    alngTmp(7) = x1 And alngTmp(1)                ' t07 = x1  & t01
    y2 = alngTmp(5) Xor alngTmp(6)                ' y2  = t05 ^ t06
    alngTmp(9) = x0 Xor alngTmp(3)                ' t09 = x0  ^ t03
    y0 = alngTmp(7) Xor alngTmp(3)                ' y0  = t07 ^ t03
    alngTmp(11) = y0 Or alngTmp(5)                ' t11 = y0  | t05
    alngTmp(12) = alngTmp(9) And alngTmp(11)      ' t12 = t09 & t11
    alngTmp(13) = x0 And y2                       ' t13 = x0  & y2
    alngTmp(14) = alngTmp(1) Xor alngTmp(5)       ' t14 = t01 ^ t05
    y1 = x1 Xor alngTmp(12)                       ' y1  = x1  ^ t12
    alngTmp(16) = x1 Or alngTmp(13)               ' t16 = x1  | t13
    y3 = alngTmp(14) Xor alngTmp(16)              ' y3  = t14 ^ t16
                                                 
End Sub
                                                 
' RND04 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND04(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(17) As Long
                                                 
    alngTmp(1) = x0 Or x1                         ' t01 = x0  | x1
    alngTmp(2) = x1 Or x2                         ' t02 = x1  | x2
    alngTmp(3) = x0 Xor alngTmp(2)                ' t03 = x0  ^ t02
    alngTmp(4) = x1 Xor x3                        ' t04 = x1  ^ x3
    alngTmp(5) = x3 Or alngTmp(3)                 ' t05 = x3  | t03
    alngTmp(6) = x3 And alngTmp(1)                ' t06 = x3  & t01
    y3 = alngTmp(3) Xor alngTmp(6)                ' y3  = t03 ^ t06
    alngTmp(8) = y3 And alngTmp(4)                ' t08 = y3  & t04
    alngTmp(9) = alngTmp(4) And alngTmp(5)        ' t09 = t04 & t05
    alngTmp(10) = x2 Xor alngTmp(6)               ' t10 = x2  ^ t06
    alngTmp(11) = x1 And x2                       ' t11 = x1  & x2
    alngTmp(12) = alngTmp(4) Xor alngTmp(8)       ' t12 = t04 ^ t08
    alngTmp(13) = alngTmp(11) Or alngTmp(3)       ' t13 = t11 | t03
    alngTmp(14) = alngTmp(10) Xor alngTmp(9)      ' t14 = t10 ^ t09
    alngTmp(15) = x0 And alngTmp(5)               ' t15 = x0  & t05
    alngTmp(16) = alngTmp(11) Or alngTmp(12)      ' t16 = t11 | t12
    y2 = alngTmp(13) Xor alngTmp(8)               ' y2  = t13 ^ t08
    y1 = alngTmp(15) Xor alngTmp(16)              ' y1  = t15 ^ t16
    y0 = Not alngTmp(14)                          ' y0  =     ~ t14
                                                 
End Sub
                                                 
' Inverse of RND04
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND04(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(16) As Long
                                                 
    alngTmp(1) = x1 Or x3                         ' t01 = x1  | x3
    alngTmp(2) = x2 Or x3                         ' t02 = x2  | x3
    alngTmp(3) = x0 And alngTmp(1)                ' t03 = x0  & t01
    alngTmp(4) = x1 Xor alngTmp(2)                ' t04 = x1  ^ t02
    alngTmp(5) = x2 Xor x3                        ' t05 = x2  ^ x3
    alngTmp(6) = Not alngTmp(3)                   ' t06 =     ~ t03
    alngTmp(7) = x0 And alngTmp(4)                ' t07 = x0  & t04
    y1 = alngTmp(5) Xor alngTmp(7)                ' y1  = t05 ^ t07
    alngTmp(9) = y1 Or alngTmp(6)                 ' t09 = y1  | t06
    alngTmp(10) = x0 Xor alngTmp(7)               ' t10 = x0  ^ t07
    alngTmp(11) = alngTmp(1) Xor alngTmp(9)       ' t11 = t01 ^ t09
    alngTmp(12) = x3 Xor alngTmp(4)               ' t12 = x3  ^ t04
    alngTmp(13) = x2 Or alngTmp(10)               ' t13 = x2  | t10
    y3 = alngTmp(3) Xor alngTmp(12)               ' y3  = t03 ^ t12
    alngTmp(15) = x0 Xor alngTmp(4)               ' t15 = x0  ^ t04
    y2 = alngTmp(11) Xor alngTmp(13)              ' y2  = t11 ^ t13
    y0 = alngTmp(15) Xor alngTmp(9)               ' y0  = t15 ^ t09
                                                 
End Sub
                                                 
' RND05 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND05(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(15) As Long
                                                 
    alngTmp(1) = x1 Xor x3                        ' t01 = x1  ^ x3
    alngTmp(2) = x1 Or x3                         ' t02 = x1  | x3
    alngTmp(3) = x0 And alngTmp(1)                ' t03 = x0  & t01
    alngTmp(4) = x2 Xor alngTmp(2)                ' t04 = x2  ^ t02
    alngTmp(5) = alngTmp(3) Xor alngTmp(4)        ' t05 = t03 ^ t04
    y0 = Not alngTmp(5)                           ' y0  =     ~ t05
    alngTmp(7) = x0 Xor alngTmp(1)                ' t07 = x0  ^ t01
    alngTmp(8) = x3 Or y0                         ' t08 = x3  | y0
    alngTmp(9) = x1 Or alngTmp(5)                 ' t09 = x1  | t05
    alngTmp(10) = x3 Xor alngTmp(8)               ' t10 = x3  ^ t08
    alngTmp(11) = x1 Or alngTmp(7)                ' t11 = x1  | t07
    alngTmp(12) = alngTmp(3) Or y0                ' t12 = t03 | y0
    alngTmp(13) = alngTmp(7) Or alngTmp(10)       ' t13 = t07 | t10
    alngTmp(14) = alngTmp(1) Xor alngTmp(11)      ' t14 = t01 ^ t11
    y2 = alngTmp(9) Xor alngTmp(13)               ' y2  = t09 ^ t13
    y1 = alngTmp(7) Xor alngTmp(8)                ' y1  = t07 ^ t08
    y3 = alngTmp(12) Xor alngTmp(14)              ' y3  = t12 ^ t14
                                                 
End Sub
                                                 
' Inverse of RND05
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND05(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(17) As Long
                                                 
    alngTmp(1) = x0 And x3                        ' t01 = x0  & x3
    alngTmp(2) = x2 Xor alngTmp(1)                ' t02 = x2  ^ t01
    alngTmp(3) = x0 Xor x3                        ' t03 = x0  ^ x3
    alngTmp(4) = x1 And alngTmp(2)                ' t04 = x1  & t02
    alngTmp(5) = x0 And x2                        ' t05 = x0  & x2
    y0 = alngTmp(3) Xor alngTmp(4)                ' y0  = t03 ^ t04
    alngTmp(7) = x0 And y0                        ' t07 = x0  & y0
    alngTmp(8) = alngTmp(1) Xor y0                ' t08 = t01 ^ y0
    alngTmp(9) = x1 Or alngTmp(5)                 ' t09 = x1  | t05
    alngTmp(10) = Not x1                          ' t10 =     ~ x1
    y1 = alngTmp(8) Xor alngTmp(9)                ' y1  = t08 ^ t09
    alngTmp(12) = alngTmp(10) Or alngTmp(7)       ' t12 = t10 | t07
    alngTmp(13) = y0 Or y1                        ' t13 = y0  | y1
    y3 = alngTmp(2) Xor alngTmp(12)               ' y3  = t02 ^ t12
    alngTmp(15) = alngTmp(2) Xor alngTmp(13)      ' t15 = t02 ^ t13
    alngTmp(16) = x1 Xor x3                       ' t16 = x1  ^ x3
    y2 = alngTmp(16) Xor alngTmp(15)              ' y2  = t16 ^ t15
                                                 
End Sub
                                                 
' RND06 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND06(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(19) As Long
                                                 
    alngTmp(1) = x0 And x3                        ' t01 = x0  & x3
    alngTmp(2) = x1 Xor x2                        ' t02 = x1  ^ x2
    alngTmp(3) = x0 Xor x3                        ' t03 = x0  ^ x3
    alngTmp(4) = alngTmp(1) Xor alngTmp(2)        ' t04 = t01 ^ t02
    alngTmp(5) = x1 Or x2                         ' t05 = x1  | x2
    y1 = Not alngTmp(4)                           ' y1  =     ~ t04
    alngTmp(7) = alngTmp(3) And alngTmp(5)        ' t07 = t03 & t05
    alngTmp(8) = x1 And y1                        ' t08 = x1  & y1
    alngTmp(9) = x0 Or x2                         ' t09 = x0  | x2
    alngTmp(10) = alngTmp(7) Xor alngTmp(8)       ' t10 = t07 ^ t08
    alngTmp(11) = x1 Or x3                        ' t11 = x1  | x3
    alngTmp(12) = x2 Xor alngTmp(11)              ' t12 = x2  ^ t11
    alngTmp(13) = alngTmp(9) Xor alngTmp(10)      ' t13 = t09 ^ t10
    y2 = Not alngTmp(13)                          ' y2  =     ~ t13
    alngTmp(15) = y1 And alngTmp(3)               ' t15 = y1  & t03
    y3 = alngTmp(12) Xor alngTmp(7)               ' y3  = t12 ^ t07
    alngTmp(17) = x0 Xor x1                       ' t17 = x0  ^ x1
    alngTmp(18) = y2 Xor alngTmp(15)              ' t18 = y2  ^ t15
    y0 = alngTmp(17) Xor alngTmp(18)              ' y0  = t17 ^ t18
                                                 
End Sub
                                                 
' Inverse of RND06
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND06(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(18) As Long
                                                 
    alngTmp(1) = x0 Xor x2                        ' t01 = x0  ^ x2
    alngTmp(2) = Not x2                           ' t02 =     ~ x2
    alngTmp(3) = x1 And alngTmp(1)                ' t03 = x1  & t01
    alngTmp(4) = x1 Or alngTmp(2)                 ' t04 = x1  | t02
    alngTmp(5) = x3 Or alngTmp(3)                 ' t05 = x3  | t03
    alngTmp(6) = x1 Xor x3                        ' t06 = x1  ^ x3
    alngTmp(7) = x0 And alngTmp(4)                ' t07 = x0  & t04
    alngTmp(8) = x0 Or alngTmp(2)                 ' t08 = x0  | t02
    alngTmp(9) = alngTmp(7) Xor alngTmp(5)        ' t09 = t07 ^ t05
    y1 = alngTmp(6) Xor alngTmp(8)                ' y1  = t06 ^ t08
    y0 = Not alngTmp(9)                           ' y0  =     ~ t09
    alngTmp(12) = x1 And y0                       ' t12 = x1  & y0
    alngTmp(13) = alngTmp(1) And alngTmp(5)       ' t13 = t01 & t05
    alngTmp(14) = alngTmp(1) Xor alngTmp(12)      ' t14 = t01 ^ t12
    alngTmp(15) = alngTmp(7) Xor alngTmp(13)      ' t15 = t07 ^ t13
    alngTmp(16) = x3 Or alngTmp(2)                ' t16 = x3  | t02
    alngTmp(17) = x0 Xor y1                       ' t17 = x0  ^ y1
    y3 = alngTmp(17) Xor alngTmp(15)              ' y3  = t17 ^ t15
    y2 = alngTmp(16) Xor alngTmp(14)              ' y2  = t16 ^ t14
                                                 
End Sub
                                                 
' RND07 bit mixing
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub RND07(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                  ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by EvaluateKey()
    '           BlockEncrypt()
    
    Dim alngTmp(18) As Long
                                                 
    alngTmp(1) = x0 And x2                        ' t01 = x0  & x2
    alngTmp(2) = Not x3                           ' t02 =     ~ x3
    alngTmp(3) = x0 And alngTmp(2)                ' t03 = x0  & t02
    alngTmp(4) = x1 Or alngTmp(1)                 ' t04 = x1  | t01
    alngTmp(5) = x0 And x1                        ' t05 = x0  & x1
    alngTmp(6) = x2 Xor alngTmp(4)                ' t06 = x2  ^ t04
    y3 = alngTmp(3) Xor alngTmp(6)                ' y3  = t03 ^ t06
    alngTmp(8) = x2 Or y3                         ' t08 = x2  | y3
    alngTmp(9) = x3 Or alngTmp(5)                 ' t09 = x3  | t05
    alngTmp(10) = x0 Xor alngTmp(8)               ' t10 = x0  ^ t08
    alngTmp(11) = alngTmp(4) And y3               ' t11 = t04 & y3
    y1 = alngTmp(9) Xor alngTmp(10)               ' y1  = t09 ^ t10
    alngTmp(13) = x1 Xor y1                       ' t13 = x1  ^ y1
    alngTmp(14) = alngTmp(1) Xor y1               ' t14 = t01 ^ y1
    alngTmp(15) = x2 Xor alngTmp(5)               ' t15 = x2  ^ t05
    alngTmp(16) = alngTmp(11) Or alngTmp(13)      ' t16 = t11 | t13
    alngTmp(17) = alngTmp(2) Or alngTmp(14)       ' t17 = t02 | t14
    y0 = alngTmp(15) Xor alngTmp(17)              ' y0  = t15 ^ t17
    y2 = x0 Xor alngTmp(16)                       ' y2  = x0  ^ t16
                                                 
End Sub
                                                 
' Inverse of RND07
' Input=x0,x1,x2,x3   output=y0,y1,y2,y3
Private Sub InvRND07(ByVal x0 As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal x3 As Long, _
                     ByRef y0 As Long, ByRef y1 As Long, ByRef y2 As Long, ByRef y3 As Long)
                                                 
    ' Called by BlockDecrypt()
    
    Dim alngTmp(17) As Long
                                                 
    alngTmp(1) = x0 And x1                        ' t01 = x0  & x1
    alngTmp(2) = x0 Or x1                         ' t02 = x0  | x1
    alngTmp(3) = x2 Or alngTmp(1)                 ' t03 = x2  | t01
    alngTmp(4) = x3 And alngTmp(2)                ' t04 = x3  & t02
    y3 = alngTmp(3) Xor alngTmp(4)                ' y3  = t03 ^ t04
    alngTmp(6) = x1 Xor alngTmp(4)                ' t06 = x1  ^ t04
    alngTmp(7) = x3 Xor y3                        ' t07 = x3  ^ y3
    alngTmp(8) = Not alngTmp(7)                   ' t08 =     ~ t07
    alngTmp(9) = alngTmp(6) Or alngTmp(8)         ' t09 = t06 | t08
    alngTmp(10) = x1 Xor x3                       ' t10 = x1  ^ x3
    alngTmp(11) = x0 Or x3                        ' t11 = x0  | x3
    y1 = x0 Xor alngTmp(9)                        ' y1  = x0  ^ t09
    alngTmp(13) = x2 Xor alngTmp(6)               ' t13 = x2  ^ t06
    alngTmp(14) = x2 And alngTmp(11)              ' t14 = x2  & t11
    alngTmp(15) = x3 Or y1                        ' t15 = x3  | y1
    alngTmp(16) = alngTmp(1) Or alngTmp(10)       ' t16 = t01 | t10
    y0 = alngTmp(13) Xor alngTmp(15)              ' y0  = t13 ^ t15
    y2 = alngTmp(14) Xor alngTmp(16)              ' y2  = t14 ^ t16
                                              
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

    ' Called by Serpent_File()
    '           Serpent_String()
    
    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select
    
    RaiseEvent CurrentBlockSize(GetBlockSize)
    
End Function

Private Sub ResetVariables()

    Erase malngSubKey()
    Erase mabytPassword()

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
    
    ResetVariables                 ' Empty arrays
    Set mobjMath32 = Nothing       ' Verify class object is not active
    Set mobjMath32 = New cMath32   ' Instantiate class object
    
    ' Preset property values
    KeyLength = 256        ' Default value
    EncryptRounds = 1      ' Number of rounds of encryption
    PrimaryKeyRounds = 1   ' Number of rounds to mix primary key array
    CreateNewFile = True   ' Default is to create a new output file

End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    ResetVariables             ' Empty arrays
    Set mobjMath32 = Nothing   ' Free class object from memory

End Sub
