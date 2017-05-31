VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSkipjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cSkipjack  (clsSkipjack.cls)
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
' Description:   Skipjack is a recent algorithm that was developed in 1987
'                and put into service in 1993.  It is a formerly secret NSA
'                (National Security Agency) encryption algorithm that was
'                declassified on June 24, 1998.  Skipjack is a
'                representative of a family of encryption algorithms
'                developed in 1980 as  part of the NSA suite of "Type I"
'                algorithms, which are  suitable for protecting all levels
'                of classified data.   Type I algorithms are typically
'                extremely secure and are  usually classified as secret.
'                Skipjack was used to encrypt  sensitive, but not
'                classified, government data. It was  implemented in two
'                government encryption devices:  the Clipper chip and
'                Fortezza PC card.  These devices have many uses and are
'                widely employed by agencies such as the FBI and NSA. They
'                provide a high level of security for sensitive
'                communications while enabling the interception of
'                telecommunications by law  enforcement officials for such
'                things as criminal  investigations.  For example, Clipper
'                chips can be used to  provide secure telephone
'                transmissions and Fortezza cards can  be used to encrypt
'                such things as e-mail and network traffic.  The key
'                characteristic of both devices is that they were designed
'                with "back doors" that allow government agents to monitor
'                encrypted transmissions given the proper authority.  This
'                feature is covered under the Escrowed Encryption Standard,
'                of which Skipjack is a part. It is implemented via a
'                mechanism called a LEAF (Law Enforcement Access Field). It
'                is important to note that Skipjack itself has nothing to
'                do with this "back door" functionality!  Skipjack is
'                entirely separate from the LEAF in the Clipper and
'                Fortezza products and is not affected in any way by its
'                presence.
'
'                Skipjack has been extensively cryptanalyzed, and has no
'                weaknesses. There are no known shortcut attacks that can
'                break Skipjack. However, the small key size makes this
'                algorithm inferior to the newer candidate algorithms for
'                the AES (Advanced Encryption Standard) competition being
'                held by NIST (National Institute of Standards and
'                Technology).  The NIST is a federal technology agency that
'                develops and promotes measurement, standards, and
'                technology.  Despite this shortcoming, Skipjack still
'                provides very strong security and it should be many years
'                before the algorithm is broken by a brute force attack.
'                Note that it was declassified in order to provide a
'                software implementation of Fortezza enabled applications.
'                The NSA does not intend  for the algorithm to be a
'                candidate for the AES.  Like  Triple DES, Skipjack is an
'                interim solution to be used until  the final AES is
'                completed and widely implemented. It offers  a safe
'                alternative to DES without having to rely on the AES.
'
'                In 1993 an Interim Report was released that gave a
'                thorough analysis of Skipjack and discussed issues
'                relating to Skipjack's integration in the Clipper and
'                Fortezza hardware. Although this report is many years old,
'                it provides some interesting information and a good
'                perspective on what it would take to break Skipjack with a
'                brute force attack. While reading the document, keep in
'                mind that the algorithm was still a closely guarded secret
'                at the time the report was written. The Final Report
'                mentioned in this document was never written, so this
'                document is the closest thing there is to an official
'                analysis of the Skipjack algorithm.
'
'                Skipjack encrypts and decrypts data in 64-bit blocks,
'                normally using an 80-bit key. It takes a 64-bit block of
'                plaintext as input and outputs a 64-bit block of
'                ciphertext. Skipjack has 32 rounds, meaning the main
'                algorithm is repeated 32 times to produce the ciphertext.
'                It has been found that the number of rounds is
'                exponentially proportional to the amount of time required
'                to find a key using a brute-force attack. So as the number
'                of rounds increases, the security of the algorithm
'                increases exponentially.
'
'                The detailed description of the actual algorithm is
'                contained in the official Skipjack specification provided
'                by the NSA after the algorithm was declassified.  Another
'                algorithm called KEA was declassified at the same time As
'                Skipjack, so the specification contains information about
'                both Skipjack and KEA. The first half of the document
'                concerns Skipjack and the second half focuses on KEA (Key
'                Exchange Algorithm).
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

'                - F-table loading is more dynamic because it is now based
'                  on key length and not a single table.  See LoadWorkArrays()
'                  routine.
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
'                sixteen bytes of salt (random) data.  After decrypting each
'                block do not forget to drop the first sixteen bytes.
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
' References:    Skipjack documentation above
'                http://csrc.nist.gov/encryption/skipjack/skipjack.pdf
'
'                Tropical software (Portions of description above)
'                http://www.tropsoft.com/strongenc/skipjack.htm
'
'                Observations on the SkipJack Encryption Algorithm
'                http://www.cs.technion.ac.il/~biham/Reports/SkipJack/
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
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Modified Skipjack_File() and Skipjack_String() routines to
'                encrypt appended data information for additional security.
'              - Moved ExpandData() routine to basCommon.bas module.
'              - Renamed enumeration SKIPJACK_ROUNDS and used standard values
'                of 0-4 to represent number of rounds of encryption.
' 06-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Modified EvaluateKey() routine for initial table loading.
'              - Updated property EncryptRounds().  Number of encryption
'                rounds is now user-defined.  See constant MAX_ROUNDS.
'              - Rewrote Encrytion/Decryption routines.  I have found several
'                C versions of Skipjack on the web.  This interpretation is
'                based on the version by Paulo Barreto with some modifications
'                by myself because the C language can manipulate numeric values
'                that are not easily done within the realm of Visual Basic 6.0.
'
'                Originally written by Panu Rissanen <bande@lut.fi> 1998.06.24
'                Optimized by Mark Tillotson <markt@chaos.org.uk> 1998.06.25
'                Optimized by Paulo Barreto <pbarreto@nw.com.br> 1998.06.30
'
'                Home page:  http://www.larc.usp.br/~pbarreto/
'                            http://www.larc.usp.br/~pbarreto/skipjack32.zip
'
' 26-Dec-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added property KeyLength() to meet and exceed AES standards.
'              (128, 192, 256, 384, 512, 1024 bits)  Modified EvaluateKey()
'              routine to use utilize the new key lengths.
' 01-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed OffSet range in EncryptData() and DecryptData() routines.
'              - Moved GetBlockSize() routine to common module.
'              - Set key length to increments of 32 bits for 128 to 448 and
'                increments of 64 bits for 448 to 1024.
' 08-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Set data block size to 16 bytes.
'              - Updated documentation.
' 15-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added additional mixing of data within maintTable() array in
'                EvaluateKey() routine.
'              - Added data swapping to EncryptData() and DecryptData()
'                routines for enhanced security.
' 21-Jan-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Removed obsolete code. Updated documentation.
'              - Correct data size capture for decryption in Skipjack_String().
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Enhanced initial loading of key array in EvaluateKey() routine.
'              - Updated LoadWorkArrays() routine with a good alternative for
'                loading the table data.
' 29-Aug-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated initial loading of key array in EvaluateKey() routine.
' 01-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced Err.Raise with InfoMsg so programs will continue to
'              flow smoothly.
' 29-Jan-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added ability to see progress display while encrypting or
'                decrypting data (EncryptData() and DecryptData() routines).
'              - Removed obsolete code.  Updated documentation.
'              - Updated initial key loading (EvaluateKey() routine).
' 02-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Replaced ShiftInt() routine with call to cMath32 module.
'              - Updated EvaluateKey() routine.
' 10-Apr-2011  Kenneth Ives  kenaso@tx.rr.com
'              Removed MixBytes() routine.  I created this routine and now
'              feel that it is no longer needed.
' 20-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated LoadWorkArrays() routine
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Made many updates as per Joe Sova's suggestions
'                1. Removed most of the erase array as they were not needed
'                2. Moved call to display progress to outside processing loop
'                   in EncryptData() and DecryptData() routines.
'                3. Moved Decryption size testing to earlier position within
'                   within Skipjack_File() routine.
'                4. Split size testing in Skipjack_String() routine.
' 21-Feb-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Added Public Event CurrentBlockSize() to track current
'                amount of data being processed.
'              - Updated GetBlockSize() routine.
'              - Increased area for number of rounds that will affect
'                multiple pass functionality.  See EncryptData() and
'                DecryptData() routines.
'              - Added new Property CreateNewFile().  Designates if creating
'                a new output file or overwriting the source file after
'                encryption or decryption.  Default value is TRUE.
'              - Updated Skipjack_File() routine to reference new property.
' 06-Jun-2012  Kenneth Ives  kenaso@tx.rr.com
'              Updated EncryptData() and DecryptData() routines to encompass
'              complete process per round if more than one iteration is
'              selected by user.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME   As String = "clsSkipjack"
  Private Const MAX_FILE_SIZE As Currency = 10000000000@   ' Approx 10 Gig
  Private Const MB_5          As Long = &H500000           ' 5242880 bytes
  Private Const BLOCK_SIZE    As Long = 8
  Private Const DATA_BLOCK    As Long = 16
  Private Const MAX_ROUNDS    As Long = 10    ' Modify as you see fit
  Private Const MAX_BYTE      As Long = 256   ' Max range of byte (0-255)

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
  Private mlngKeyLength     As Long
  Private maintTable()      As Integer
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
    '             X.Password = abytPwd()  <-- Highly recommend using a byte array
    
    Dim intIndex As Integer
    
    Const ROUTINE_NAME As String = "Password Property"
    
    On Error GoTo Password_Error

    Erase mabytPassword()
    
    ' Is password data in an array
    If IsArray(vntData) Then
    
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
' Routine:       Skipjack_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)  ' convert file location to byte array 
'                objSkipjack.Skipjack_File abytFile(), blnEncrypt             ' Encrypt/Decrypt data
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
' 15-Mar-2009  Kenneth Ives  kenaso@tx.rr.com
'              Updated EvaluateKey() and LoadWorkArrays() routines.
' 20-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Moved Decryption size testing to an earlier position as per
'              Joe Sova's suggestion.
' ***************************************************************************
Public Function Skipjack_File(ByRef abytFile() As Byte, _
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

    Const ROUTINE_NAME As String = "Skipjack_File"
    
    On Error GoTo Skipjack_File_Error
    
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
        GoTo Skipjack_File_CleanUp
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
        GoTo Skipjack_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file
        
        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
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
                GoTo Skipjack_File_CleanUp
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
                GoTo Skipjack_File_CleanUp
                
            End If
        End If
        
        mcurMaxSize = mcurOrigLength * mlngRounds   ' Save file size for progress calcs
        
        mcurAccrued = 0@   ' Init accrued progress
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@
        
        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
        End If
        
        strTempFile = CreateTempFile()  ' Create a temp target file
        
        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Skipjack_File_CleanUp
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
            GoTo Skipjack_File_CleanUp
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
                GoTo Skipjack_File_CleanUp
                
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

Skipjack_File_CleanUp:
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
        Skipjack_File = False
    Else
        Skipjack_File = True
    End If
        
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Skipjack_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Skipjack_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Skipjack_String
'
' Description:   Function to encrypt/decrypt a string of data.
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)          ' convert string data to byte array
'                objSkipjack.Skipjack_String abytData(), blnEncrypt  ' Encrypt/Decrypt data & return as byte array
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
Public Function Skipjack_String(ByRef abytData() As Byte, _
                       Optional ByVal blnEncrypt As Boolean = True) As Boolean

    Dim curDataLength As Currency
    Dim abytSize(7)   As Byte
    
    Const ROUTINE_NAME As String = "Skipjack_String"

    On Error GoTo Skipjack_String_Error
    
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
        GoTo Skipjack_String_CleanUp
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
    
Skipjack_String_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Skipjack_String = False
    Else
        Skipjack_String = True
    End If
    
    ResetVariables    ' Reset class variables
    Erase abytSize()  ' Always empty arrays when not needed
    On Error GoTo 0   ' Nullify this error trap
    Exit Function
    
Skipjack_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Skipjack_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function EvaluateKey() As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngIdx      As Long
    Dim lngIndex    As Long
    Dim lngKeyLen   As Long
    Dim lngTemp     As Long
    Dim abytPwd()   As Byte
    Dim abytTable() As Byte

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
    
    Erase abytPwd()    ' Always start with empty arrays
    Erase abytTable()
    
    lngKeyLen = mlngKeyLength \ 8                        ' Ex:  16 bytes = 128 bits \ 8 bits
    abytPwd() = ExpandData(mabytPassword(), lngKeyLen)   ' Expand password key to desired length
    abytTable() = LoadWorkArrays()                       ' Load F-table data (ASCII 0-255 mixed)
    
    ReDim maintTable(0 To 9, MAX_BYTE)   ' Size integer table array
    
    ' for (i = 0; i < 10; i++)
    '     byte *t = tab[i], k = key[i];
    '     for (c = 0; c < 256; c++)
    '         t[c] = fTable[c ^ k];
    For lngIndex = 0 To 9
        
        lngTemp = CLng(abytPwd(lngIndex))   ' Use first 10 values of password array
        
        ' Load integer table data
        For lngIdx = 0 To (MAX_BYTE - 1)
            maintTable(lngIndex, lngIdx) = CInt(abytTable((lngIdx Xor lngTemp) Mod MAX_BYTE))
        Next lngIdx
    Next lngIndex
    
EvaluateKey_CleanUp:
    Erase abytPwd()    ' Always empty arrays when not needed
    Erase abytTable()
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EvaluateKey = False
    Else
        EvaluateKey = True
    End If
    
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

EvaluateKey_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ResetVariables
    gblnStopProcessing = True
    Resume EvaluateKey_CleanUp

End Function

Private Function EncryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngRound      As Long
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim intW1         As Integer
    Dim intW2         As Integer
    Dim intW3         As Integer
    Dim intW4         As Integer
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte
    
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
          
        For lngRound = 1 To mlngRounds

            ' Transfer eight bytes to work array
            CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                    
            ' Convert eight bytes to four integers
            CopyMemory intW1, abytInput(0), 2&
            CopyMemory intW2, abytInput(2), 2&
            CopyMemory intW3, abytInput(4), 2&
            CopyMemory intW4, abytInput(6), 2&
                
            ' Rule A1
            '    /* stepping rule A: */
            '    g0(tab, w1); w4 ^= w1 ^ 1;
            '    g1(tab, w4); w3 ^= w4 ^ 2;
            '    g2(tab, w3); w2 ^= w3 ^ 3;
            '    g3(tab, w2); w1 ^= w2 ^ 4;
            '    g4(tab, w1); w4 ^= w1 ^ 5;
            '    g0(tab, w4); w3 ^= w4 ^ 6;
            '    g1(tab, w3); w2 ^= w3 ^ 7;
            '    g2(tab, w2); w1 ^= w2 ^ 8;
            GG intW1, 0, 1, 2, 3: intW4 = intW4 Xor (intW1 Xor 1)
            GG intW4, 4, 5, 6, 7: intW3 = intW3 Xor (intW4 Xor 2)
            GG intW3, 8, 9, 0, 1: intW2 = intW2 Xor (intW3 Xor 3)
            GG intW2, 2, 3, 4, 5: intW1 = intW1 Xor (intW2 Xor 4)
            GG intW1, 6, 7, 8, 9: intW4 = intW4 Xor (intW1 Xor 5)
            GG intW4, 0, 1, 2, 3: intW3 = intW3 Xor (intW4 Xor 6)
            GG intW3, 4, 5, 6, 7: intW2 = intW2 Xor (intW3 Xor 7)
            GG intW2, 8, 9, 0, 1: intW1 = intW1 Xor (intW2 Xor 8)

            ' Rule B1
            '    /* stepping rule B: */
            '    w2 ^= w1 ^  9; g3(tab, w1);
            '    w1 ^= w4 ^ 10; g4(tab, w4);
            '    w4 ^= w3 ^ 11; g0(tab, w3);
            '    w3 ^= w2 ^ 12; g1(tab, w2);
            '    w2 ^= w1 ^ 13; g2(tab, w1);
            '    w1 ^= w4 ^ 14; g3(tab, w4);
            '    w4 ^= w3 ^ 15; g4(tab, w3);
            '    w3 ^= w2 ^ 16; g0(tab, w2);
            intW2 = intW2 Xor (intW1 Xor 9):  GG intW1, 2, 3, 4, 5
            intW1 = intW1 Xor (intW4 Xor 10): GG intW4, 6, 7, 8, 9
            intW4 = intW4 Xor (intW3 Xor 11): GG intW3, 0, 1, 2, 3
            intW3 = intW3 Xor (intW2 Xor 12): GG intW2, 4, 5, 6, 7
            intW2 = intW2 Xor (intW1 Xor 13): GG intW1, 8, 9, 0, 1
            intW1 = intW1 Xor (intW4 Xor 14): GG intW4, 2, 3, 4, 5
            intW4 = intW4 Xor (intW3 Xor 15): GG intW3, 6, 7, 8, 9
            intW3 = intW3 Xor (intW2 Xor 16): GG intW2, 0, 1, 2, 3

            ' Rule A2
            '    /* stepping rule A: */
            '    g1(tab, w1); w4 ^= w1 ^ 17;
            '    g2(tab, w4); w3 ^= w4 ^ 18;
            '    g3(tab, w3); w2 ^= w3 ^ 19;
            '    g4(tab, w2); w1 ^= w2 ^ 20;
            '    g0(tab, w1); w4 ^= w1 ^ 21;
            '    g1(tab, w4); w3 ^= w4 ^ 22;
            '    g2(tab, w3); w2 ^= w3 ^ 23;
            '    g3(tab, w2); w1 ^= w2 ^ 24;
            GG intW1, 4, 5, 6, 7: intW4 = intW4 Xor (intW1 Xor 17)
            GG intW4, 8, 9, 0, 1: intW3 = intW3 Xor (intW4 Xor 18)
            GG intW3, 2, 3, 4, 5: intW2 = intW2 Xor (intW3 Xor 19)
            GG intW2, 6, 7, 8, 9: intW1 = intW1 Xor (intW2 Xor 20)
            GG intW1, 0, 1, 2, 3: intW4 = intW4 Xor (intW1 Xor 21)
            GG intW4, 4, 5, 6, 7: intW3 = intW3 Xor (intW4 Xor 22)
            GG intW3, 8, 9, 0, 1: intW2 = intW2 Xor (intW3 Xor 23)
            GG intW2, 2, 3, 4, 5: intW1 = intW1 Xor (intW2 Xor 24)
    
            ' Perform 8 more iterations of mixing.
            ' Rule B2
            '    /* stepping rule B: */
            '    w2 ^= w1 ^ 25; g4(tab, w1);
            '    w1 ^= w4 ^ 26; g0(tab, w4);
            '    w4 ^= w3 ^ 27; g1(tab, w3);
            '    w3 ^= w2 ^ 28; g2(tab, w2);
            '    w2 ^= w1 ^ 29; g3(tab, w1);
            '    w1 ^= w4 ^ 30; g4(tab, w4);
            '    w4 ^= w3 ^ 31; g0(tab, w3);
            '    w3 ^= w2 ^ 32; g1(tab, w2);
            intW2 = intW2 Xor (intW1 Xor 25): GG intW1, 6, 7, 8, 9
            intW1 = intW1 Xor (intW4 Xor 26): GG intW4, 0, 1, 2, 3
            intW4 = intW4 Xor (intW3 Xor 27): GG intW3, 4, 5, 6, 7
            intW3 = intW3 Xor (intW2 Xor 28): GG intW2, 8, 9, 0, 1
            intW2 = intW2 Xor (intW1 Xor 29): GG intW1, 2, 3, 4, 5
            intW1 = intW1 Xor (intW4 Xor 30): GG intW4, 6, 7, 8, 9
            intW4 = intW4 Xor (intW3 Xor 31): GG intW3, 0, 1, 2, 3
            intW3 = intW3 Xor (intW2 Xor 32): GG intW2, 4, 5, 6, 7
                
            ' Convert four integers to bytes
            CopyMemory abytOutput(0), intW1, 2&
            CopyMemory abytOutput(2), intW2, 2&
            CopyMemory abytOutput(4), intW3, 2&
            CopyMemory abytOutput(6), intW4, 2&
                    
            ' Transfer encrypted data back to the temporary work array
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
                
        Next lngRound
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
    
EncryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        EncryptData = False
    Else
        EncryptData = True
    End If

    On Error GoTo 0
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte) As Boolean

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
    Dim lngRound      As Long
    Dim lngOffSet     As Long
    Dim lngProgress   As Long
    Dim lngDataLength As Long
    Dim intW1         As Integer
    Dim intW2         As Integer
    Dim intW3         As Integer
    Dim intW4         As Integer
    Dim abytInput(7)  As Byte
    Dim abytOutput(7) As Byte

    Const ROUTINE_NAME  As String = "DecryptData"

    On Error GoTo DecryptData_Error

    lngDataLength = UBound(abytData) + 1   ' Get size of incoming array
    
    ' 1-Jan-2009 Fixed lngOffSet range
    ' Decrypt the data
    For lngOffSet = 0 To (lngDataLength - 1) Step BLOCK_SIZE

        ' Added feature.
        ' Perform requested number of rounds
        For lngRound = 1 To mlngRounds
            
            ' Transfer eight bytes to work array
            CopyMemory abytInput(0), abytData(lngOffSet), BLOCK_SIZE
                    
            ' Convert eight bytes to four integers
            CopyMemory intW1, abytInput(0), 2&
            CopyMemory intW2, abytInput(2), 2&
            CopyMemory intW3, abytInput(4), 2&
            CopyMemory intW4, abytInput(6), 2&
        
            ' Perform 8 iterations of mixing.
            ' Inverse of Rule B2
            '    /* stepping rule A: */
            '    h1(tab, w2); w3 ^= w2 ^ 32;
            '    h0(tab, w3); w4 ^= w3 ^ 31;
            '    h4(tab, w4); w1 ^= w4 ^ 30;
            '    h3(tab, w1); w2 ^= w1 ^ 29;
            '    h2(tab, w2); w3 ^= w2 ^ 28;
            '    h1(tab, w3); w4 ^= w3 ^ 27;
            '    h0(tab, w4); w1 ^= w4 ^ 26;
            '    h4(tab, w1); w2 ^= w1 ^ 25;
            HH intW2, 4, 5, 6, 7: intW3 = intW3 Xor (intW2 Xor 32)
            HH intW3, 0, 1, 2, 3: intW4 = intW4 Xor (intW3 Xor 31)
            HH intW4, 6, 7, 8, 9: intW1 = intW1 Xor (intW4 Xor 30)
            HH intW1, 2, 3, 4, 5: intW2 = intW2 Xor (intW1 Xor 29)
            HH intW2, 8, 9, 0, 1: intW3 = intW3 Xor (intW2 Xor 28)
            HH intW3, 4, 5, 6, 7: intW4 = intW4 Xor (intW3 Xor 27)
            HH intW4, 0, 1, 2, 3: intW1 = intW1 Xor (intW4 Xor 26)
            HH intW1, 6, 7, 8, 9: intW2 = intW2 Xor (intW1 Xor 25)
    
            ' Use 24 iterations to mix the data for each round.
            '
            ' Inverse of Rule A2
            '    /* stepping rule B: */
            '    w1 ^= w2 ^ 24; h3(tab, w2);
            '    w2 ^= w3 ^ 23; h2(tab, w3);
            '    w3 ^= w4 ^ 22; h1(tab, w4);
            '    w4 ^= w1 ^ 21; h0(tab, w1);
            '    w1 ^= w2 ^ 20; h4(tab, w2);
            '    w2 ^= w3 ^ 19; h3(tab, w3);
            '    w3 ^= w4 ^ 18; h2(tab, w4);
            '    w4 ^= w1 ^ 17; h1(tab, w1);
            intW1 = intW1 Xor (intW2 Xor 24): HH intW2, 2, 3, 4, 5
            intW2 = intW2 Xor (intW3 Xor 23): HH intW3, 8, 9, 0, 1
            intW3 = intW3 Xor (intW4 Xor 22): HH intW4, 4, 5, 6, 7
            intW4 = intW4 Xor (intW1 Xor 21): HH intW1, 0, 1, 2, 3
            intW1 = intW1 Xor (intW2 Xor 20): HH intW2, 6, 7, 8, 9
            intW2 = intW2 Xor (intW3 Xor 19): HH intW3, 2, 3, 4, 5
            intW3 = intW3 Xor (intW4 Xor 18): HH intW4, 8, 9, 0, 1
            intW4 = intW4 Xor (intW1 Xor 17): HH intW1, 4, 5, 6, 7
    
            ' Inverse of Rule B1
            '    /* stepping rule A: */
            '    h0(tab, w2); w3 ^= w2 ^ 16;
            '    h4(tab, w3); w4 ^= w3 ^ 15;
            '    h3(tab, w4); w1 ^= w4 ^ 14;
            '    h2(tab, w1); w2 ^= w1 ^ 13;
            '    h1(tab, w2); w3 ^= w2 ^ 12;
            '    h0(tab, w3); w4 ^= w3 ^ 11;
            '    h4(tab, w4); w1 ^= w4 ^ 10;
            '    h3(tab, w1); w2 ^= w1 ^  9;
            HH intW2, 0, 1, 2, 3: intW3 = intW3 Xor (intW2 Xor 16)
            HH intW3, 6, 7, 8, 9: intW4 = intW4 Xor (intW3 Xor 15)
            HH intW4, 2, 3, 4, 5: intW1 = intW1 Xor (intW4 Xor 14)
            HH intW1, 8, 9, 0, 1: intW2 = intW2 Xor (intW1 Xor 13)
            HH intW2, 4, 5, 6, 7: intW3 = intW3 Xor (intW2 Xor 12)
            HH intW3, 0, 1, 2, 3: intW4 = intW4 Xor (intW3 Xor 11)
            HH intW4, 6, 7, 8, 9: intW1 = intW1 Xor (intW4 Xor 10)
            HH intW1, 2, 3, 4, 5: intW2 = intW2 Xor (intW1 Xor 9)

            ' Inverse of Rule A1
            '    /* stepping rule B: */
            '    w1 ^= w2 ^ 8; h2(tab, w2);
            '    w2 ^= w3 ^ 7; h1(tab, w3);
            '    w3 ^= w4 ^ 6; h0(tab, w4);
            '    w4 ^= w1 ^ 5; h4(tab, w1);
            '    w1 ^= w2 ^ 4; h3(tab, w2);
            '    w2 ^= w3 ^ 3; h2(tab, w3);
            '    w3 ^= w4 ^ 2; h1(tab, w4);
            '    w4 ^= w1 ^ 1; h0(tab, w1);
            intW1 = intW1 Xor (intW2 Xor 8): HH intW2, 8, 9, 0, 1
            intW2 = intW2 Xor (intW3 Xor 7): HH intW3, 4, 5, 6, 7
            intW3 = intW3 Xor (intW4 Xor 6): HH intW4, 0, 1, 2, 3
            intW4 = intW4 Xor (intW1 Xor 5): HH intW1, 6, 7, 8, 9
            intW1 = intW1 Xor (intW2 Xor 4): HH intW2, 2, 3, 4, 5
            intW2 = intW2 Xor (intW3 Xor 3): HH intW3, 8, 9, 0, 1
            intW3 = intW3 Xor (intW4 Xor 2): HH intW4, 4, 5, 6, 7
            intW4 = intW4 Xor (intW1 Xor 1): HH intW1, 0, 1, 2, 3

            ' Convert four integers to bytes
            CopyMemory abytOutput(0), intW1, 2&
            CopyMemory abytOutput(2), intW2, 2&
            CopyMemory abytOutput(4), intW3, 2&
            CopyMemory abytOutput(6), intW4, 2&
                    
            ' Transfer decrypted data back to the temporary work array
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
                
        Next lngRound
    
        ' An error occurred or user opted to STOP processing
        If gblnStopProcessing Then
            Exit For    ' exit For..Next loop
        End If
    
    Next lngOffSet
    
DecryptData_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        DecryptData = False
    Else
        DecryptData = True
    End If

    On Error GoTo 0
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

'    /**
'    * The key-dependent permutation GG on V^16 is a four-round Feistel network.
'    * The round function is a fixed byte-substitution table (permutation on V^8),
'    * the F-table.  Each round of GG incorporates a single byte from the key.
'    */
'    #define g(tab, w, i, j, k, l) \
'    { \
'       w ^= (word32)tab[i][w & 0xff] << 8; \
'       w ^= (word32)tab[j][w >>   8]; \
'       w ^= (word32)tab[k][w & 0xff] << 8; \
'       w ^= (word32)tab[l][w >>   8]; \
Private Sub GG(ByRef intWn As Integer, _
               ByVal a As Byte, _
               ByVal b As Byte, _
               ByVal c As Byte, _
               ByVal d As Byte)
                   
    intWn = intWn Xor (maintTable(a, (w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))  ' Shift left
    intWn = intWn Xor (maintTable(b, (w16Shift(intWn, -8) Mod MAX_BYTE)))             ' Shift right
    intWn = intWn Xor (maintTable(c, (w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))
    intWn = intWn Xor (maintTable(d, (w16Shift(intWn, -8) Mod MAX_BYTE)))
    
End Sub

'    * The inverse of the GG permutation.
'    #define h(tab, w, i, j, k, l) \
'    { \
'        w ^= (word32)tab[l][w >>   8]; \
'        w ^= (word32)tab[k][w & 0xff] << 8; \
'        w ^= (word32)tab[j][w >>   8]; \
'        w ^= (word32)tab[i][w & 0xff] << 8; \
Private Sub HH(ByRef intWn As Integer, _
               ByVal a As Byte, _
               ByVal b As Byte, _
               ByVal c As Byte, _
               ByVal d As Byte)
               
    intWn = intWn Xor (maintTable(d, (w16Shift(intWn, -8) Mod MAX_BYTE)))              ' Shift right
    intWn = intWn Xor (maintTable(c, (w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))   ' Shift left
    intWn = intWn Xor (maintTable(b, (w16Shift(intWn, -8) Mod MAX_BYTE)))
    intWn = intWn Xor (maintTable(a, (w16Shift((intWn And &HFF&), 8) Mod MAX_BYTE)))

End Sub

' ***************************************************************************
' Routine:       w16Shift (Short Integer)
'
' Description:   Shifts the bits to the right or left the specified number of
'                positions and returns the new value.  Bits "falling off"
'                the edge do not wrap around.  The fill bits are zeroes on
'                the opposite side.  Some common languages like C/C++ or
'                this job: ">>" or "<<".
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
Private Function w16Shift(ByVal lngValue As Long, _
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

    ' Called by Skipjack_File()
    '           Skipjack_String()
    
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
' Description:   F-table data consist of ASCII 0-255 values mixed with
'                no duplicates.
'
'                This is an alternate way of loading the F-table array.
'
'                For added security, I have made this table a little more
'                dynamic by creating several data strings and select an
'                appropriate one based on key length provided by user.
'
' Returns:       Mixed byte array of ASCII values 0-255
'
' ---------------------------------------------------------------------------
' Original Skipjack table data
'    strData = strData & "A3 D7 09 83 F8 48 F6 F4 B3 21 15 78 99 B1 AF F9 "
'    strData = strData & "E7 2D 4D 8A CE 4C CA 2E 52 95 D9 1E 4E 38 44 28 "
'    strData = strData & "0A DF 02 A0 17 F1 60 68 12 B7 7A C3 E9 FA 3D 53 "
'    strData = strData & "96 84 6B BA F2 63 9A 19 7C AE E5 F5 F7 16 6A A2 "
'    strData = strData & "39 B6 7B 0F C1 93 81 1B EE B4 1A EA D0 91 2F B8 "
'    strData = strData & "55 B9 DA 85 3F 41 BF E0 5A 58 80 5F 66 0B D8 90 "
'    strData = strData & "35 D5 C0 A7 33 06 65 69 45 00 94 56 6D 98 9B 76 "
'    strData = strData & "97 FC B2 C2 B0 FE DB 20 E1 EB D6 E4 DD 47 4A 1D "
'    strData = strData & "42 ED 9E 6E 49 3C CD 43 27 D2 07 D4 DE C7 67 18 "
'    strData = strData & "89 CB 30 1F 8D C6 8F AA C8 74 DC C9 5D 5C 31 A4 "
'    strData = strData & "70 88 61 2C 9F 0D 2B 87 50 82 54 64 26 7D 03 40 "
'    strData = strData & "34 4B 1C 73 D1 C4 FD 3B CC FB 7F AB E6 3E 5B A5 "
'    strData = strData & "AD 04 23 9C 14 51 22 F0 29 79 71 7E FF 8C 0E E2 "
'    strData = strData & "0C EF BC 72 75 6F 37 A1 EC D3 8E 62 8B 86 10 E8 "
'    strData = strData & "08 77 11 BE 92 4F 24 C5 32 36 9D CF F3 A6 BB AC "
'    strData = strData & "5E 6C A9 13 57 25 B5 E3 BD A8 3A 01 05 59 2A 46 "
' ---------------------------------------------------------------------------
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 12-Sep-2011  Kenneth Ives  kenaso@tx.rr.com
'              Documented routine
' 20-Oct-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated number of data mixing iterations
' ***************************************************************************
Private Function LoadWorkArrays() As Byte()

    ' Called by EvaluateKey()

    Dim lngIndex  As Long
    Dim strData   As String
    Dim avntData  As Variant
    Dim abytTmp() As Byte
    
    avntData = Empty        ' Always start with empty variants
    strData = vbNullString  ' Verify data string is empty
    
    ' See Skipjack_tbl.txt for additional sets of data
    ' strings from which to choose.  See application
    ' BuildTables.vbp to create different style CASE
    ' statements.
    '
    ' Determine which string of data to use based on key
    ' length selected by the user.
    '
    ' Each case statement:
    '    1.  Consists of all ASCII values (0-255) in two
    '        character hex equivalent with no duplicates
    '    2.  Mixed 100-400 iterations using Knuth Shuffle
    '
    Select Case mlngKeyLength
           Case 128, 448
                strData = strData & "70 60 A3 DA 1D C1 6B 76 D5 39 07 9E 17 DF E3 64 "
                strData = strData & "28 21 2B 41 8A 85 9F B1 B9 14 90 D7 71 AF 95 B8 "
                strData = strData & "B5 3B E8 35 6E FE 7B B0 88 49 20 15 F9 BB 79 E7 "
                strData = strData & "91 C3 18 6C 36 53 61 86 F7 3C 32 4C AE 38 10 BF "
                strData = strData & "55 06 D0 B7 BC 8B F1 AB EC 27 C4 EB 5B AA 65 89 "
                strData = strData & "1A 46 8F C9 A4 87 BD 67 CD 63 DD 7F 74 50 4A 33 "
                strData = strData & "3D 48 43 9D B3 ED CA 69 23 58 A7 8E 97 E4 AC 81 "
                strData = strData & "31 4B 1B C2 2C 45 62 4D 80 6D 05 0D 83 47 9A 99 "
                strData = strData & "24 EA C7 08 4F A6 7A 34 0C BA 0E FA DC 66 11 B6 "
                strData = strData & "5D BE 59 1E 25 7C 9C CE AD 5E F0 96 A9 93 A8 0F "
                strData = strData & "54 22 82 5C 5A CC EF 29 0B 2E B4 FD D1 8C 5F 2A "
                strData = strData & "FF A2 6A 51 F4 D3 FB CF A0 92 40 68 D8 98 FC 42 "
                strData = strData & "E2 2F D6 3F 2D E5 C8 E1 B2 16 CB D4 9B 8D EE 09 "
                strData = strData & "13 F3 0A 12 57 D9 4E D2 C5 04 19 A1 3E C0 7E 26 "
                strData = strData & "73 37 1F 30 DE 03 94 77 F2 E0 C6 3A 56 F6 75 DB "
                strData = strData & "E9 84 01 02 7D 72 44 52 1C F5 F8 00 A5 E6 78 6F "

           Case 160, 512
                strData = strData & "FE 4B DF 27 ED 8C 42 EC CE 86 C4 84 EF 40 A3 05 "
                strData = strData & "A6 E8 35 8F 07 CA A1 AE 22 5F F0 19 E5 44 EA F8 "
                strData = strData & "59 FA 13 3B 8A 18 D5 DB CD 74 29 CC 4C 98 E7 D6 "
                strData = strData & "0B 03 6F 4A 93 F1 2D 52 02 E1 2F 6B D3 AF 47 28 "
                strData = strData & "68 3D 4D F9 6A F5 EB 9F B2 B7 83 95 3F C6 E2 D1 "
                strData = strData & "80 55 67 61 81 B3 1A 25 BD A2 6C 36 C7 BE 0A A9 "
                strData = strData & "00 F6 B1 0D 1F 1C 16 43 79 1B 56 14 4F 9C BB 3E "
                strData = strData & "C5 FC 97 11 2B 04 50 0C DD 24 A5 F7 FF F3 7F CB "
                strData = strData & "60 E0 D7 33 31 72 E6 53 5B 4E 92 99 30 B6 B5 37 "
                strData = strData & "FB 57 AD 9E 12 94 23 77 01 69 70 2C 5D BF 82 AC "
                strData = strData & "48 C8 15 75 D4 20 87 34 8B F2 5C 32 91 2E 58 08 "
                strData = strData & "49 73 7C 5E DA B0 BA 62 90 9A 9B 10 21 54 9D DC "
                strData = strData & "39 7D AB BC 7B DE 8D 26 51 F4 85 0F 0E E3 45 5A "
                strData = strData & "B4 A0 EE 1D 8E 38 63 D9 65 D2 C3 A4 A7 89 7A 17 "
                strData = strData & "A8 1E AA 76 E9 6E 64 D0 78 C2 D8 88 E4 41 FD 46 "
                strData = strData & "3A C9 09 2A 3C 71 7E CF 06 C0 96 B8 C1 6D 66 B9 "

           Case 192, 576
                strData = strData & "43 72 34 C4 E1 BE 75 D1 89 D7 60 9C AA 8E 2D A3 "
                strData = strData & "10 4C 61 70 B5 64 E4 DE 7A 98 59 3E 68 BA 77 7C "
                strData = strData & "85 13 65 B7 21 35 4A C8 F4 EF E3 24 EA 8C E8 7D "
                strData = strData & "96 A7 A1 FE 04 4B 63 FA 57 38 F7 1F CF 95 17 C6 "
                strData = strData & "03 5C E6 48 8A DA EB CC CB B1 88 2B 80 3F 25 FC "
                strData = strData & "DD D9 AC 45 ED D3 7E 23 53 36 67 01 0A A6 4D DB "
                strData = strData & "09 74 73 8B 42 1E 29 33 FD 69 A8 F6 40 CA 1A A5 "
                strData = strData & "6F 90 B8 E2 62 EC 3D F2 F8 37 78 7B 47 B2 52 87 "
                strData = strData & "99 AD 71 44 32 FB 26 C1 AE D8 E7 8D 4F BB 19 82 "
                strData = strData & "39 E5 F0 5A 8F 12 A9 B3 D5 B4 11 6A 9B CE DF 83 "
                strData = strData & "C9 18 9D E9 9A 5B F3 2F 6B DC 3C 79 A4 0B 3B 76 "
                strData = strData & "9E 0C 2E 05 66 FF 6D 22 6C 3A 51 20 B6 D2 1B C3 "
                strData = strData & "C5 14 86 30 56 BD D4 06 28 00 49 A0 F1 16 5E AB "
                strData = strData & "F9 93 BF 94 7F 84 54 BC 46 E0 02 31 1D 0D EE 0E "
                strData = strData & "CD 2C 5F 6E C7 0F 55 5D 58 2A A2 81 97 D0 50 4E "
                strData = strData & "F5 07 C0 41 1C 91 92 15 B0 B9 9F D6 AF 27 C2 08 "

           Case 224, 640
                strData = strData & "DA B7 50 CA E7 8D B0 A3 87 31 6D F3 81 3D 9E C4 "
                strData = strData & "70 40 14 B6 8E 1C 8F 56 73 FD 1E A8 A9 07 27 4F "
                strData = strData & "5B F8 92 4C B1 23 F5 00 FC 35 AB 71 D0 39 ED 20 "
                strData = strData & "74 96 13 FF 9F 94 BC 86 3E 30 EF 93 BD 42 1A 26 "
                strData = strData & "37 06 F7 38 2F E5 9B 91 CD 85 7D 41 80 D6 C9 1F "
                strData = strData & "EA 83 E9 03 11 63 22 0E EB A1 32 E4 77 B2 46 C2 "
                strData = strData & "79 BA A4 5A 64 A0 EE C1 BB 66 97 55 F1 AC 12 C7 "
                strData = strData & "9C 28 59 4E 6F B8 16 2C 62 49 02 54 7F C5 89 AE "
                strData = strData & "43 09 9D 75 19 44 88 1B 33 8A E3 7B BE 04 D4 E2 "
                strData = strData & "6E 29 5F 76 A2 D7 48 68 3F F2 15 D9 5D 60 F9 61 "
                strData = strData & "B3 95 5C D1 9A E1 DB AF 08 4D 52 A6 E6 84 7E 36 "
                strData = strData & "4B 57 FB 6C 2B 72 DE 2A 18 D5 69 2E B4 CE B5 DF "
                strData = strData & "01 17 F0 5E F4 C8 A7 0A DC 0F 4A 58 1D 34 65 0B "
                strData = strData & "C6 A5 AA 3C F6 6B 99 D3 CF 0D CB 21 7A 8C EC 51 "
                strData = strData & "53 45 FA E0 D2 2D 6A D8 7C 67 25 C3 E8 3B FE AD "
                strData = strData & "CC 47 8B BF C0 DD 98 24 3A 78 05 10 0C 90 B9 82 "

           Case 256, 704
                strData = strData & "8C 5C EC 5A 30 48 E5 FB DA 28 38 F9 81 72 21 84 "
                strData = strData & "EA 46 93 89 1E 76 0A 66 25 78 33 44 96 A7 7D FE "
                strData = strData & "98 29 08 B1 3D 24 FC A1 1A C1 FF D8 D4 40 95 65 "
                strData = strData & "68 F3 6D CA 8B 9D 9F 4D 5F 2D 43 0E 06 7E AB C0 "
                strData = strData & "49 54 6F 58 AA CC EE B5 ED A4 6B 14 6C 27 CF AE "
                strData = strData & "03 CE 9E BD 87 BE 5B A3 2A 37 B2 FA 10 77 82 EB "
                strData = strData & "B4 7F 62 D3 0C BB 97 B3 DB AC D6 56 00 DD BF AF "
                strData = strData & "E4 13 4C 9C 3E 79 FD A8 7B DC 52 75 C6 3A E9 F4 "
                strData = strData & "91 85 1B D7 D2 F8 69 0F 90 55 39 57 45 11 20 0D "
                strData = strData & "E7 74 BC F6 A6 42 73 B6 8D EF 7A 01 B9 A2 92 16 "
                strData = strData & "BA F7 CD 94 1F 51 64 47 C9 AD 83 3F 0B E1 C5 E0 "
                strData = strData & "17 6A DE 9A C8 80 31 26 23 DF 02 61 8A E6 35 B7 "
                strData = strData & "C2 2F 59 88 4F C4 A0 4A 71 4E 4B 70 36 53 50 86 "
                strData = strData & "63 1D 19 34 E8 07 7C 18 22 E2 3C D1 99 F0 D5 15 "
                strData = strData & "1C 9B 3B 12 2E B8 60 04 F2 09 2B 5D D9 8F 32 D0 "
                strData = strData & "67 6E 2C A5 8E 41 E3 C3 A9 F5 CB F1 05 5E B0 C7 "

           Case 288, 768
                strData = strData & "0A 16 0F 2E 19 4D 3E 8C 82 A2 5D 91 DD D8 B1 E0 "
                strData = strData & "5A 2D 9E 27 81 86 FC C0 61 DE 13 C6 06 08 22 2F "
                strData = strData & "7C 95 01 E1 6C 5C 11 75 24 9B 3C 05 C7 B4 F7 AA "
                strData = strData & "14 E4 8D 34 F9 6E 4C 47 A1 20 B2 FD AF 0D E8 0E "
                strData = strData & "23 D4 F8 CC 30 79 C8 62 5F 42 E2 ED 18 32 44 FF "
                strData = strData & "B0 F0 4A CE 71 C2 D1 25 66 AD 52 C3 BF 8B 89 80 "
                strData = strData & "D9 FA F6 15 57 A9 2B 36 50 A0 7B 94 9F 21 DC 92 "
                strData = strData & "70 78 17 03 CD 83 BA 73 43 DA 3D 65 29 9D A6 8A "
                strData = strData & "B9 96 51 4B 33 4E 74 AC B8 7E EB 3F 55 D7 3A 59 "
                strData = strData & "1C 7F 12 A7 6F 76 BE AB F3 F2 1E 41 35 64 DF 2A "
                strData = strData & "39 E7 48 31 D3 3B F4 B6 97 09 2C 9C D2 85 5E F5 "
                strData = strData & "0C 88 AE 53 7A 8E B5 68 A5 BC F1 BB 37 A4 45 46 "
                strData = strData & "CF C1 D6 99 07 4F 7D CA EA 9A BD 26 58 E9 A3 54 "
                strData = strData & "6A 6D C9 0B 28 FB 8F 38 B3 E5 69 C4 77 60 6B 67 "
                strData = strData & "98 B7 40 90 93 C5 87 EE E3 04 10 A8 56 5B 02 72 "
                strData = strData & "DB 63 1B CB D5 EC E6 49 1A 84 D0 1D FE 00 EF 1F "

           Case 320, 832
                strData = strData & "F6 1D F9 B8 AC AD 3A 9A A3 95 5D 47 9C D1 FD EB "
                strData = strData & "04 4F 0A 16 A2 FC 91 09 E4 86 AB 00 33 6F 38 13 "
                strData = strData & "8F 77 B5 8E 0D 41 5A 11 62 54 D0 78 31 71 3E 17 "
                strData = strData & "01 50 D4 32 B9 DB B3 06 E8 51 25 40 5C 6B FE AA "
                strData = strData & "A9 CB E3 53 8C A8 6D 4B ED BE 2F E6 C6 DD 97 18 "
                strData = strData & "FB 67 27 C5 DA 9D 92 FF EA 3C F1 F4 A1 46 59 39 "
                strData = strData & "0B F2 2B BB 80 1C 3D A7 E0 A4 82 DC 1E 9B 68 BF "
                strData = strData & "AE CD 19 D7 0E 57 1A C2 D2 6A 87 08 BC B2 34 EE "
                strData = strData & "14 8B 72 99 20 B6 7F C3 49 8D 7E 2A 52 9E 4E 55 "
                strData = strData & "10 28 42 FA 98 26 7D 3B BD C4 23 5B 75 03 60 F7 "
                strData = strData & "69 22 79 6C B0 C1 89 94 74 D8 E2 4D CF D3 05 5E "
                strData = strData & "A0 84 0F A6 58 64 29 15 C7 B4 45 0C CC 35 30 76 "
                strData = strData & "73 90 36 2C E1 D5 DF 21 7C E9 E7 2D 07 3F B7 AF "
                strData = strData & "4C C9 E5 9F 96 C0 02 F3 24 1B 7B D6 88 CE 65 44 "
                strData = strData & "F5 12 BA 7A F0 63 B1 43 83 81 F8 93 CA 66 61 2E "
                strData = strData & "37 4A 48 EC EF C8 6E 70 A5 5F D9 56 85 8A DE 1F "

           Case 352, 896
                strData = strData & "05 B7 DD 22 9D 8B 00 FC 72 B9 F4 88 8A A8 E6 6D "
                strData = strData & "77 A7 04 4C 61 40 41 43 0A 8F 30 AE F3 58 D3 EA "
                strData = strData & "5F 92 0F A3 52 DB 48 23 7E 4E 68 C1 44 4F 13 B5 "
                strData = strData & "0D 84 18 F8 D4 AA F6 31 81 CB A9 BF 75 80 19 34 "
                strData = strData & "69 09 36 37 96 94 38 F1 93 9C EF D1 86 60 66 BE "
                strData = strData & "71 5A 2E D0 DF 9A 06 2A D2 2D FE 21 C9 D6 F5 1F "
                strData = strData & "5B D9 A0 57 02 11 2B C2 51 CA 3C 78 E7 C4 27 DC "
                strData = strData & "DE 89 BD B2 E3 62 7F 8C 1E 0E CD 4B A5 95 E1 35 "
                strData = strData & "6A 85 4D 20 29 55 59 B6 FD 45 FF F2 7B 83 91 BB "
                strData = strData & "32 F0 9E 07 AF A4 DA 14 C8 1C 6B BC 42 7A E9 F7 "
                strData = strData & "AB 50 5D A1 3A 63 53 E2 87 67 CE 74 9F EC B1 7C "
                strData = strData & "8D FA ED B0 76 D5 3F C3 8E 70 C7 B8 AC BA 0C 47 "
                strData = strData & "6F 9B 08 25 C0 3E 33 B3 73 7D 1D 97 17 26 98 6C "
                strData = strData & "10 54 65 3B 39 12 2F 82 99 16 5E 24 F9 B4 E4 2C "
                strData = strData & "0B 1A E5 D8 C6 46 5C FB A6 A2 EE 03 C5 15 E8 3D "
                strData = strData & "1B D7 28 EB 56 AD 4A 90 E0 49 CF 01 79 CC 6E 64 "

           Case 384, 960
                strData = strData & "10 DF B7 DE 8F 3A 15 68 5E 5A 65 EC BA F8 7B 95 "
                strData = strData & "A6 46 AE FF 61 0E D5 CE 7F FA 27 CB 66 57 24 4B "
                strData = strData & "8A F9 1F D2 E3 4D 21 53 E2 79 06 28 87 B8 12 6E "
                strData = strData & "AF 9D 49 54 CC C7 6A F7 A5 31 9F 6B 16 82 5D 69 "
                strData = strData & "56 F1 C4 1D 93 7D BB 2E 17 B9 9C 0D 84 0F 05 0B "
                strData = strData & "A9 DD 34 1E 33 8E 78 6C F5 E6 EF 80 0C E1 B6 ED "
                strData = strData & "CF 62 A0 4E 60 37 6D 38 35 4F DA 81 EA 44 52 B1 "
                strData = strData & "EE 1C F3 A3 7A D3 6F FB 71 48 25 3B 42 8D D8 86 "
                strData = strData & "41 C0 72 7C 70 A4 E9 DC 23 91 9A B4 AB 74 67 03 "
                strData = strData & "2B 02 32 A2 20 36 AA 5B F2 07 C6 18 09 85 99 B5 "
                strData = strData & "43 C3 3E FE FD 9E BC AD D0 C9 29 73 98 2C BF 7E "
                strData = strData & "B3 00 76 47 14 08 5C 59 F0 04 C1 E8 3F 89 97 A8 "
                strData = strData & "4A FC C2 2D BD 1A AC 77 92 58 22 A1 63 30 11 8B "
                strData = strData & "26 40 E0 75 64 3C F4 CA 1B E4 D7 C8 01 19 50 0A "
                strData = strData & "F6 90 2A 8C CD D1 96 A7 DB 94 83 C5 B0 45 EB E5 "
                strData = strData & "4C 39 D6 3D BE 88 5F 13 B2 51 55 2F E7 D4 9B D9 "

           Case 416, 1024
                strData = strData & "67 84 1F E9 49 2D EA 55 71 7D 5C AB 0E E6 F4 A4 "
                strData = strData & "1E B5 7E 21 39 53 DD 6C 4E D4 20 14 2B D2 77 FA "
                strData = strData & "4D CE 06 0C EE E7 65 7F 2C AF CB 0F 4A A0 6A 56 "
                strData = strData & "94 F9 F3 27 B2 ED D9 DC B7 2E 33 43 AA 51 4C 83 "
                strData = strData & "96 04 C7 38 C8 E0 8E 48 03 FE F0 28 8D C9 87 05 "
                strData = strData & "6B D1 22 B9 5F BC 72 07 34 99 08 D6 2F A3 47 C5 "
                strData = strData & "3A 54 A6 6F 44 79 70 CD 91 D0 E1 9C 25 9D FF 7B "
                strData = strData & "3D BA 32 DF C2 0B C3 AE F5 B8 DE A2 30 3F 88 59 "
                strData = strData & "93 86 BB 1B 5A DB F8 B1 61 10 D5 BE FC F6 FD E3 "
                strData = strData & "D3 60 15 50 74 5D 52 58 46 8F D7 45 6D 78 E5 1A "
                strData = strData & "66 7A AD 9A F2 B3 A5 BF 82 24 16 64 68 0A 9F B0 "
                strData = strData & "0D 13 5B 8C 69 EF 98 29 3C A1 40 CF E4 90 1C 3B "
                strData = strData & "2A 19 FB AC 63 9E 17 41 4B 31 81 75 8A 80 35 4F "
                strData = strData & "EB 23 8B 3E 7C A8 62 73 F7 00 09 57 95 92 E8 42 "
                strData = strData & "1D B6 36 37 DA 76 F1 89 18 6E D8 5E E2 02 A7 C0 "
                strData = strData & "B4 C4 9B CA EC CC BD 97 C1 11 A9 85 01 26 12 C6 "
    End Select

    avntData = Split(strData, Chr$(32))   ' convert data string into an array
    ReDim abytTmp(MAX_BYTE)               ' Resize temp byte array
    
    ' Convert data to a smaller footprint
    For lngIndex = 0 To (MAX_BYTE - 1)
        abytTmp(lngIndex) = CByte("&H" & avntData(lngIndex))
    Next lngIndex
    
    LoadWorkArrays = abytTmp()    ' Return data in byte array format
    
    Erase abytTmp()          ' Always empty arrays when not needed
    avntData = Empty         ' Always empty variants when not needed
    strData = vbNullString   ' Empty data string
    
End Function

Private Sub ResetVariables()

    Erase maintTable()  ' Always empty arrays when not needed
    mcurOrigLength = 0@
    
End Sub

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.
                
    ResetVariables
    Erase mabytPassword()   ' Always start with empty arrays
         
    ' Set property values
    KeyLength = 256        ' Default value
    EncryptRounds = 1      ' Number of rounds of encryption
    CreateNewFile = True   ' Default is to create a new output file
    
End Sub

Private Sub Class_Terminate()
    
    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
            
    ResetVariables
    Erase mabytPassword()      ' Always empty arrays when not needed

End Sub
