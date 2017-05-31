VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCryptoAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cCryptoAPI  (clsCryptoAPI.cls)
'
' WARNING:       This has been tested on Windows 32-bit and 64-bit versions
'                of Vista and newer.  Windows XP or older are no longer
'                supported.
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
' Description:  CryptoAPI is Microsoft's answer to the encryption world.
'
'               Ciphers used are:
'                   RC2 (Rivest Cipher 2)
'                   RC4 (Rivest Cipher 4)
'                   DES (Data Encryption Standard)
'                   DES3 (also known as Triple DES)
'                   AES-128 (Advanced Encryption Standard [Rijndael])
'                   AES-192
'                   AES-256
'                   BASE64 (Base64, also known as MIME encoding, translates
'                           binary into safe text. It is used to send
'                           attachments in email and to change small bits of
'                           unsafe high-character data into stuff that is a
'                           lot nicer for text-based system)
'
' Warning:      The weakest link in any encryption is the password. This
'               includes the security, length and components that make up
'               the password.  Password length and components are usually
'               dictated by organizational standards.  Security is the
'               responsibility of the user.  Do not write down or discuss
'               your password with others.  Lock your screensaver when not
'               at your workstation.
'
' Note:         After encryption, data sizes will not match original sizes.
'               This is due to internal padding and the storing of information
'               required to decrypt the data later.
'
' Suggestions:  To strengthen your encryption, prefix each block of data with
'               four to twelve bytes of salt (random) data. After decrypting
'               each block do not forget to drop the x-number of prefixed
'               random bytes.
'
'                                     or
'
'                Compress the file or data string first using your favorite
'                archive utility (Zip, Rar, etc.).  Now encrypt the data.
'
'                                     or
'
'               Encode data with Base64 and then perform your encryption.
'               After decryption, decode your data with Base64.  Most
'               messaging is done with string data and not file data.
'
'               The above are just suggestions.  Happy coding.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 15-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Added Base64 encryption/decryption
'              - Updated EncryptData(), DecryptData() routines
'              - Cleaned up code and added documentation
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME            As String = "clsCryptoAPI"
  Private Const kB_64                  As Long = &H10000            ' 65536 bytes
  Private Const MB_5                   As Long = &H500000           ' 5242880 bytes
  Private Const BLOCK_SIZE             As Long = 16
  Private Const MAX_FILE_SIZE          As Currency = 10000000000@   ' Approx 10 Gig

  ' Microsoft Cipher constants
  Private Const CRYPT_NO_SALT          As Long = &H10&        ' Used to designate no salt values while creating session key
  Private Const CRYPT_STRING_BASE64    As Long = 1            ' Base64 string cipher
  Private Const ALG_TYPE_BLOCK         As Long = 1536         ' Used by block ciphers
  Private Const ALG_TYPE_STREAM        As Long = 2048         ' Used by stream ciphers
  Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576        ' Used by all ciphers
  Private Const ALG_SID_RC4            As Long = 1            ' RC4 stream cipher
  Private Const ALG_SID_DES            As Long = 1            ' DES block cipher
  Private Const ALG_SID_RC2            As Long = 2            ' RC2 block cipher
  Private Const ALG_SID_3DES           As Long = 3            ' Triple DES block cipher
  Private Const ALG_SID_AES_128        As Long = 14           ' AES-128 block cipher
  Private Const ALG_SID_AES_192        As Long = 15           ' AES-192 block cipher
  Private Const ALG_SID_AES_256        As Long = 16           ' AES-256 block cipher
  Private Const CALG_RC4               As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4       ' 26625
  Private Const CALG_DES               As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES        ' 26113
  Private Const CALG_RC2               As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2        ' 26114
  Private Const CALG_3DES              As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES       ' 26115
  Private Const CALG_AES_128           As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128    ' 26126
  Private Const CALG_AES_192           As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192    ' 26127
  Private Const CALG_AES_256           As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256    ' 26128
  ' Microsoft Hash constants
  Private Const ALG_TYPE_ANY           As Long = 0            ' Used by all hashes
  Private Const ALG_SID_MD2            As Long = 1            ' 16 byte hashed return length
  Private Const ALG_SID_MD4            As Long = 2            ' 16 byte hashed return length
  Private Const ALG_SID_MD5            As Long = 3            ' 16 byte hashed return length
  Private Const ALG_SID_SHA1           As Long = 4            ' 20 byte hashed return length
  Private Const ALG_SID_SHA_256        As Long = 12           ' 32 byte hashed return length
  Private Const ALG_SID_SHA_384        As Long = 13           ' 48 byte hashed return length
  Private Const ALG_SID_SHA_512        As Long = 14           ' 64 byte hashed return length
  Private Const ALG_CLASS_HASH         As Long = 32768        ' Used by all hashes
  Private Const CALG_MD2               As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2       ' 32769
  Private Const CALG_MD4               As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4       ' 32770
  Private Const CALG_MD5               As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5       ' 32771
  Private Const CALG_SHA1              As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1      ' 32772
  Private Const CALG_SHA_256           As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256   ' 32780
  Private Const CALG_SHA_384           As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384   ' 32781
  Private Const CALG_SHA_512           As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512   ' 32782
  ' Microsoft Provider type constants
  Private Const PROV_RSA_FULL          As Long = 1            ' Provider type ID
  Private Const PROV_RSA_AES           As Long = 24           ' AES Provider type ID
  Private Const CRYPT_NEWKEYSET        As Long = &H8&         ' If needed, create a container
  Private Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000   ' -268435456
  ' Microsoft provider names
  ' Reg key:  HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography\Defaults\Provider\
  Private Const MS_DEFAULT_PROV        As String = "Microsoft Base Cryptographic Provider v1.0"
  Private Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"
  Private Const MS_ENH_RSA_AES_PROV    As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"
  Private Const MS_ENH_RSA_AES_PROV_XP As String = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumCryptoAPI_HashAlgorithms
      eCryptoAPI_MD2       ' 0
      eCryptoAPI_MD4       ' 1
      eCryptoAPI_MD5       ' 2
      eCryptoAPI_SHA1      ' 3
      eCryptoAPI_SHA256    ' 4
      eCryptoAPI_SHA384    ' 5
      eCryptoAPI_SHA512    ' 6
  End Enum

  Public Enum enumCryptoAPI_CipherAlgorithms
      eCryptoAPI_RC2       ' 0
      eCryptoAPI_RC4       ' 1
      eCryptoAPI_DES       ' 2
      eCryptoAPI_3DES      ' 3
      eCryptoAPI_AES_128   ' 4
      eCryptoAPI_AES_192   ' 5
      eCryptoAPI_AES_256   ' 6
      eCryptoAPI_BASE64    ' 7
  End Enum

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The CryptCreateHash function initiates the hashing of a stream of
  ' data. It creates and returns to the calling application a handle
  ' to a CSP hash object. This handle is used in subsequent calls to
  ' CryptHashData and CryptHashSessionKey to hash session keys and
  ' other streams of data.
  Private Declare Function CryptCreateHash Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal algid As Long, _
          ByVal hkey As Long, ByVal dwFlags As Long, _
          ByRef phHash As Long) As Long

  ' The CryptHashData function adds data to a specified hash object.
  ' This function and CryptHashSessionKey can be called multiple
  ' times to compute the hash of long or discontinuous data streams.
  Private Declare Function CryptHashData Lib "advapi32.dll" _
          (ByVal hhash As Long, ByVal pbData As String, _
          ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

  'The CryptDestroyHash function destroys the hash object referenced
  ' by the hHash parameter. After a hash object has been destroyed,
  ' it can no longer be used.  The destruction of hash objects after
  ' their use is finished is recommended for sedblity reasons.
  Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
          (ByVal hhash As Long) As Long

  ' The CryptAcquireContext function is used to acquire a handle to a
  ' particular key container within a particular cryptographic service
  ' provider (CSP). This returned handle can then be used to make
  ' calls to the selected CSP.  This function performs two operations.
  ' It first attempts to find a CSP with the characteristics described
  ' in the dwProvType and pszProvider parameters. If the CSP is found,
  ' the function attempts to find a key container within the CSP
  ' matching the name specified by the pszContainer parameter. With the
  ' appropriate setting of dwFlags, this function can also create and
  ' destroy key containers. If function succeeds, return value is nonzero.
  Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
          Alias "CryptAcquireContextA" (ByRef phProv As Long, _
          ByVal pszContainer As String, ByVal pszProvider As String, _
          ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

  ' The CryptReleaseContext function releases the handle of a
  ' cryptographic service provider (CSP) and a key container. At each
  ' call to this function, the reference count on the CSP is reduced
  ' by one. When the reference count reaches zero, the context is fully
  ' released and it can no longer be used by any function in the application.
  ' An application calls this function after finishing the use of the CSP.
  ' After this function is called, the released CSP handle is no longer
  ' valid. This function does not destroy key containers or key pairs.
  Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal dwFlags As Long) As Long

  ' The CryptDeriveKey function generates cryptographic session keys derived
  ' from a base data value. This function guarantees that when the same CSP
  ' and algorithms are used, the keys generated from the same base data are
  ' identical. The base data can be a password or any other user data.  This
  ' function is the same as CryptGenKey, except that the generated session
  ' keys are derived from base data instead of being random. CryptDeriveKey
  ' can only be used to generate session keys. It cannot generate
  ' public/private key pairs.
  Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
          (ByVal hProv As Long, ByVal algid As Long, _
          ByVal hBaseData As Long, ByVal dwFlags As Long, _
          ByRef phKey As Long) As Long

  ' The CryptEncrypt function encrypts data. The algorithm used to encrypt
  ' the data is designated by the key held by the CSP module and is
  ' referenced by the hKey parameter.
  Private Declare Function CryptEncrypt Lib "advapi32.dll" _
          (ByVal hkey As Long, ByVal hhash As Long, ByVal Final As Long, _
          ByVal dwFlags As Long, ByVal pbData As String, _
          ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long

  ' The CryptDecrypt function encrypts data. The algorithm used to decrypt
  ' the data is designated by the key held by the CSP module and is
  ' referenced by the hKey parameter.
  Private Declare Function CryptDecrypt Lib "advapi32.dll" _
          (ByVal hkey As Long, ByVal hhash As Long, ByVal Final As Long, _
          ByVal dwFlags As Long, ByVal pbData As String, _
          ByRef pdwDataLen As Long) As Long

  ' The CryptDestroyKey function releases the handle referenced by the hKey
  ' parameter. After a key handle has been released, it becomes invalid and
  ' cannot be used again.
  Private Declare Function CryptDestroyKey Lib "advapi32.dll" _
          (ByVal hkey As Long) As Long

  ' The CryptBinaryToString function is used to convert data to
  ' Base64 encrypted data format.  Failure is return value of zero.
  Private Declare Function CryptBinaryToString Lib "Crypt32" _
          Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, _
          ByVal cbBinary As Long, ByVal dwFlags As Long, _
          ByVal pszString As Long, ByRef pcchString As Long) As Long

  ' The CryptStringToBinary function is used to convert Base64 encrypted
  ' data back to its original format.  Failure is return value of zero.
  Private Declare Function CryptStringToBinary Lib "Crypt32" _
          Alias "CryptStringToBinaryW" (ByVal pszString As Long, _
          ByVal cchString As Long, ByVal dwFlags As Long, _
          ByVal pbBinary As Long, ByRef pcbBinary As Long, _
          ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

  ' The GetLastError function returns the calling thread's last-error
  ' code value.  Most Win32 functions set their calling thread's
  ' last-error value when they fail; a few functions set it when they
  ' succeed.  You should call the GetLastError function immediately when
  ' a function's return value indicates that such a call will return
  ' useful data. That is because some functions call SetLastError(0) when
  ' they succeed, wiping out the error code set by the most recently
  ' failed function.
  Private Declare Function GetLastError Lib "kernel32" () As Long

  ' The CopyMemory function copies a block of memory from one location to
  ' another.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

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
  Private mblnAESProvider    As Boolean    ' Flag denoting AES strong provider
  Private mblnCreateNewFile  As Boolean
  Private mlngBlockSize      As Long
  Private mlngKeyLength      As Long
  Private mlngHashObject     As Long
  Private mlngSessionKey     As Long
  Private mlngHashMethod     As Long
  Private mlngCipherMethod   As Long
  Private mlngProviderHandle As Long       ' CryptoAPI provider handle
  Private mcurMaxSize        As Currency
  Private mcurAccrued        As Currency
  Private mcurOrigLength     As Currency
  Private mabytPassword()    As Byte
  Private mlngHashAlgo       As enumCryptoAPI_HashAlgorithms
  Private mlngCipherAlgo     As enumCryptoAPI_CipherAlgorithms

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

' Comment out Version Property if not used in this application
Public Property Get Version() As String
    ' Version of the DLL this class is associated with.
    Version = DLL_NAME & ".dll v" & GetFileVerData(App.Path & "\" & DLL_NAME & ".dll") & _
              vbNewLine & " Copyright (c) 2004-2017 - Kenneth Ives  kenaso@tx.rr.com"
End Property

Public Property Get AES_Ready() As Boolean
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.AES_Ready

    ' Test for a provider handle
    If mlngProviderHandle < 1 Then
        GetProviderHandle   ' Get provider handle
    End If

    AES_Ready = mblnAESProvider

End Property

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

Public Property Let CreateNewFile(ByVal blnNewFile As Boolean)

    ' 21-Feb-2012 Added Property to desiginate if input file
    '             is to be overwritten after encryption/decryption.
    '
    ' TRUE - Create new file to hold encrypted/decrypted data
    ' FALSE - Overwrite input file after encryption/decryption
    mblnCreateNewFile = blnNewFile

End Property

Public Property Let HashMethod(ByVal lngHashMethod As enumCryptoAPI_HashAlgorithms)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashMethod = 4
    Select Case lngHashMethod
           Case 0 To 6: mlngHashMethod = lngHashMethod
           Case Else:   mlngHashMethod = 4               ' Default is SHA-256
    End Select

End Property

Public Property Let CipherMethod(ByVal lngCipherMethod As enumCryptoAPI_CipherAlgorithms)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.CipherMethod = 1
    Select Case lngCipherMethod
           Case 0 To 7: mlngCipherMethod = lngCipherMethod
           Case Else:   mlngCipherMethod = 1                 ' Default is RC4
    End Select

End Property

Public Property Let BlockSize(ByVal lngBlockSize As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.BlockSize = 128
    ' Used with AES (Rijndael) cipher only.
    Select Case lngBlockSize
           Case 128:  mlngBlockSize = 128
           Case 192:  mlngBlockSize = 192
           Case 256:  mlngBlockSize = 256
           Case Else: mlngBlockSize = 128   ' Use default value
    End Select

End Property

Public Property Let KeyLength(ByVal lngKeyLen As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.KeyLength = 128
    ' Used with AES (Rijndael) cipher only.
    Select Case lngKeyLen
           Case 128:  mlngKeyLength = 128
           Case 192:  mlngKeyLength = 192
           Case 256:  mlngKeyLength = 256
           Case Else: mlngKeyLength = 128   ' Use default value
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

        ReDim mabytPassword(UBound(vntData) + 1)

        For lngIndex = 0 To UBound(vntData)
            mabytPassword(lngIndex) = CByte(vntData(lngIndex))
        Next lngIndex

        ReDim Preserve mabytPassword(UBound(vntData))

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
' Routine:       CryptoAPI_File
'
' Description:   Function to encrypt/decrypt a file.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)   ' Convert file location to byte array 
'                objCryptoAPI.CryptoAPI_File abytFile(), blnEncrypt            ' Encrypt/Decrypt data
' 
' Parameters:    abytFile() - Path/filename in byte array
'                blnEncrypt - Boolean flag designating to encrypt data
'                             TRUE - Encrypt data
'                             FALSE - Decrypt data
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function CryptoAPI_File(ByRef abytFile() As Byte, _
                               ByVal blnEncrypt As Boolean) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String  ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim hFile1       As Long
    Dim hFile2       As Long
    Dim lngBlockSize As Long
    Dim lngFinalPass As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curProcessed As Currency
    Dim abytData()   As Byte
    Dim abytSize()   As Byte
    Dim abytTemp()   As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "CryptoAPI_File"

    On Error GoTo CryptoAPI_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    CryptoAPI_File = False            ' Preset to FALSE
    Set objBigFiles = New cBigFiles   ' Instantiate class objects
    Erase abytData()                  ' Always start with empty arrays
    Erase abytSize()
    lngFinalPass = 0

    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal

    ' See if a new file is to be created after
    ' performing encryption or decryption.
    ' ex:  Append "ENC" extension to target file
    '      designating the file is encrypted.
    If mblnCreateNewFile Then
        Select Case blnEncrypt

               Case True
                    ' First verify file name has an extension
                    If InStr(strFileName, ".") > 0 Then
                        strTarget = Mid$(strFileName, 1, InStrRev(strFileName, ".") - 1) & ENCRYPT_EXT
                    Else
                        ' No file extension
                        strTarget = strFileName & ENCRYPT_EXT
                    End If

               Case False
                    ' First verify file name has an extension
                    If InStr(strFileName, ".") > 0 Then
                        strTarget = Mid$(strFileName, 1, InStrRev(strFileName, ".") - 1) & DECRYPT_EXT
                    Else
                        ' No file extension
                        strTarget = strFileName & DECRYPT_EXT
                    End If
        End Select
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo CryptoAPI_File_CleanUp
    End If

    With objBigFiles
        .CalcFileSize strSource, curAmtLeft   ' Get size of file

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
            GoTo CryptoAPI_File_CleanUp
        End If

        ' If performing decryption then do not
        ' process the last few bytes of the file
        ' because this is where we stored the
        ' size of the original file.
        If blnEncrypt Then

            mcurOrigLength = curAmtLeft                  ' Capture total byte count for file
            ReDim abytSize(8)                            ' Size array to hold data length
            CopyMemory abytSize(0), mcurOrigLength, 8&   ' Copy original length to temp array
            ReDim Preserve abytSize(7)                   ' Drop trailing null value
            MixAppendedData abytSize()                   ' Encrypt appended data

        Else

            curAmtLeft = curAmtLeft - 8   ' Adjust amount that has to be decrypted
            curFilePos1 = curAmtLeft      ' Set pointer to capture the last 8 bytes

            If Not .OpenReadOnly(strSource, hFile1) Then      ' Open the source file
                gblnStopProcessing = True
            End If

            ReDim abytSize(7)                                 ' Size array to hold data length
            .API_ReadFile hFile1, curFilePos1, abytSize()     ' Capture last 8 bytes of source file
            .API_CloseFile hFile1                             ' close source file

            ' An error occurred or user opted to STOP processing
            DoEvents
            If gblnStopProcessing Then
                GoTo CryptoAPI_File_CleanUp
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
                GoTo CryptoAPI_File_CleanUp

            End If
        End If

        mcurMaxSize = mcurOrigLength   ' Save file size for progress calcs

        mcurAccrued = 0@   ' Init accrued progress
        curProcessed = 0@
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo CryptoAPI_File_CleanUp
        End If

        strTempFile = CreateTempFile()   ' Create a temp target file

        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo CryptoAPI_File_CleanUp
        End If

        Do
            Erase abytData()                          ' Empty data array
            lngBlockSize = GetBlockSize(curAmtLeft)   ' Process the source file in chunks

            ReDim abytData(lngBlockSize - 1)                   ' Size receiving array
            curAmtLeft = curAmtLeft - CCur(lngBlockSize)       ' Adjust amount of file left to process
            curProcessed = curProcessed + CCur(lngBlockSize)   ' Calc current processed

            DoEvents
            If curAmtLeft < 1 Then
                lngFinalPass = 1   ' Set flag to process end of data
            End If

            ' Read source file
            If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
                gblnStopProcessing = True
                Exit Do   ' exit Do..Loop
            End If

            curFilePos1 = curFilePos1 + CCur(UBound(abytData) + 1)   ' Adjust pointers accordingly

            If blnEncrypt Then

                ' Encrypt data
                If Not EncryptData(abytData(), lngFinalPass) Then
                    gblnStopProcessing = True
                    Exit Do   ' exit Do..Loop
                End If

            Else

                ' Decrypt data
                If DecryptData(abytData(), lngFinalPass) Then

                    ' Remove trailing padding on last pass
                    If lngFinalPass = 1 Then

                        ReDim abytTemp(lngBlockSize)                        ' Size temp hold array
                        CopyMemory abytTemp(0), abytData(0), lngBlockSize   ' Copy decrypted data
                        ReDim abytData(lngBlockSize)                        ' Resize data array

                        If curProcessed > mcurOrigLength Then
                            lngBlockSize = UBound(abytTemp) - (curProcessed - mcurOrigLength)
                            ReDim Preserve abytTemp(lngBlockSize)   ' Resize data array to exact length
                        End If

                        CopyMemory abytData(0), abytTemp(0), UBound(abytTemp)   ' Copy temp array to data array
                        ReDim Preserve abytData(lngBlockSize - 1)               ' Resize data array to exact length

                    End If

                Else
                    gblnStopProcessing = True
                    Exit Do   ' exit Do..Loop
                End If

            End If

            ' Write to target file
            If Not .API_WriteFile(hFile2, curFilePos2, abytData()) Then
                gblnStopProcessing = True
                Exit Do   ' exit Do..Loop
            End If

            curFilePos2 = curFilePos2 + CCur(UBound(abytData) + 1)    ' Adjust pointers accordingly

            ' If nothing else to read then leave
            DoEvents
            If curAmtLeft < 1 Then
                Exit Do   ' exit Do..Loop
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
            GoTo CryptoAPI_File_CleanUp
        End If

        .API_CloseFile hFile1    ' Close all files opened by this routine
        .API_CloseFile hFile2

        ' Add important data to end of encrypted file
        If blnEncrypt Then

            CopyMemory abytSize(0), mcurOrigLength, 8&   ' Save original file size [0-7]
            .CalcFileSize strTempFile, curFilePos2       ' Get size of target file
            MixAppendedData abytSize()                   ' Encrypt appended data

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
                GoTo CryptoAPI_File_CleanUp

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
    CryptoAPI_File = True            ' Good finish

CryptoAPI_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Verify all open files have been closed
        .API_CloseFile hFile2
    End With
    Set objBigFiles = Nothing   ' Free class object from memory

    ' Destroy hash object
    If mlngHashObject <> 0 Then
        Call CryptDestroyHash(mlngHashObject)
        mlngHashObject = 0
    End If

    ' Destroy session key
    If mlngSessionKey <> 0 Then
        Call CryptDestroyKey(mlngSessionKey)
        mlngSessionKey = 0
    End If

    ReleaseProvider   ' Release provider handle
    Erase abytData()  ' Always empty arrays when not needed
    Erase abytSize()

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

CryptoAPI_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CryptoAPI_File_CleanUp

End Function

' ***************************************************************************
' Routine:       Base64_File
'
' Description:   Function to encrypt/decrypt a file to Base64 format.
' 
' Syntax:        abytFile() = StrConv("C:\Test\Testfile.txt", vbFromUnicode)   ' Convert file location to byte array 
'                objCryptoAPI.CryptoAPI_File abytFile(), blnEncrypt            ' Encrypt/Decrypt data
' 
' Parameters:    abytFile() - Path/filename in byte array
'                blnEncrypt - Boolean flag designating to encrypt data
'                             TRUE - Encrypt data
'                             FALSE - Decrypt data
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function Base64_File(ByRef abytFile() As Byte, _
                            ByVal blnEncrypt As Boolean) As Boolean

    Dim strPath      As String
    Dim strSource    As String
    Dim strTarget    As String   ' Not used when overlaying source file
    Dim strTempFile  As String
    Dim strFileName  As String
    Dim hFile1       As Long
    Dim hFile2       As Long
    Dim curFilePos1  As Currency
    Dim curFilePos2  As Currency
    Dim curProcessed As Currency
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    Const ROUTINE_NAME As String = "Base64_File"

    On Error GoTo Base64_File_Error

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If

    Base64_File = False        ' Preset to FALSE
    Set objBigFiles = New cBigFiles   ' Instantiate class objects

    Erase abytData()   ' Always start with empty arrays

    strSource = ByteArrayToString(abytFile())           ' convert incoming array to string
    strPath = QualifyPath(GetPath(strSource))           ' Capture source path
    strFileName = GetFilename(strSource)                ' Capture source filename
    SetFileAttributes strSource, FILE_ATTRIBUTE_NORMAL  ' Reset file attributes to normal

    ' See if a new file is to be created after
    ' performing encryption or decryption.
    ' ex:  Append "ENC" extension to target file
    '      designating the file is encrypted.
    If mblnCreateNewFile Then
        Select Case blnEncrypt

               Case True
                    ' First verify file name has an extension
                    If InStr(strFileName, ".") > 0 Then
                        strTarget = Mid$(strFileName, 1, InStrRev(strFileName, ".") - 1) & ENCRYPT_EXT
                    Else
                        ' No file extension
                        strTarget = strFileName & ENCRYPT_EXT
                    End If

               Case False
                    ' First verify file name has an extension
                    If InStr(strFileName, ".") > 0 Then
                        strTarget = Mid$(strFileName, 1, InStrRev(strFileName, ".") - 1) & DECRYPT_EXT
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

    With objBigFiles
        .CalcFileSize strSource, mcurOrigLength   ' Get size of file

        ' See if file is empty
        If mcurOrigLength < 1 Then

            InfoMsg "There is no data to encrypt." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

            gblnStopProcessing = True
            GoTo Base64_File_CleanUp

        End If

        If blnEncrypt Then

            ' Is the file too big?  A Base64 encrypted file
            ' is approximately one third larger than the
            ' original and may exceed limits.
            '
            ' ex:  6973030 = MB_5 + Fix(MB_5 * 0.33)
            '      6973030 = 5242880 + 1730150
            If mcurOrigLength > CCur(MB_5 + Fix(MB_5 * 0.33)) Then

                InfoMsg "File:  " & strSource & vbNewLine & vbNewLine & _
                        "File size exceeds " & Format$(MB_5, "#,##0") & _
                        " bytes after encryption." & vbNewLine & _
                        "Base64 is not designed to process large files." & _
                        vbNewLine & vbNewLine & _
                        "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                gblnStopProcessing = True
                GoTo Base64_File_CleanUp
            End If

        End If

        mcurMaxSize = mcurOrigLength   ' Save file size for progress calcs

        mcurAccrued = 0@   ' Init accrued progress
        curProcessed = 0@
        curFilePos1 = 0@   ' Reset pointers to first position in file
        curFilePos2 = 0@

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile1) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If

        strTempFile = CreateTempFile()   ' Create a temp target file

        ' Open target file
        If Not .OpenReadWrite(strTempFile, hFile2) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If

        ReDim abytData(mcurOrigLength - 1)   ' Size receiving array

        ' Read source file
        If Not .API_ReadFile(hFile1, curFilePos1, abytData()) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If

        If blnEncrypt Then

            ' Encrypt data
            If Not Base64Encode(abytData()) Then
                gblnStopProcessing = True
                GoTo Base64_File_CleanUp
            End If

        Else

            ' Decrypt data
            If Not Base64Decode(abytData()) Then
                gblnStopProcessing = True
                GoTo Base64_File_CleanUp
            End If

        End If

        ' Write to target file
        If Not .API_WriteFile(hFile2, curFilePos2, abytData()) Then
            gblnStopProcessing = True
            GoTo Base64_File_CleanUp
        End If

        .API_CloseFile hFile1   ' Close all files opened by this routine
        .API_CloseFile hFile2
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
    Base64_File = True        ' Good finish

Base64_File_CleanUp:
    With objBigFiles
        .API_CloseFile hFile1   ' Verify all open files have been closed
        .API_CloseFile hFile2
    End With

    Set objBigFiles = Nothing   ' Free class object from memory
    Erase abytData()            ' Always empty arrays when not needed

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

Base64_File_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume Base64_File_CleanUp

End Function

' ***************************************************************************
' Routine:       CryptoAPI_String
'
' Description:   Function to encrypt/decrypt a string of data. 
' 
' Syntax:        abytData() = StrConv("abc", vbFromUnicode)             ' Convert string data to byte array
'                objCryptoAPI.CryptoAPI_String abytFile(), blnEncrypt   ' Encrypt/Decrypt data
' 
' Parameters:    abytData() - String data in byte array
'                blnEncrypt - Boolean flag designating to encrypt data
'                             TRUE - Encrypt data
'                             FALSE - Decrypt data
'
' Returns:       TRUE - Successful completion
'                FALSE - An error occurred
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function CryptoAPI_String(ByRef abytData() As Byte, _
                                 ByVal blnEncrypt As Boolean) As Boolean

    Dim lngFinalPass  As Long
    Dim curDataLength As Currency
    Dim abytSize()    As Byte
    Dim abytTemp()    As Byte

    Const ROUTINE_NAME As String = "CryptoAPI_String"

    On Error GoTo CryptoAPI_String_Error

    CryptoAPI_String = False   ' Preset to FALSE

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

    Erase abytSize()                           ' Always start with empty arrays
    mcurMaxSize = CCur(UBound(abytData)) + 1   ' Save string length for progress calcs
    mcurAccrued = 0@                           ' Init accrued progress
    lngFinalPass = 1                           ' String data only gets one pass

    ' Encrypt the data
    If blnEncrypt Then

        ' See if there is any data
        If UBound(abytData) < 1 Then

            InfoMsg "There is no data to encrypt." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

            gblnStopProcessing = True
            GoTo CryptoAPI_String_CleanUp

        End If

        ' Get string length for progress display
        mcurOrigLength = CCur(UBound(abytData) + 1)

        ' Text boxes have limitations as to the largest amount of data
        ' they can hold, based on the operating system in use. In any
        ' application, it is prudent to first determine both the
        ' operating system and the size of the file being loaded. The
        ' limit for Windows NT-based systems the limit is 64k.
        ' Applications requiring the loading of files or portions of
        ' files exceeding these maximums should use the Rich Textbox
        ' control. You can then change the maximum string size accordingly.
        '
        ' 26-Dec-2011  Randy Birch
        ' http://vbnet.mvps.org/code/faq/fileloadtext.htm
        If mcurOrigLength > kB_64 Then

            InfoMsg "This data string exceeds " & Format$(kB_64, "#,##0") & " bytes and" & vbNewLine & _
                    "will not fit inside a text box.  Place data" & vbNewLine & _
                    "inside a file and then encrypt the file." & vbNewLine & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

            gblnStopProcessing = True
            GoTo CryptoAPI_String_CleanUp

        End If

        ReDim abytTemp(mcurOrigLength)                        ' Readjust array to hold data
        CopyMemory abytTemp(0), abytData(0), mcurOrigLength   ' Copy input data to work array
        ReDim Preserve abytTemp(mcurOrigLength - 1)           ' Drop trailing null

        Select Case mlngCipherMethod

               Case 0 To 6
                    If EncryptData(abytTemp(), lngFinalPass) Then

                        ReDim abytSize(8)                            ' Size array to hold data length
                        CopyMemory abytSize(0), mcurOrigLength, 8&   ' Copy original length to temp array
                        ReDim Preserve abytSize(7)                   ' Drop trailing null value
                        MixAppendedData abytSize()                   ' Encrypt data to be appended

                        curDataLength = UBound(abytTemp)             ' Data length after encryption
                        ReDim Preserve abytTemp(curDataLength + 8)   ' Resize to hold original length

                        ' Append original data length to end of encrypted data
                        CopyMemory abytTemp(curDataLength + 1), abytSize(0), 8&
                        ReDim abytData(UBound(abytTemp) + 1)
                        CopyMemory abytData(0), abytTemp(0), (UBound(abytTemp) + 1)
                        ReDim Preserve abytData(UBound(abytTemp))

                        CryptoAPI_String = True   ' Good finish
                    Else
                        gblnStopProcessing = True
                    End If

               Case 7   ' Base64
                    If Base64Encode(abytTemp()) Then
                        Erase abytData()
                        abytData() = abytTemp()
                        CryptoAPI_String = True   ' Good finish
                    Else
                        gblnStopProcessing = True
                    End If
        End Select

    Else

        ' Decrypt the data
        Select Case mlngCipherMethod

               Case 0 To 6
                    ReDim abytSize(8)
                    CopyMemory abytSize(0), abytData(UBound(abytData) - 7), 8&   ' Capture original data length
                    ReDim Preserve abytData(UBound(abytData) - 8)                ' Adjust data length
                    ReDim Preserve abytSize(7)                                   ' Adjust size array
                    MixAppendedData abytSize()                                   ' Decrypt appended data
                    CopyMemory mcurOrigLength, abytSize(0), 8&                   ' Save original data length

                    ' Must have data and cannot exceed size of textbox.
                    ' Text boxes have limitations as to the largest amount of data
                    ' they can hold, based on the operating system in use. In any
                    ' application, it is prudent to first determine both the
                    ' operating system and the size of the file being loaded. The
                    ' limit for Windows NT-based systems the limit is 64k.
                    ' Applications requiring the loading of files or portions of
                    ' files exceeding these maximums should use the Rich Textbox
                    ' control. You can then change the maximum string size accordingly.
                    '
                    ' 26-Dec-2011  Randy Birch
                    ' http://vbnet.mvps.org/code/faq/fileloadtext.htm
                    If (UBound(abytData) < 1) Or _
                       (mcurOrigLength > kB_64) Then

                        InfoMsg "This data cannot be Decrypted, it is" & vbNewLine & _
                                "either not Encrypted, too large or corrupt." & vbNewLine & vbNewLine & _
                                "Original size:  " & Format$(mcurOrigLength, "#,##0") & " bytes" & vbNewLine & _
                                "Max allowed size:  " & Format$(kB_64, "#,##0") & " bytes" & _
                                vbNewLine & vbNewLine & _
                                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                        gblnStopProcessing = True
                        GoTo CryptoAPI_String_CleanUp

                    End If

                    ' Decrypt the data
                    If DecryptData(abytData(), lngFinalPass) Then

                        ReDim abytTemp(mcurOrigLength)
                        CopyMemory abytTemp(0), abytData(0), mcurOrigLength

                        ReDim abytData(mcurOrigLength)
                        CopyMemory abytData(0), abytTemp(0), mcurOrigLength
                        ReDim Preserve abytData(mcurOrigLength - 1)           ' Resize data to exact length
                        curDataLength = CCur(UBound(abytData) + 1)            ' data length after decryption

                        ' Compare data length
                        If curDataLength <> mcurOrigLength Then

                            InfoMsg "Encrypted data string is corrupted.  Cannot Decrypt." & _
                                    vbNewLine & vbNewLine & _
                                    "Current size:  " & Format$(curDataLength, "#,##0") & vbNewLine & _
                                    "Original size:  " & Format$(mcurOrigLength, "#,##0") & _
                                    vbNewLine & vbNewLine & _
                                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                            gblnStopProcessing = True
                        Else
                            CryptoAPI_String = True   ' Good finish
                        End If
                    Else
                        gblnStopProcessing = True
                    End If

               Case 7   ' Base64
                    If Base64Decode(abytData()) Then
                        CryptoAPI_String = True   ' Good finish
                    Else
                        gblnStopProcessing = True
                    End If
        End Select
    End If

CryptoAPI_String_CleanUp:
    ' Destroy hash object
    If mlngHashObject <> 0 Then
        Call CryptDestroyHash(mlngHashObject)
        mlngHashObject = 0
    End If

    ' Destroy session key
    If mlngSessionKey <> 0 Then
        Call CryptDestroyKey(mlngSessionKey)
        mlngSessionKey = 0
    End If

    ReleaseProvider    ' Release provider handle
    Erase abytSize()   ' Always empty arrays when not needed
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

CryptoAPI_String_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume CryptoAPI_String_CleanUp

End Function


' ***************************************************************************
' ****              Internal Functions and Procedures                    ****
' ***************************************************************************

Private Function GetSessionKey() As Boolean

    ' Called by EncryptData()
    '           DecryptData()

    Dim strPassword As String

    Const ROUTINE_NAME As String = "GetSessionKey"

    GetSessionKey = False   ' Preset to FALSE
    mlngSessionKey = 0
    mlngHashObject = 0
    strPassword = vbNullString

    If mlngProviderHandle < 1 Then
        GetProviderHandle   ' Capture provider handle
    End If

    ' Something went wrong, time to leave
    If mlngProviderHandle < 1 Then
        GoTo GetSessionKey_CleanUp
    End If

    mlngHashAlgo = GetHashSelection()    ' Check validity of hash selection
    mlngCipherAlgo = GetCipherMethod()   ' Check validity of cipher selection

    ' Create an empty hash object. The CryptCreateHash
    ' function initiates the hashing of a stream of data.
    ' It creates and returns to the calling application
    ' a handle to a CSP hash object. This handle is used
    ' in subsequent calls to CryptHashData to hash session
    ' keys and other streams of data.
    If Not CBool(CryptCreateHash(mlngProviderHandle, mlngHashAlgo, _
                                 ByVal 0&, ByVal 0&, mlngHashObject)) Then

        InfoMsg "Error: " & CStr(GetLastError) & " Failed to create a hash object." & _
                vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME

        gblnStopProcessing = True
        GoTo GetSessionKey_CleanUp
    End If

    ' Convert password to uppercase string data
    strPassword = UCase$(ByteArrayToString(mabytPassword()))

    ' If AES-nnn (Rijndael) cipher used then
    ' capture first 16-32 bytes of password
    Select Case mlngCipherMethod
           Case 4, 5, 6
                ' Capture only what is needed
                '    16 bytes = 128 \ 8
                '    24 bytes = 192 \ 8
                '    32 bytes = 256 \ 8
                strPassword = Left$(strPassword, (mlngKeyLength \ 8))
    End Select

    ' Hash in password and create hash object
    If Not CBool(CryptHashData(mlngHashObject, ByVal strPassword, _
                               Len(strPassword), ByVal 0&)) Then

        InfoMsg "Error: " & CStr(GetLastError) & " Failed to hash in the password.", _
                vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME

        gblnStopProcessing = True
        GoTo GetSessionKey_CleanUp
    End If

    ' Create session key based on cipher ID and hash object (AES supported)
    If Not CBool(CryptDeriveKey(mlngProviderHandle, mlngCipherAlgo, _
                                mlngHashObject, ByVal CRYPT_NO_SALT, mlngSessionKey)) Then

        InfoMsg "Error: " & CStr(GetLastError) & " Failed to create a session key.", _
                vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME

        gblnStopProcessing = True
        GoTo GetSessionKey_CleanUp
    End If

    GetSessionKey = True   ' Successful finish

GetSessionKey_CleanUp:
    strPassword = vbNullString
    On Error GoTo 0

End Function

Private Function EncryptData(ByRef abytData() As Byte, _
                             ByVal lngFinalPass As Long) As Boolean

    ' Called by CryptoAPI_File()
    '           CryptoAPI_String()

    Dim strData      As String
    Dim strBuffer    As String
    Dim lngDataLen   As Long
    Dim lngProgress  As Long
    Dim lngBufferLen As Long

    Const ROUTINE_NAME As String = "EncryptData"

    On Error GoTo EncryptData_Error

    EncryptData = False        ' Preset to FALSE
    strData = vbNullString     ' Empty string variables
    strBuffer = vbNullString

    lngDataLen = UBound(abytData) + 1   ' Size of incoming array

    ' Data length must be a multiple of 16 bytes
    If (lngDataLen Mod BLOCK_SIZE <> 0) Then
        lngDataLen = (lngDataLen + BLOCK_SIZE) - (lngDataLen Mod BLOCK_SIZE)
        ReDim Preserve abytData(lngDataLen - 1)
    End If

    strData = ByteArrayToString(abytData())   ' Convert byte array to string
    lngBufferLen = lngDataLen * 2             ' Double work area

    ' See if AES-nnn (Rijndael) cipher used
    Select Case mlngCipherMethod
           Case 4, 5, 6
                ' Pad output buffer for encryption
                '    16 bytes = 128 \ 8
                '    24 bytes = 192 \ 8
                '    32 bytes = 256 \ 8
                lngBufferLen = lngBufferLen + (mlngBlockSize \ 8)
    End Select

    strBuffer = String$(lngBufferLen, 0)   ' Preload with null values
    LSet strBuffer = strData               ' Forse data to left-most position

    ' Create a session key
    If Not GetSessionKey() Then
        GoTo EncryptData_CleanUp
    End If

    ' Encrypt input data
    If CryptEncrypt(mlngSessionKey, ByVal 0&, lngFinalPass, _
                    ByVal 0&, strBuffer, lngDataLen, lngBufferLen) = 0 Then

        ' Something is wrong
        If GetLastError > 0 Then

            InfoMsg "Error: " & CStr(GetLastError) & " Failed to encrypt data.", _
                    vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
        Else

            ' User may not have full Aministrator rights
            InfoMsg "This application did not function correctly possibly due to" & vbNewLine & _
                    "current logged in user may not have sufficient rights." & vbNewLine & vbNewLine & _
                    "If running Windows 8 or newer then Microsoft has split" & vbNewLine & _
                    "Administrator authority via User Access Control (UAC)." & vbNewLine & vbNewLine & _
                    "Recommended action:" & vbNewLine & _
                    vbTab & "1.  Close this application" & vbNewLine & _
                    vbTab & "2.  Right mouse click executable name" & vbNewLine & _
                    vbTab & "3.  Select option to 'Run as Administrator'" & vbNewLine & _
                    vbTab & vbTab & "or" & vbNewLine & _
                    vbTab & "4.  Select 'Properties>>Compatibility' tab and check" & vbNewLine & _
                    vbTab & "the box to 'Run this program as an administrator'." & vbNewLine & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        End If

        GoTo EncryptData_CleanUp
    End If

    Erase abytData()                           ' Empty return array
    strBuffer = Left$(strBuffer, lngDataLen)   ' Capture encrypted data
    lngDataLen = Len(strBuffer)                ' Capture data length
    abytData() = StringToByteArray(strBuffer)  ' Convert data to byte array

    ' Update progress bar. Separate
    ' steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLen)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)

    EncryptData = True   ' Successful finish

EncryptData_CleanUp:
    ' Destroy session key
    If mlngSessionKey <> 0 Then
        Call CryptDestroyKey(mlngSessionKey)
        mlngSessionKey = 0
    End If

    strData = vbNullString     ' Empty string variables
    strBuffer = vbNullString
    On Error GoTo 0            ' Nullify this error trap
    Exit Function

EncryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume EncryptData_CleanUp

End Function

Private Function DecryptData(ByRef abytData() As Byte, _
                             ByVal lngFinalPass As Long) As Boolean

    ' Called by CryptoAPI_File()
    '           CryptoAPI_String()

    Dim strData     As String
    Dim strBuffer   As String
    Dim lngDataLen  As Long
    Dim lngProgress As Long

    Const ROUTINE_NAME As String = "DecryptData"

    On Error GoTo DecryptData_Error

    DecryptData = False        ' Preset to FALSE
    strData = vbNullString     ' Empty string variables
    strBuffer = vbNullString

    strData = ByteArrayToString(abytData())   ' Convert byte array to string
    lngDataLen = Len(strData)                 ' Capture input string length
    strBuffer = String$(lngDataLen, 0)        ' Preload with nulls
    LSet strBuffer = strData                  ' Force data to left-most position

    ' Create a session key
    If Not GetSessionKey() Then
        GoTo DecryptData_CleanUp
    End If

    ' Decrypt data
    If CryptDecrypt(mlngSessionKey, ByVal 0&, ByVal lngFinalPass, _
                    ByVal 0&, strBuffer, lngDataLen) = 0 Then

        ' Something is wrong
        If GetLastError <> 0 Then

            InfoMsg "Error: " & CStr(GetLastError) & " Failed to decrypt data.", _
                     vbNewLine & vbNewLine & MODULE_NAME & ":" & ROUTINE_NAME
        Else

            ' Wrong password or user may not have full Aministrator rights
            InfoMsg "An incorrect password may have been used." & vbNewLine & _
                    vbTab & "or" & vbNewLine & _
                    "This application may not function correctly possibly due to" & vbNewLine & _
                    "current logged in user may not have sufficient rights." & vbNewLine & vbNewLine & _
                    "Recommended action:" & vbNewLine & _
                    vbTab & "1.  Close and restart this application." & vbNewLine & _
                    vbTab & vbTab & "or" & vbNewLine & _
                    vbTab & "1.  Close this application" & vbNewLine & _
                    vbTab & "2.  Right mouse click executable name" & vbNewLine & _
                    vbTab & "3.  Select option to 'Run as Administrator'" & vbNewLine & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
        End If

        GoTo DecryptData_CleanUp

    End If

    Erase abytData()                           ' Empty return array
    strBuffer = Left$(strBuffer, lngDataLen)   ' Capture decrypted data
    abytData() = StringToByteArray(strBuffer)  ' Convert data to byte array

    ' Update progress bar. Separate
    ' steps for easier debugging
    mcurAccrued = mcurAccrued + CCur(lngDataLen)
    lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
    RaiseEvent CipherProgress(lngProgress)

    DecryptData = True   ' Successful finish

DecryptData_CleanUp:
    ' Destroy session key
    If mlngSessionKey <> 0 Then
        Call CryptDestroyKey(mlngSessionKey)
        mlngSessionKey = 0
    End If

    strData = vbNullString     ' Empty string variables
    strBuffer = vbNullString
    On Error GoTo 0            ' Nullify this error trap
    Exit Function

DecryptData_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume DecryptData_CleanUp

End Function

' ***************************************************************************
' Routine:       Base64Encode
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
' Parameters:   abytData() - Byte array holding data to be encoded
'
' Returns:      TRUE if successful else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function Base64Encode(ByRef abytData() As Byte) As Boolean

    Dim strData    As String
    Dim lngDataLen As Long

    Const ROUTINE_NAME As String = "Base64Encode"

    On Error GoTo Base64Encode_Error

    Base64Encode = False   ' Preset to FALSE

    ' Determine Base64 output String length required
    CryptBinaryToString abytData(0), UBound(abytData) + 1, _
                        CRYPT_STRING_BASE64, _
                        StrPtr(vbNullString), lngDataLen

    strData = String$(lngDataLen, 0)   ' Preload with nulls

    ' Convert binary to Base64
    If CryptBinaryToString(abytData(0), UBound(abytData) + 1, _
                           CRYPT_STRING_BASE64, _
                           StrPtr(strData), lngDataLen) = 0 Then

        InfoMsg "Failed to encode data to Base64." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    Else
        Erase abytData()                          ' Empty original array
        strData = Left$(strData, lngDataLen)      ' Capture actual length
        abytData() = StringToByteArray(strData)   ' Convert string to byte array
        Base64Encode = True                       ' Successful finish
    End If

Base64Encode_CleanUp:
    strData = vbNullString   ' Empty variable
    On Error GoTo 0          ' Nullify this error trap
    Exit Function

Base64Encode_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Base64Encode_CleanUp

End Function

' ***************************************************************************
' Routine:       Base64Decode
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
' Parameters:   abytData() - Byte array holding data to be decoded
'
' Returns:      TRUE if successful else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 27-Jul-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function Base64Decode(ByRef abytData() As Byte) As Boolean

    Dim strData       As String
    Dim lngDataLen    As Long
    Dim lngActualUsed As Long
    Dim abytTemp()    As Byte

    Const ROUTINE_NAME As String = "Base64Decode"

    On Error GoTo Base64Decode_Error

    Base64Decode = False    ' Preset to FALSE
    strData = vbNullString  ' Empty variable

    strData = ByteArrayToString(abytData())   ' Convert to string data

    ' Determine output buffer length required
    CryptStringToBinary StrPtr(strData), Len(strData), CRYPT_STRING_BASE64, _
                        StrPtr(vbNullString), lngDataLen, 0&, lngActualUsed

    ReDim abytTemp(lngDataLen - 1)   ' Size temp work array

    ' Convert Base64 to binary
    If CryptStringToBinary(StrPtr(strData), Len(strData), CRYPT_STRING_BASE64, _
                           VarPtr(abytTemp(0)), lngDataLen, 0&, lngActualUsed) = 0 Then

        InfoMsg "Failed to decode Base64 data to string." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME
    Else
        ReDim abytData(lngDataLen)                        ' Resize original array
        CopyMemory abytData(0), abytTemp(0), lngDataLen   ' Transfer data to original array
        ReDim Preserve abytData(lngDataLen - 1)           ' Drop trailing null
        Base64Decode = True                               ' Successful finish
    End If

Base64Decode_CleanUp:
    Erase abytTemp()   ' Empty temp work array
    On Error GoTo 0    ' Nullify this error trap
    Exit Function

Base64Decode_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume Base64Decode_CleanUp

End Function

' **************************************************************************
' Routine:       GetBlockSize
'
' Description:   Determines amount of the data to be processed.
'
' Parameters:    curAmtLeft - Amount of data left
'
' Returns:       New block size as a long integer
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

' ***************************************************************************
' Procedure:     GetHashSelection
'
' Description:   User passes a numeric value, usually 0-6, and this routine
'                will determine the hash selection.  No data needs to be
'                passed to the HashLength parameter as this is output only.
'                If the AES provider is not available then the parameter
'                passed will be calculated to use a lesser hash algorithm.
'
' Parameters:    lngHashMethod - Input - Selected hash algorithm
'
' Returns:       Five digit code representing appropriate hash selection
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function GetHashSelection() As Long

    ' Called by GetSessionKey()

    Dim lngHashMethod As enumCryptoAPI_HashAlgorithms

    lngHashMethod = mlngHashMethod

    Select Case mblnAESProvider

           Case True   ' Strong (AES) provider available
                Select Case lngHashMethod
                       Case eCryptoAPI_MD2:    GetHashSelection = CALG_MD2       ' 0
                       Case eCryptoAPI_MD4:    GetHashSelection = CALG_MD4       ' 1
                       Case eCryptoAPI_MD5:    GetHashSelection = CALG_MD5       ' 2
                       Case eCryptoAPI_SHA1:   GetHashSelection = CALG_SHA1      ' 3
                       Case eCryptoAPI_SHA256: GetHashSelection = CALG_SHA_256   ' 4
                       Case eCryptoAPI_SHA384: GetHashSelection = CALG_SHA_384   ' 5
                       Case eCryptoAPI_SHA512: GetHashSelection = CALG_SHA_512   ' 6
                       Case Else:              GetHashSelection = CALG_SHA_256   ' 4 Default
                End Select

           Case False  ' Default or Enhanced provider available
                If lngHashMethod > eCryptoAPI_SHA1 Then
                    lngHashMethod = eCryptoAPI_SHA1
                End If

                Select Case lngHashMethod
                       Case eCryptoAPI_MD2:  GetHashSelection = CALG_MD2    ' 0
                       Case eCryptoAPI_MD4:  GetHashSelection = CALG_MD4    ' 1
                       Case eCryptoAPI_MD5:  GetHashSelection = CALG_MD5    ' 2
                       Case eCryptoAPI_SHA1: GetHashSelection = CALG_SHA1   ' 3
                       Case Else:            GetHashSelection = CALG_SHA1   ' 3 Default
                End Select
    End Select

End Function

' ***************************************************************************
' Procedure:     CipherSelection
'
' Description:   User passes a numeric value, usually 0-8, and this routine
'                will determine the cipher selection.  If the AES provider
'                is not available then the parameter passed will be
'                calculated to use a lesser hash algorithm.
'
' Parameters:    lngCipherMethod - Input - Selected hash algorithm
'
' Returns:       Five digit code representing appropriate cipher selection
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function GetCipherMethod() As Long

    ' Call by GetSessionKey()

    Dim lngCipherMethod As enumCryptoAPI_CipherAlgorithms

    lngCipherMethod = mlngCipherMethod

    Select Case mblnAESProvider

           Case True   ' Strong (AES) provider available
                Select Case lngCipherMethod
                       Case eCryptoAPI_RC2:      GetCipherMethod = CALG_RC2        ' 0
                       Case eCryptoAPI_RC4:      GetCipherMethod = CALG_RC4        ' 1
                       Case eCryptoAPI_DES:      GetCipherMethod = CALG_DES        ' 2
                       Case eCryptoAPI_3DES:     GetCipherMethod = CALG_3DES       ' 3
                       Case eCryptoAPI_AES_128:  GetCipherMethod = CALG_AES_128    ' 4
                       Case eCryptoAPI_AES_192:  GetCipherMethod = CALG_AES_192    ' 5
                       Case eCryptoAPI_AES_256:  GetCipherMethod = CALG_AES_256    ' 6
                       Case Else:                GetCipherMethod = CALG_RC4        ' 1 Default
                End Select

           Case False  ' Default or Enhanced provider available
                If lngCipherMethod > eCryptoAPI_RC4 Then
                    lngCipherMethod = eCryptoAPI_RC4
                End If

                Select Case lngCipherMethod
                       Case eCryptoAPI_RC2: GetCipherMethod = CALG_RC2   ' 0
                       Case eCryptoAPI_RC4: GetCipherMethod = CALG_RC4   ' 1
                       Case eCryptoAPI_DES: GetCipherMethod = CALG_DES   ' 2
                       Case Else:           GetCipherMethod = CALG_RC4   ' 1 Default
                End Select
    End Select

End Function

' ***************************************************************************
' Routine:       GetProviderHandle
'
' Description:   This is a generic routine to capture the MS provider
'                handle to see if any enhanced encryption capabilities are
'                available.
'
'                The AES (Advanced Encryption Standard) cipher algorithm
'                and SHA2 family of hash algorithms were not available until
'                Windows XP with SP3 installed.  The container provider for
'                Windows XP SP3 was named the same as later versions except
'                with the word prototype appended.  (ex:  "(Prototype)")
'                See constant MS_ENH_RSA_AES_PROV_XP in Declare section of
'                this module.
'
'                Windows XP has supported SHA-256, SHA-384 and SHA-512
'                since the release of Service Pack 3 in 2008, and Windows
'                Server 2003 can also support SHA-2 if the KB938397 hotfix
'                has been installed.
'
' Parameters:    lngProvider - Which provider to search for
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Feb-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 10-Feb-2010  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for bringing to my attention the
'              need to update this routine.  Rewrote this routine to test
'              for availability of SHA2 hash functionality
' 04-Mar-2010  Kenneth Ives  kenaso@tx.rr.com
'              Changed code structure to reduce number of GoTo statements.
' 24-May-2015  Kenneth Ives  kenaso@tx.rr.com
'              Added ability to determine which type of provider to look for
' ***************************************************************************
Private Sub GetProviderHandle()

    ' Called by GetSessionKey()
    '           AES_Ready() Property

    Dim lngIndex       As Long
    Dim strContainer   As String   ' Container value
    Dim astrProvider() As String   ' List of providers

    Const ROUTINE_NAME As String = "GetProviderHandle"

    On Error GoTo GetProviderHandle_Error

    ReleaseProvider   ' Release hold on any provider handle

    ReDim astrProvider(5)   ' Size provider array

    ' Load provider names
    astrProvider(0) = MS_ENH_RSA_AES_PROV & vbNullChar      ' Windows Vista or newer
    astrProvider(1) = MS_ENH_RSA_AES_PROV_XP & vbNullChar   ' Windows XP with service pack 3
    astrProvider(2) = MS_ENHANCED_PROV & vbNullChar         ' Enhanced provider
    astrProvider(3) = MS_DEFAULT_PROV & vbNullChar          ' Base default provider
    astrProvider(4) = vbNullChar

    ' Set search parameters
    For lngIndex = 0 To UBound(astrProvider) - 1

        strContainer = vbNullChar

        Select Case lngIndex
               Case 0 To 3
                    ' Attempt to acquire a provider handle.
                    ' Failure is a return value of zero.
                    If CBool(CryptAcquireContext(mlngProviderHandle, _
                                                 ByVal strContainer, ByVal astrProvider(lngIndex), _
                                                 PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
                        If lngIndex < 2 Then
                            mblnAESProvider = True   ' Set property value (AES capability is available)
                        End If

                        Exit For   ' Exit FOR..NEXT loop
                    End If

               Case Else
                    ' If no luck acquiring a provider handle then create
                    ' default provider using current user's logon ID.
                    ' Failure is a return value of zero.
                    If CBool(CryptAcquireContext(mlngProviderHandle, _
                                                 vbNullString, vbNullString, _
                                                 PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then

                        ' Successful aquisition of a provider handle
                    Else
                        ' Serious problem - Stop processing
                        InfoMsg "An unknown error occurred while retrieving" & vbNewLine & _
                                "CryptoAPI provider handle." & _
                                vbNewLine & vbNewLine & _
                                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

                        ReleaseProvider             ' Verify provider handle is clear
                        gblnStopProcessing = True   ' Time to stop processing

                    End If     ' Create a provider

                    Exit For   ' Exit FOR..NEXT loop
        End Select

    Next lngIndex

GetProviderHandle_CleanUp:
    Erase astrProvider()   ' Empty array
    On Error GoTo 0        ' Nullify this error trap
    Exit Sub

GetProviderHandle_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ReleaseProvider                   ' Release provider handle
    gblnStopProcessing = True         ' Time to stop processing
    Resume GetProviderHandle_CleanUp

End Sub

Private Sub ReleaseProvider()

    ' Called by GetProviderHandle()
    '           Class_Terminate()

    If mlngProviderHandle <> 0 Then
        CryptReleaseContext mlngProviderHandle, 0&
    End If

    mblnAESProvider = False  ' Set property value (AES capability not available)
    mlngProviderHandle = 0   ' Numeric value depicting provider handle

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
    ' Enum enumCryptoAPI_HashAlgorithms
    Const eCryptoAPI_MD2     As Long = 0&
    Const eCryptoAPI_MD4     As Long = 1&
    Const eCryptoAPI_MD5     As Long = 2&
    Const eCryptoAPI_SHA1    As Long = 3&
    Const eCryptoAPI_SHA256  As Long = 4&
    Const eCryptoAPI_SHA384  As Long = 5&
    Const eCryptoAPI_SHA512  As Long = 6&

    ' Enum enumCryptoAPI_CipherAlgorithms
    Const eCryptoAPI_RC2     As Long = 0&
    Const eCryptoAPI_RC4     As Long = 1&
    Const eCryptoAPI_DES     As Long = 2&
    Const eCryptoAPI_3DES    As Long = 3&
    Const eCryptoAPI_AES_128 As Long = 4&
    Const eCryptoAPI_AES_192 As Long = 5&
    Const eCryptoAPI_AES_256 As Long = 6&
    Const eCryptoAPI_BASE64  As Long = 7&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ReleaseProvider          ' Release provider handle
    StopProcessing = False   ' Preset property to FALSE
    CreateNewFile = True     ' Default is to create a new output file
    HashMethod = 4           ' Default is SHA-256
    CipherMethod = 1         ' Default is RC4

    ' Ignored unless AES-nnn cipher
    BlockSize = 128          ' Default AES_128 (16 bytes)
    KeyLength = 128          ' Default AES_128 (16 bytes)

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ReleaseProvider   ' Release provider handle

End Sub



