VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAPI_Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        cAPI_Hash  (clsAPI_Hash.cls)
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
' Description:   This is a class which accesses Microsoft's CryptoAPI hash
'                algorithms.  These include:
'
'                    MD2  MD4  MD5  SHA-1  SHA-256  SHA-384  SHA-512
'
'                The Secure Hash Algorithm (SHA) is required for use with
'                the Digital Signature Algorithm (DSA) as specified in the
'                Digital Signature Standard (DSS) and whenever a secure
'                hash algorithm is required for federal applications.  For
'                a message of length < 2^64 bits, this algorithm produces a
'                condensed representation of the message called a
'                message digest. The message digest is used during
'                generation of a signature for the message.  This
'                also used to compute a message digest for the received
'                version of the message during the process of verifying the
'                signature.  Any change to the message in transit will,
'                with very high probability, result in a different message
'                digest, and the signature will fail to verify.
'
'                The MD5 algorithm is an extension of the MD4 message-digest
'                algorithm 1,2. MD5 is slightly slower than MD4, but is more
'                "conservative" in design. MD5 was designed because it was
'                felt that MD4 was perhaps being adopted for use more quickly
'                than justified by the existing critical review; because MD4
'                was designed to be exceptionally fast, it is "at the edge"
'                in terms of risking successful cryptanalytic attack. MD5
'                backs off a bit, giving up a little in speed for a much
'                greater likelihood of ultimate security. It incorporates
'                some suggestions made by various reviewers, and contains
'                additional optimizations. The MD5 algorithm is being placed
'                in the public domain for review and possible adoption as a
'                standard.
'
'                These algorithms have been tested to be accurate in
'                accordance with FIPS-180-2, FIPS-180-3 and FIPS-180-4
'                publications.
'
'                SHA-256 Rollout (Nov 19, 2014)
'                16-Jan-2015  CAs will be forbidden to issue SHA-1 certificates
'                             that expire past December 31, 2016
'                01-Jan_2016  Microsoft will cease trusting Code Signing
'                             Certificates using SHA-1
'                01-Jan-2017  Microsoft will cease trusting SSL Certificates
'                             using SHA-1
'
' REFERENCE:     Microsoft Windows SHA-3 Comments
'                http://bradconte.com/files/misc/Windows SHA-3 comments.pdf
'
'                The Cryptography API, or How to Keep a Secret
'                http://msdn.microsoft.com/en-us/library/ms867086.aspx
'
'                CryptoAPI Cryptographic Service Providers
'                http://msdn.microsoft.com/en-us/library/bb931357(VS.85).aspx
'
'                Secure Hash Algorithm-2 (SHA-2) - This hashing algorithm was
'                developed as a successor to SHA-1 by the National Institute of
'                Standards and Technology (NIST) and the National Security Agency
'                (NSA). It has four variants SHA-224, SHA-256, SHA-384, and SHA-512
'                which are named according to the number of bits in their outputs.
'                Of these, SHA-256, SHA-384, and SHA-512 are implemented in the
'                Microsoft AES Cryptographic Provider.
'                http://msdn.microsoft.com/en-us/library/windows/desktop/aa382459(v=vs.85).aspx
'
'                SHA-2 support on MS Windows (23 Jan 2009)
'                Paraphrasing:  Regarding SHA-224 support, SHA-224 offers
'                less security than SHA-256 but takes the same amount of
'                resources.  Also SHA-224 is not generally used by protocols
'                and applications.  The National Security Agency (NSA)
'                Suite B standards also does not include it.  Microsoft
'                has no plans to add it to future versions of their
'                Cryptographic Service Providers (CSP).
'                http://blogs.msdn.com/b/alejacma/archive/2009/01/23/sha-2-support-on-windows-xp.aspx
'
'                SHA-224, SHA-512/224 and SHA-512/256 have not yet been
'                implemented into the Microsoft crypto suite of hashes.
'
'                NIST (National Institute of Standards and Technology)
'                FIPS (Federal Information Processing Standards Publication)
'                SP (Special Publications)
'                http://csrc.nist.gov/publications/PubsFIPS.html
'
'                FIPS 180-2 (Federal Information Processing Standards Publication)
'                dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
'                http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
'                FIPS 180-3 (Federal Information Processing Standards Publication)
'                dated Oct-2008 (supercedes FIPS 180-2)
'                http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
'                FIPS 180-4 (Federal Information Processing Standards Publication)
'                dated Mar-2012 (Supercedes FIPS-180-3)
'                http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
'
'                Examples of the implementation of the secure hash algorithms
'                SHA-1, SHA-224, SHA-256, SHA-384, SHA-512 can be found at:
'                http://csrc.nist.gov/groups/ST/toolkit/examples.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/Examples/SHA2_Additional.pdf
'
'                Aaron Gifford's additional test vectors
'                http://www.adg.us/computers/sha.html
'
' Add'l Info:    MD4 Message-Digest Algorithm has been compromised at the rump
'                session of Crypto 2004 it was announced that Xiaoyun Wang,
'                Dengguo Feng, Xuejia Lai and Hongbo Yu found collisions for
'                MD4, MD5, RIPEMD, and the 128-bit version of HAVAL.
'                http://eprint.iacr.org/2004/199.pdf
'
'                Feb-2005:  SHA-1 has been compromised.  Recommended that
'                you do not use for password or document authentication.
'                http://www.schneier.com/blog/archives/2005/02/sha1_broken.html
'                http://csrc.nist.gov/groups/ST/toolkit/documents/shs/NISTHashComments-final.pdf
'
'                Mar-2005 Demonstrating a technique for finding MD5 collisions quickly.
'                Eight hours on 1.6 GHz computer.
'                http://cryptography.hyperlink.cz/md5/MD5_collisions.pdf
'
'                Jun-2005 Two researchers from the Institute for Cryptology and
'                IT-Security have generated PostScript files with identical MD5-sums
'                but entirely different (but meaningful!) content.
'                http://www.schneier.com/blog/archives/2005/06/more_md5_collis.html
'
'                March 15, 2006:  The SHA-2 family of hash functions
'                (i.e., SHA-224, SHA-256, SHA-384 and SHA-512) may be used
'                by Federal agencies for all applications using secure hash
'                algorithms. Federal agencies should stop using SHA-1 for
'                digital signatures, digital time stamping and other
'                applications that require collision resistance as soon as
'                practical, and must use the SHA-2 family of hash functions
'                for these applications after 2010. After 2010, Federal
'                agencies may use SHA-1 only for the following applications:
'                     - hash-based message authentication codes (HMACs)
'                     - key derivation functions (KDFs)
'                     - random number generators (RNGs)
'                Regardless of use, NIST encourages application and protocol
'                designers to use the SHA-2 family of hash functions for all
'                new applications and protocols.
'                http://csrc.nist.gov/groups/ST/hash/policy.html
'
'                Export Control: Certain cryptographic devices and technical
'                data regarding them are subject to Federal export controls.
'                Exports of cryptographic modules implementing this standard
'                and technical data regarding them must comply with these
'                Federal regulations and be licensed by the Bureau of Export
'                Administration of the U.S. Department of Commerce.
'                Information about export regulations is available at:
'                http://www.bis.doc.gov/index.htm
'
' AddIn tools    Callers Add-in v3.6 dtd 04-Sep-2016 by RD Edwards (RDE)
' for VB6:       Fantastic VB6 add-in to indentify if a routine calls another
'                routine or is called by other routines within a project. A must
'                have tool for any VB6 programmer.
'                http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74734&lngWId=1
'
'                NOTE:  Under Windows 10, if you have problems recognizing
'                a VB6 addin, try recompiling it directly into the System32
'                folder.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote this module
' 25-Mar-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 04-May-2012  Kenneth Ives  kenaso@tx.rr.com
'              Added new property ReturnLowercase().  Designates if returned
'              hashed data string should be in upper or lowercase format.
' 04-Aug-2013  Kenneth Ives  kenaso@tx.rr.com
'              Made two local variables module level(mlngHashLength,
'              mlngHashObject) for speed.
' 22-Nov-2013  Kenneth Ives  kenaso@tx.rr.com
'              Fixed bug in FormatOutput() routine.  Only appeared if user
'              did not want to return data as a hex string.
' 28-Dec-2014  Kenneth Ives  kenaso@tx.rr.com
'              Updated documentation
' 07-Nov-2015  Kenneth Ives  kenaso@tx.rr.com
'              - Fixed logic flaw for multiple hash iterations.
'                Updated ProcessData() routine and documentation.
'              - Removed obsolete code.
' 05-Oct-2016  Kenneth Ives  kenaso@tx.rr.com
'              Fixed logic flaw for converting output to hex.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module constants
' ***************************************************************************
  Private Const MODULE_NAME            As String = "clsAPI_Hash"
  Private Const MB_5                   As Long = &H500000     ' 5242880 bytes
  Private Const MAX_ROUNDS             As Long = 10

  ' Microsoft Hash constants
  Private Const HP_HASHVAL             As Long = 2
  Private Const ALG_TYPE_ANY           As Long = 0            ' Used by all hashes
  Private Const ALG_SID_MD2            As Long = 1            ' 16 byte hash result
  Private Const ALG_SID_MD4            As Long = 2            ' 16 byte hash result
  Private Const ALG_SID_MD5            As Long = 3            ' 16 byte hash result
  Private Const ALG_SID_SHA1           As Long = 4            ' 20 byte hash result
  Private Const ALG_SID_SHA_256        As Long = 12           ' 32 byte hash result
  Private Const ALG_SID_SHA_384        As Long = 13           ' 48 byte hash result
  Private Const ALG_SID_SHA_512        As Long = 14           ' 64 byte hash result
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
  Private Const CRYPT_NEWKEYSET        As Long = &H8          ' If needed, create a container
  Private Const CRYPT_VERIFYCONTEXT    As Long = &HF0000000   ' -268435456
  ' Verify provider names
  ' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography\Defaults\Provider\
  Private Const MS_DEFAULT_PROV        As String = "Microsoft Base Cryptographic Provider v1.0"
  Private Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"
  Private Const MS_ENH_RSA_AES_PROV    As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"
  Private Const MS_ENH_RSA_AES_PROV_XP As String = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumAPI_HashAlgorithms
      eAPI_MD2      ' 0
      eAPI_MD4      ' 1
      eAPI_MD5      ' 2
      eAPI_SHA1     ' 3
      eAPI_SHA256   ' 4
      eAPI_SHA384   ' 5
      eAPI_SHA512   ' 6
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

  ' The CryptGetHashParam function retrieves data that governs the
  ' operations of a hash object. The actual hash value can be
  ' retrieved by using this function.
  Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
          (ByVal hhash As Long, ByVal dwParam As Long, _
          ByVal pbData As String, pdwDataLen As Long, _
          ByVal dwFlags As Long) As Long

  'The CryptDestroyHash function destroys the hash object referenced
  ' by the hHash parameter. After a hash object has been destroyed,
  ' it can no longer be used.  The destruction of hash objects after
  ' their use is finished is recommended for security reasons.
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

  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'                    +-------------- Module level designator
'                    |  +----------- Data type (Currency)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m cur Accrued
' Variable name:     mcurAccrued
' ***************************************************************************
  Private mblnAESProvider     As Boolean   ' Flag denoting AES crypto provider
  Private mblnRetLowercase    As Boolean
  Private mblnReturnHexString As Boolean
  Private mcurMaxSize         As Currency
  Private mcurAccrued         As Currency
  Private mlngHashLength      As Long      ' Raw data hash length
  Private mlngHashObject      As Long
  Private mlngHashRounds      As Long
  Private mlngProviderHandle  As Long      ' CryptoAPI provider handle
  Private mlngHashMethod      As enumAPI_HashAlgorithms


' ***************************************************************************
' ****                      Events                                       ****
' ***************************************************************************

' Update progress bar
Public Event HashProgress(ByVal lngProgress As Long)


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

Public Property Let HashMethod(ByVal lngHashMethod As enumAPI_HashAlgorithms)
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  X.HashMethod = eAPI_SHA512
    '
    mlngHashMethod = lngHashMethod

End Property

Public Property Let ReturnHexString(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnHexString = True
    mblnReturnHexString = blnData
End Property

Public Property Let ReturnLowercase(ByVal blnData As Boolean)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.ReturnLowercase = True
    mblnRetLowercase = blnData
End Property

Public Property Get AES_Ready() As Boolean

    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.AES_Ready

    ' Test for a provider handle
    If mlngProviderHandle < 1 Then
        GetProviderHandle   ' Get provider handle
    End If

    AES_Ready = mblnAESProvider   ' Is advanced hashing available?

End Property

Public Property Let HashRounds(ByVal lngData As Long)
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.HashRounds = 1
    Select Case lngData
           Case 1 To MAX_ROUNDS: mlngHashRounds = lngData   ' Good selection
           Case Else:            mlngHashRounds = 1         ' Default value
    End Select

End Property


' *************************************************************************** 
' ****                      Methods                                      **** 
' *************************************************************************** 

' *************************************************************************** 
' Routine:       HashFile 
' 
' Description:   Function to create a unique hex string representation of 
'                the data passed. Can handle files greater than 2 gigabytes.
' 
' Parameters:    abytFile() - Path/filename in byte array to be hashed 
' 
' Returns:       Hashed data string in byte array 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' *************************************************************************** 
Public Function HashFile(ByRef abytFile() As Byte) As Byte()

    Dim hFile        As Long
    Dim lngProgress  As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim curFilePos   As Currency
    Dim strSource    As String
    Dim abytData()   As Byte
    Dim objBigFiles  As cBigFiles

    On Error GoTo HashFile_CleanUp

    hFile = 0
    Set objBigFiles = New cBigFiles  ' Instantiate class object

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    ' convert byte array to string
    strSource = ByteArrayToString(abytFile())

    ' Verify file exists
    If Not IsPathValid(strSource) Then
        gblnStopProcessing = True
        GoTo HashFile_CleanUp
    End If

    ReleaseProvider    ' Verify any provider handles are released
    curFilePos = 0@    ' set to first position in the file
    mcurAccrued = 0@   ' Initialize accumulators
    mcurMaxSize = 0@

    With objBigFiles
        .CalcFileSize strSource, mcurMaxSize        ' Get file size
        curAmtLeft = mcurMaxSize                    ' Save starting length
        mcurMaxSize = mcurMaxSize * mlngHashRounds  ' for progress display

        ' If zero byte file then leave
        If curAmtLeft < 1 Then
            gblnStopProcessing = True
        End If

        ' An error occurred or user opted to STOP processing
        DoEvents
        If gblnStopProcessing Then
            GoTo HashFile_CleanUp
        End If

        '*********************************
        ' Step #1  Create hash object
        '*********************************
        If Not CreateHashObject() Then
            GoTo HashFile_CleanUp
        End If

        ' Open source file
        If Not .OpenReadOnly(strSource, hFile) Then
            gblnStopProcessing = True
            GoTo HashFile_CleanUp
        End If

        ' Loop while processing data chunks
        Do
            lngBlockSize = GetBlockSize(curAmtLeft)   ' Process data in chunks
            ReDim abytData(lngBlockSize - 1)          ' Size receiving array

            curAmtLeft = curAmtLeft - CCur(lngBlockSize)     ' Adjust amount of file left to process
            mcurAccrued = mcurAccrued + CCur(lngBlockSize)   ' Current progress

            ' Read source file
            If Not .API_ReadFile(hFile, curFilePos, abytData()) Then
                gblnStopProcessing = True
                Exit Do  ' exit Do..Loop
            End If

            ' Adjust pointers accordingly
            curFilePos = curFilePos + CCur(UBound(abytData) + 1)

            '*********************************
            ' Step #2  Hash data
            '*********************************
            If Not ProcessData(abytData()) Then
                Exit Do  ' exit Do..Loop
            End If

            ' Update progress bar
            lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
            RaiseEvent HashProgress(lngProgress)
            DoEvents

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

        .API_CloseFile hFile    ' Close all files opened by this routine
    End With

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashFile_CleanUp
    End If

    '*********************************
    ' Step #3  Format hashed data and
    '          return as byte array
    '*********************************
    Erase abytData()                 ' Empty array
    abytData() = FormatOutput()      ' Convert hashed data to an array

    HashFile = abytData()            ' Return array
    RaiseEvent HashProgress(100)     ' Update progress bar

HashFile_CleanUp:
    objBigFiles.API_CloseFile hFile  ' Close file opened by this routine
    Set objBigFiles = Nothing        ' Always free objects from memory

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        RaiseEvent HashProgress(0)   ' Update progress bar
        DoEvents
    End If

    Erase abytData()   ' Empty array
    ReleaseProvider    ' Verify any provider handles are released
    On Error GoTo 0    ' Nullify this error trap

End Function

' *************************************************************************** 
' Routine:       HashString 
' 
' Description:   Function to create a unique hex string representation of 
'                the data passed.  See Readme.txt file for more information. 
' 
' Parameters:    abytData() - data string in byte array to be hashed 
' 
' Returns:       Hashed data string in byte array 
' 
' =========================================================================== 
'    DATE      NAME / eMAIL 
'              DESCRIPTION 
' -----------  -------------------------------------------------------------- 
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' *************************************************************************** 
Public Function HashString(ByRef abytData() As Byte) As Byte()

    Dim lngPointer   As Long
    Dim lngProgress  As Long
    Dim lngBlockSize As Long
    Dim curAmtLeft   As Currency
    Dim abytHash()   As Byte

    On Error GoTo HashString_CleanUp

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashString_CleanUp
    End If

    ReleaseProvider    ' Verify any provider handles are released
    mcurAccrued = 0@   ' Initialize accumulators
    mcurMaxSize = 0@
    lngPointer = 0     ' Initialize data pointer

    mcurMaxSize = CCur(UBound(abytData) + 1)    ' Capture string size
    curAmtLeft = mcurMaxSize                    ' Save starting length
    mcurMaxSize = mcurMaxSize * mlngHashRounds  ' for progress display

    ' If empty string then leave
    If curAmtLeft < 1 Then
        gblnStopProcessing = True
    End If

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        GoTo HashString_CleanUp
    End If

    '*********************************
    ' Step #1  Create hash object
    '*********************************
    If Not CreateHashObject() Then
        GoTo HashString_CleanUp
    End If

    ' Loop while processing data chunks
    Do
        lngBlockSize = GetBlockSize(curAmtLeft)                      ' Process data in chunks
        ReDim abytHash(lngBlockSize)                                 ' Size temp array
        CopyMemory abytHash(0), abytData(lngPointer), lngBlockSize   ' Load temp array
        ReDim Preserve abytHash(lngBlockSize - 1)                    ' Adjust temp array

        lngPointer = lngPointer + lngBlockSize + 1    ' Adjust data pointer accordingly
        curAmtLeft = curAmtLeft - CCur(lngBlockSize)  ' Adjust amount of data left to process

        '*********************************
        ' Step #2  Hash data
        '*********************************
        If Not ProcessData(abytHash()) Then
            Exit Do  ' exit Do..Loop
        End If

        ' Update progress bar
        lngProgress = CalcProgress(mcurAccrued, mcurMaxSize)
        RaiseEvent HashProgress(lngProgress)
        DoEvents

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
        GoTo HashString_CleanUp
    End If

    '*********************************
    ' Step #3  Format hashed data and
    '          return as byte array
    '*********************************
    Erase abytHash()                 ' Empty array
    abytHash() = FormatOutput()      ' Convert hashed data to an array

    HashString = abytHash()          ' return array
    RaiseEvent HashProgress(100)     ' Update progress bar

HashString_CleanUp:
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        RaiseEvent HashProgress(0)   ' Update progress bar
        DoEvents
    End If

    Erase abytHash()   ' Empty array
    ReleaseProvider    ' Verify any provider handles are released
    On Error GoTo 0    ' Nullify this error trap

End Function


' ***************************************************************************
' ****                  Internal functions and Procedures                ****
' ***************************************************************************

Private Function CreateHashObject() As Boolean

    '*********************************
    ' Step #1  Create hash object
    '*********************************

    ' Called by HashFile()
    '           HashString()

    Dim lngHashAlgo As Long

    Const ROUTINE_NAME As String = "CreateHashObject"

    On Error Resume Next

    lngHashAlgo = 0
    mlngHashObject = 0
    mlngHashLength = 0

    CreateHashObject = False  ' Preset to FALSE

    ' Aquire provider handle
    If mlngProviderHandle < 1 Then

        GetProviderHandle   ' Capture strongest provider handle

        ' Something went wrong, time to leave
        If mlngProviderHandle < 1 Then
            GoTo CreateHashObject_CleanUp
        End If

    End If   ' mlngProviderHandle

    ' Get hash algorithm and return length value
    HashSelection lngHashAlgo

    ' Create a hash object.
    ' API CryptCreateHash function initiates the hashing of
    ' a stream of data.  It creates and returns to the calling
    ' application a handle to a CSP hash object.  This handle
    ' is used in subsequent calls to CryptHashData to hash
    ' session keys and other streams of data.  Failure is a
    ' return value of zero.
    If CryptCreateHash(mlngProviderHandle, lngHashAlgo, _
                       ByVal 0&, ByVal 0&, mlngHashObject) = 0 Then

        InfoMsg "Failed to create hash object key." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

        gblnStopProcessing = True

    Else

        CreateHashObject = True  ' Good finish

    End If

CreateHashObject_CleanUp:
    On Error GoTo 0   ' Nullify this error trap

End Function

Private Function ProcessData(ByRef abytData() As Byte) As Boolean

    '*********************************
    ' Step #2  Hash data
    '*********************************

    ' Called by HashFile()
    '           HashString()

    Dim lngRounds As Long
    Dim strInput  As String

    Const ROUTINE_NAME As String = "ProcessData"

    On Error Resume Next

    ProcessData = True                         ' Preset to good finish
    strInput = ByteArrayToString(abytData())   ' Convert byte array to string

    ' If multiple rounds are to be performed
    For lngRounds = 1 To mlngHashRounds

        ' API CryptHashData function adds data to a specified
        ' hash object. This API can be called multiple times.
        ' If multiple rounds are performed, input data (which
        ' does not change) is hashed with the internal results
        ' of the previous iteration using current hash object.
        ' Failure is a return value of zero.
        If CryptHashData(mlngHashObject, strInput, Len(strInput), ByVal 0&) = 0 Then

            InfoMsg "Failed to hash data." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

            gblnStopProcessing = True   ' Error occurred
            ProcessData = False         ' Bad finish
            Exit For                    ' exit For..Next loop

        End If

    Next lngRounds

ProcessData_CleanUp:
    strInput = vbNullString   ' Empty variable
    On Error GoTo 0           ' Nullify this error trap

End Function

Private Function FormatOutput() As Byte()

    '*********************************
    ' Step #3  Format hashed data and
    '          return as byte array
    '*********************************

    ' Called by HashFile()
    '           HashString()

    Dim strData    As String
    Dim abytData() As Byte

    Const ROUTINE_NAME As String = "FormatOutput"

    On Error Resume Next

    strData = vbNullString             ' Verify an empty string
    strData = Space$(mlngHashLength)   ' Preload with blank spaces (01-Jul-2011 Alfred Hellmüller)

    ' API CryptGetHashParam function retrieves data
    ' that governs the operations of a hash object.
    ' The actual hash value can be retrieved by using
    ' this function. Failure is a return value of zero.
    If CryptGetHashParam(mlngHashObject, HP_HASHVAL, strData, _
                         mlngHashLength, ByVal 0&) = 0 Then

        InfoMsg "Failed to retrieve hashed data." & _
                vbNewLine & vbNewLine & _
                "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

        gblnStopProcessing = True

    Else

        strData = Left$(strData, mlngHashLength)   ' Capture amount of data required

        ' Convert raw data to hex string.
        If mblnReturnHexString Then
            strData = ConvertToHex(strData)
        End If

        ' See if there was an error in conversion
        ' or user opted to stop processing.
        If Len(strData) = 0 Then

            InfoMsg "Failed to convert byte array into string data." & _
                    vbNewLine & vbNewLine & _
                    "Source:  " & MODULE_NAME & "." & ROUTINE_NAME

            gblnStopProcessing = True   ' Time to stop processing
            ReDim abytData(0)           ' Adjust array to one null byte

        Else

            ' 05-Oct-2016  Fixed logic flaw for converting output to hex
            If mblnRetLowercase Then
                strData = LCase$(strData)   ' Convert string to lowercase
            Else
                strData = UCase$(strData)   ' Convert string to uppercase
            End If

            Erase abytData()                          ' Verify an empty array to receive data
            abytData() = StringToByteArray(strData)   ' convert string to byte array

        End If

    End If

FormatOutput_CleanUp:
    ' CryptDestroyHash() function
    ' releases hash object references
    DoEvents
    If mlngHashObject <> 0 Then
        CryptDestroyHash mlngHashObject
        mlngHashObject = 0
    End If

    strData = vbNullString      ' Empty variable
    FormatOutput = abytData()   ' Return byte array
    On Error GoTo 0             ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       ConvertToHex
'
' Description:   Convert hashed data into 2-char hex string format.
'                Please note that it is faster to load an output string
'                with spaces and then insert the data.  Constantly
'                appending data will cause the system to have to
'                reallocate space, re-evaluate new string size and
'                possibly designate a new internal address to hold this
'                data.  All of which takes extra cycles and time.
'
' Parameters:    strData - incoming string data
'
' Returns:       Hex string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-DEC-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function ConvertToHex(ByVal strData As String) As String

    ' Called by FormatOutput()

    Dim lngIndex   As Long
    Dim lngLength  As Long
    Dim lngPointer As Long
    Dim strHex     As String

    On Error GoTo ConvertToHex_Error

    lngPointer = 1             ' Initialize output pointer
    lngLength = Len(strData)   ' Capture length of incoming data
    strHex = Space$(512)       ' Faster to preload output string with spaces

    ' parse data and convert each character to hex
    For lngIndex = 1 To lngLength

        Mid$(strHex, lngPointer, 2) = Right$("0" & Hex$(Asc(Mid$(strData, lngIndex, 1))), 2)
        lngPointer = lngPointer + 2   ' increment pointer

    Next lngIndex

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        strHex = vbNullString      ' Verify return data string is empty
    Else
        strHex = TrimStr(strHex)   ' Capture correct amount of data
    End If

ConvertToHex_CleanUp:
    ConvertToHex = strHex   ' Return converted data
    On Error GoTo 0         ' Nullify this error trap
    Exit Function

ConvertToHex_Error:
    Err.Clear
    strHex = vbNullString
    gblnStopProcessing = True
    Resume ConvertToHex_CleanUp

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

    ' Called by HashFile()
    '           HashString()

    ' Determine amount of data to process
    Select Case curAmtLeft
           Case Is >= MB_5: GetBlockSize = MB_5
           Case Else:       GetBlockSize = CLng(curAmtLeft)
    End Select

End Function

' ***************************************************************************
' Procedure:     HashSelection
'
' Description:   User passes a numeric value, usually 0-6, and this routine
'                will determine the hash selection.  No data needs to be
'                passed to the HashLength parameter as this is output only.
'                If the AES provider is not available then the parameter
'                passed will be calculated to use a lesser hash algorithm.
'
' Parameters:    lngHashAlgo  - Output - Hash algorithm selection
'
' Returns:       Five digit code representing appropriate hash selection
'                along with the hashed value return string length.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine 
' ***************************************************************************
Private Sub HashSelection(ByRef lngHashAlgo As Long)

    ' Called by CreateHashObject()

    Select Case mblnAESProvider

           Case True   ' Strong (AES) provider available
                Select Case mlngHashMethod
                       Case eAPI_MD2:    lngHashAlgo = CALG_MD2:     mlngHashLength = 16   ' 0
                       Case eAPI_MD4:    lngHashAlgo = CALG_MD4:     mlngHashLength = 16   ' 1
                       Case eAPI_MD5:    lngHashAlgo = CALG_MD5:     mlngHashLength = 16   ' 2
                       Case eAPI_SHA1:   lngHashAlgo = CALG_SHA1:    mlngHashLength = 20   ' 3
                       Case eAPI_SHA256: lngHashAlgo = CALG_SHA_256: mlngHashLength = 32   ' 4
                       Case eAPI_SHA384: lngHashAlgo = CALG_SHA_384: mlngHashLength = 48   ' 5
                       Case eAPI_SHA512: lngHashAlgo = CALG_SHA_512: mlngHashLength = 64   ' 6
                End Select

           Case False  ' Default or enhanced provider available
                DoEvents
                ' Reset property value, if needed
                If mlngHashMethod > 3 Then
                    HashMethod = eAPI_SHA1  ' 3
                End If

                Select Case mlngHashMethod
                       Case eAPI_MD2:  lngHashAlgo = CALG_MD2:  mlngHashLength = 16   ' 0
                       Case eAPI_MD4:  lngHashAlgo = CALG_MD4:  mlngHashLength = 16   ' 1
                       Case eAPI_MD5:  lngHashAlgo = CALG_MD5:  mlngHashLength = 16   ' 2
                       Case eAPI_SHA1: lngHashAlgo = CALG_SHA1: mlngHashLength = 20   ' 3
                End Select
    End Select

End Sub

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
' ***************************************************************************
Private Sub GetProviderHandle()

    ' Called by AES_Ready()
    '           CreateHashObject()

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

                    End If   ' Create a provider

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

    ' Called by HashFile()
    '           HashString()
    '           GetProviderHandle()
    '           Class_Initialize()
    '           Class_Terminate()

    If mlngProviderHandle <> 0 Then
        CryptReleaseContext mlngProviderHandle, 0
    End If

    mblnAESProvider = False  ' Flag denotes Strong (AES) provider is available
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
    ' Enum enumAPI_HashAlgorithms
    Const eAPI_MD2    As Long = 0&
    Const eAPI_MD4    As Long = 1&
    Const eAPI_MD5    As Long = 2&
    Const eAPI_SHA1   As Long = 3&
    Const eAPI_SHA256 As Long = 4&
    Const eAPI_SHA384 As Long = 5&
    Const eAPI_SHA512 As Long = 6&
End Sub
#End If

Private Sub Class_Initialize()

    ' Whenever a class object is instantiated (activated), this
    ' routine is called automatically if it exist with code inside.

    ReleaseProvider            ' Release provider handle

    ' Preset property values
    HashMethod = eAPI_SHA512   ' Set to default hash algorithm
    HashRounds = 1             ' Set to default number of passes
    ReturnHexString = True     ' Return data in string format
    ReturnLowercase = False    ' Return hash string in uppercase
    StopProcessing = False     ' Continue processing

End Sub

Private Sub Class_Terminate()

    ' Whenever a class object is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.

    ReleaseProvider   ' Release provider handle

End Sub
