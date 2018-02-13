Attribute VB_Name = "modMD5"
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" _
              Alias "CryptAcquireContextA" (ByRef phProv As Long, _
              ByVal pszContainer As String, ByVal pszProvider As String, _
              ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
              
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptBinHashData Lib "advapi32.dll" Alias "CryptHashData" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long

Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const MS_ENH_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const MS_ENH_RSA = "Microsoft Enhanced RSA and AES Cryptographic Provider" 'xpsp3+ only
Private Const MS_ENH_RSA_AES_PROV_XP As String = "Microsoft Enhanced RSA and AES Cryptographic Provider (Prototype)"

Private Const CRYPT_NEWKEYSET = &H8
Private Const CRYPT_VERIFYCONTEXT  As Long = &HF0000000
  
Private Const PROV_RSA_FULL = 1
Private Const PROV_RSA_AES = 24
Private Const HP_HASHVAL = 2
Private Const ALG_CLASS_HASH = 32768 '0x8000
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_SHA_256 As Long = &H800C& 'trailing & required
Private Const CALG_SHA_512 As Long = &H800E&


Private lpHashObj As Long
Private bHashPad(200) As Byte
Private bHashPadLen As Long
Private CryptoProvider As Long
Private HashType As Long
Public error_message As String


Private Function InitProvider() As Boolean
     'make sure we can acqire a basic crypto provider or raise err
    Dim sProvider As String
    Dim sContainer As String
    Dim try As Byte
    Dim pType As Long
    
tryAgain:
    sContainer = vbNullChar
    
    Select Case try
        Case 0: sProvider = MS_ENH_RSA_AES_PROV_XP & vbNullChar
        Case 1: sProvider = MS_ENH_RSA & vbNullChar
        Case 2: sProvider = MS_ENH_PROV & vbNullChar
        Case 3: sProvider = MS_DEF_PROV & vbNullChar
        Case 4: Exit Function
    End Select
    
    pType = IIf(try = 0, PROV_RSA_AES, PROV_RSA_FULL)
    
    If Not CBool(CryptAcquireContext(CryptoProvider, ByVal sContainer, ByVal sProvider, pType, CRYPT_VERIFYCONTEXT)) Then
        sContainer = vbNullChar
        If Not CBool(CryptAcquireContext(CryptoProvider, ByVal sContainer, ByVal sProvider, pType, CRYPT_NEWKEYSET)) Then
              If try = 2 Then
                    errLog "Could not Acquire a Crypto Context on this machine"
              Else
                    try = try + 1
                    GoTo tryAgain
              End If
        End If
    End If
    
    InitProvider = True
    
End Function

Private Function InitHash() As Boolean

    Dim lReturn As Long
   
    error_message = Empty
   
    InitProvider
        
    'Attempt to acquire a handle to a Hash object
    If Not CBool(CryptCreateHash(CryptoProvider, HashType, 0, 0, lpHashObj)) Then
            errLog "InitProvider - Could not Acquire hash Context"
    End If
    
    InitHash = True
    
End Function

Function MD5(sData As String) As String
On Error GoTo hadErr

    'is NOT binary unicode safe!
    HashType = CALG_MD5
    
    InitHash
    If Not HashDigestData(sData) Then GoTo hadErr
    MD5 = GetDigestedData()
    DestroyHash
    CryptReleaseContext CryptoProvider, 0
    
    Exit Function
hadErr:
     DestroyHash
End Function

Private Function SetHashData() As Boolean
   Dim lLength As Long
   
   lLength = 200&     ' actual length of the digested data (16 or 20)
   
   If Not CBool(CryptGetHashParam(lpHashObj, HP_HASHVAL, bHashPad(0), lLength, 0)) Then
        bHashPadLen = 0
        errLog "No Hash Data"
   End If
    
   'Set the module variable to the actual length of the hash value
   bHashPadLen = lLength
   SetHashData = True
   
End Function


Private Function HashDigestData(ByVal sData As String) As Boolean
    
    bHashPadLen = 0
    
    InitHash
    
    Dim lDataLen As Long
    
    lDataLen = Len(sData)
    
    If Not CBool(CryptHashData(lpHashObj, sData, lDataLen, 0)) Then
       errLog "HashData - Unable to digest the data."
    End If
    
    'SetHashData sets the variable to holds the result
    Call SetHashData
    
    HashDigestData = True
   
End Function

Private Function GetDigestedData() As String
    Dim lError As Long
    
    Dim sData As String, sHex As String
    Dim icounter As Long
    Dim spacerChar As Byte
    
    If bHashPadLen = 0 Then errLog "GetDigest - No Data to get"
     
    For icounter = 0 To bHashPadLen - 1
        'Debug.Print bHashPad(icounter)
        sHex = Hex(bHashPad(icounter))
        If Len(sHex) = 1 Then sHex = "0" & sHex
        sData = sData & sHex
    Next
    
    GetDigestedData = sData
   
End Function

Private Sub DestroyHash()
    CryptDestroyHash lpHashObj
    bHashPadLen = 0
End Sub

Private Sub errLog(sErr As String)
    error_message = error_message & sErr & vbCrLf
    Err.Raise 1
End Sub
