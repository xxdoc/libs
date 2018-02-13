VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'9/7/07  fixed error in crypt acquire context, for some reason old method broke
'        with a latter service pack :(
'
'5.17.12 fixed overflow error in HashFile on very large files (changed var to long)
'1.21.14 bugfix - HashFile now only opens files in Read Mode (now can hash readonly files)

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

'CALG_SHA_256    0x0000800c 'these 3 require xp sp3+ and:
'CALG_SHA_384    0x0000800d '   Microsoft Enhanced RSA and AES Cryptographic Provider
'CALG_SHA_512    0x0000800e

Public Enum HashTypes
    hMD2 = CALG_MD2
    hMD4 = CALG_MD4
    hmd5 = CALG_MD5
    hSHA = CALG_SHA
End Enum

Enum HashFormats
    HexFormat
    DelimitedHexFormat
    NumericFormat
    AsciiFormat        'not unicode binary safe
End Enum

 
Private lpHashObj As Long
Private bHashPad(200) As Byte
Private bHashPadLen As Long
Private CryptoProvider As Long
Private HashType As HashTypes

Public error_message As String

Event HadError(sMsg As String)

Private Sub Class_Terminate()
    CryptReleaseContext CryptoProvider, 0
End Sub

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

Function HashString(sData As String, Optional method As HashTypes = hmd5, Optional returnFormat As HashFormats = HexFormat) As String
On Error GoTo hadErr

    'is NOT binary unicode safe!
    If method = 256 Then method = CALG_SHA_256
    If method = 512 Then method = CALG_SHA_512
    
    HashType = method
    
    InitHash
    If Not HashDigestData(sData) Then GoTo hadErr
    
    HashString = GetDigestedData(returnFormat)
    DestroyHash
    
    Exit Function
hadErr:
     DestroyHash
End Function


Function HashBytes(bytes() As Byte, Optional method As HashTypes = hmd5, Optional returnFormat As HashFormats = HexFormat) As String
On Error GoTo hell
    
    'binary unicode safe!
    If method = 256 Then method = CALG_SHA_256
    If method = 512 Then method = CALG_SHA_512

    HashType = method
    
    InitHash
    CryptBinHashData lpHashObj, bytes(LBound(bytes)), AryElems(bytes), 0
    SetHashData
    
    HashBytes = GetDigestedData(returnFormat)
    DestroyHash
    
    Exit Function
hell:
      DestroyHash
End Function

Private Function AryElems(x As Variant) As Long
    'returns number of elements in array regardless of where bounds start
    On Error GoTo hell
    AryElems = UBound(x) - LBound(x) + 1
    Exit Function
hell: AryElems = -1
End Function

Function HashFile(sFilePath As String, Optional method As HashTypes = hmd5, Optional returnFormat As HashFormats = HexFormat) As String
    On Error GoTo hadErr

    If Not FileExists(sFilePath) Then errLog "HashFile - File not found"
        
    If method = 256 Then method = CALG_SHA_256
    If method = 512 Then method = CALG_SHA_512

    HashType = method
    
    InitHash
    If Not HashDigestFile(sFilePath) Then GoTo hadErr
    
    HashFile = GetDigestedData(returnFormat)
    DestroyHash
    
    Exit Function
hadErr:
    DestroyHash
End Function




'Initializes bHashPad byte array with the hash value
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

Private Function GetDigestedData(ByVal hasHashFormats As HashFormats) As String
    Dim lError As Long
    
    Dim sData As String, sHex As String
    Dim icounter As Long
    Dim spacerChar As Byte
    
    If bHashPadLen = 0 Then errLog "GetDigest - No Data to get"
    
    If hasHashFormats = DelimitedHexFormat Then spacerChar = Asc(" ")
    
    Select Case hasHashFormats
        Case HexFormat, DelimitedHexFormat
            For icounter = 0 To bHashPadLen - 1
                'Debug.Print bHashPad(icounter)
                sHex = Hex(bHashPad(icounter))
                If Len(sHex) = 1 Then sHex = "0" & sHex
                sData = sData & sHex & IIf(spacerChar <> 0, Chr(spacerChar), Empty)
                sHex = ""
            Next
        Case NumericFormat
            For icounter = 0 To bHashPadLen - 1
                sData = sData & CStr(bHashPad(icounter))
            Next
        Case AsciiFormat
            For icounter = 0 To bHashPadLen - 1
                sData = sData & Chr(bHashPad(icounter))
            Next
    End Select
    
    GetDigestedData = sData
   
End Function

Private Function HashDigestFile(ByVal sSourceFile As String) As Boolean
    On Error GoTo oops

    bHashPadLen = 0
    
    Dim fNum As Long, fLen As Long, fBlocks As Long, fExtraBytes As Long, i As Long
    Dim fDat() As Byte
    Const BlockSize = 1000
    
    fNum = FreeFile
    Open sSourceFile For Binary Access Read As fNum
    
    fLen = LOF(fNum)
    fBlocks = fLen \ BlockSize
    fExtraBytes = fLen Mod BlockSize
    
    ReDim fDat(1 To BlockSize)
    
    For i = 1 To fBlocks
        Get fNum, , fDat()
        CryptBinHashData lpHashObj, fDat(1), BlockSize, 0
    Next
    
    If fExtraBytes <> 0 Then
        ReDim fDat(1 To fExtraBytes)
        Get fNum, , fDat()
        CryptBinHashData lpHashObj, fDat(1), fExtraBytes, 0
    End If
     
    Close fNum
    Erase fDat()
    SetHashData
    
    HashDigestFile = True
    
    Exit Function
oops:
     Close fNum
     errLog "Runtime error HashFile - " & Err.Description
End Function


Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Sub DestroyHash()
    CryptDestroyHash lpHashObj
    bHashPadLen = 0
End Sub

Private Sub errLog(sErr As String)
    error_message = error_message & sErr & vbCrLf
    RaiseEvent HadError(sErr)
    Err.Raise 1
End Sub
