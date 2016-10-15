Attribute VB_Name = "mWinVerify"
Private Declare Function WinVerifyTrust Lib "wintrust.dll" (ByVal hwnd As Long, ByRef pgActionID As GUID, ByRef pWVTData As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (Destination As Any, ByRef length As Long)
Private Declare Sub RtlFillMemory Lib "kernel32.dll" (Destination As Long, length As Long, Fill As Byte)

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type WINTRUST_FILE_INFO
    cbStruct As Long
    pcwszFilePath As String
    hFile As Long
    pgKnownSubject As GUID
End Type

Private Type WINTRUST_DATA
    cbStruct As Long
    pPolicyCallbackData As Long
    pSIPClientData As Long
    dwUIChoice As Long
    fdwRevocationChecks As Long
    dwUnionChoice As Long
    pFile As Long
    pCatalog As Long
    pBlob As Long
    pSgnr As Long
    pCert As Long
    dwStateAction As Long
    hWVTStateData As Long
    pwszURLReference As String
    dwProvFlags As Long
    dwUIContext As Long
End Type


Private Const WTD_UI_NONE = 2
Private Const WTD_REVOKE_NONE = 0
Private Const WTD_SAFER_FLAG = &H100

Private Const TRUST_E_ACTION_UNKNOWN = -2146762750
Private Const ERROR_SUCCESS = 0&
Private Const TRUST_E_NOSIGNATURE = &H800B0100
Private Const TRUST_E_EXPLICIT_DISTRUST = &H800B0111
Private Const CRYPT_E_SECURITY_SETTINGS = &H80092026
Private Const WTD_UI_ALL = 1& 'Display all UI.
Private Const WTD_UI_NOBAD = 3& ' Do not display any negative UI.
Private Const WTD_UI_NOGOOD = 4& ' Do not display any positive UI.
Private Const WTD_REVOKE_WHOLECHAIN = 1& ' Revocation checking will be done on the whole chain.
Private Const WTD_CHOICE_FILE = 1& ' Use the file pointed to by pFile.
Private Const WTD_CHOICE_CATALOG = 2& ' Use the catalog pointed to by pCatalog.
Private Const WTD_CHOICE_BLOB = 3& ' Use the BLOB pointed to by PBlob.
Private Const WTD_CHOICE_SIGNER = 4& ' Use the WINTRUST_SGNR_INFO structure pointed to by pSgnr.
Private Const WTD_CHOICE_CERT = 5& ' Use the certificate pointed to by pCert.
Private Const INVALID_HANDLE_VALUE = (-1)
Private Const TRUST_E_SUBJECT_FORM_UNKNOWN = (&H800B0003)
Private Const TRUST_E_PROVIDER_UNKNOWN = (&H800B0001)
Private Const TRUST_E_SUBJECT_NOT_TRUSTED = (&H800B0004)
Private Const SignatureOrFileCorrupt = &H80096010
Private Const SignatureExpired = &H800B0101

Public Enum SigResults
    srNotSigned = 0
    srSignedOK = 1
    srSignedFail = 2
    srError = 3
    srCorrupt = 4
    srSigExpired = 5
End Enum

Private lastError As Long

Public Function isSigned(x As SigResults) As Boolean
    isSigned = True
    If x = srNotSigned Then isSigned = False
End Function

Public Function SigToColor(x As SigResults) As ColorConstants
    If x = srSigExpired Then SigToColor = vbRed
    If x = srError Then SigToColor = vbRed
    If x = srSignedFail Then SigToColor = vbRed
    If x = srCorrupt Then SigToColor = vbRed
    If x = srNotSigned Then SigToColor = vbBlack
    If x = srSignedOK Then SigToColor = vbBlue
End Function

Public Function SigToStr(x As SigResults)
    If x = srError Then SigToStr = "Error " & Hex(lastError)
    If x = srNotSigned Then SigToStr = "Not Signed"
    If x = srSignedOK Then SigToStr = "Valid"
    If x = srSignedFail Then SigToStr = "Invalid"
    If x = srCorrupt Then SigToStr = "Corrupt"
    If x = srSigExpired Then SigToStr = "Expired"
End Function

Public Function VerifyFileSignature(sFile$) As SigResults

    Dim uVerifyV2 As GUID, uWTfileinfo As WINTRUST_FILE_INFO
    Dim uWTdata As WINTRUST_DATA, lRet&
    Dim ret As SigResults
    
    ret = srError
    
    RtlFillMemory ByVal VarPtr(uWTfileinfo), ByVal LenB(uWTfileinfo), ByVal 0
    RtlFillMemory ByVal VarPtr(uWTdata), ByVal LenB(uWTdata), ByVal 0
     
    With uWTfileinfo
        .cbStruct = Len(uWTfileinfo)
        .pcwszFilePath = sFile
    End With
    
    With uWTdata
        .cbStruct = Len(uWTdata)
        .dwUIChoice = WTD_UI_NONE
        .fdwRevocationChecks = WTD_REVOKE_NONE
        .dwUnionChoice = WTD_CHOICE_FILE
        .dwProvFlags = WTD_SAFER_FLAG
        .pFile = VarPtr(uWTfileinfo)
    End With
    
    If CLSIDFromString(StrPtr("{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}"), uVerifyV2) = 0 Then
        lStatus = WinVerifyTrust(0, uVerifyV2, uWTdata)
    
         Select Case (lStatus)
            Case ERROR_SUCCESS
                'MsgBox "The file """ & sFile & """ is signed and the signature was verified."
                ret = srSignedOK
                
            Case TRUST_E_NOSIGNATURE
            
                dwLastError = GetLastError()
                
                If (TRUST_E_NOSIGNATURE = dwLastError) Or (TRUST_E_SUBJECT_FORM_UNKNOWN = dwLastError) Or (TRUST_E_PROVIDER_UNKNOWN = dwLastError) Then '// The file was not signed.
                    'MsgBox "The file """ & sFile & """ is not signed."
                    ret = srNotSigned
                Else
                    'The signature was not valid or there was an error opening the file.
                    'MsgBox "An unknown error occurred trying to verify the signature of the """ & sFile & """ file."
                    If dwLastError = 0 Then
                        ret = srNotSigned
                    Else
                        lastError = dwLastError
                        ret = srError
                    End If
                End If
                
            Case TRUST_E_PROVIDER_UNKNOWN
                ret = srNotSigned
                
            Case TRUST_E_SUBJECT_FORM_UNKNOWN ' 800b0003
                ret = srNotSigned
                
            Case SignatureExpired
                ret = srSigExpired
                
            Case TRUST_E_EXPLICIT_DISTRUST ' // The hash that represents the subject or the publisher ' // is not allowed by the admin or user.
                'MsgBox "The signature is present, but specifically disallowed."
                ret = srSignedFail
                
            Case TRUST_E_SUBJECT_NOT_TRUSTED '// The user clicked "No" when asked to install and run.
                'MsgBox "The signature is present, but not trusted."
                ret = srSignedFail
                
            Case CRYPT_E_SECURITY_SETTINGS
                ret = srSignedFail
                'MsgBox "CRYPT_E_SECURITY_SETTINGS - The hash " & _
                "representing the subject or the publisher wasn't " & _
                "explicITly trusted by the admin and admin policy " & _
                "has disabled user trust. No signature, publisher " & _
                "or timestamp errors."
                
            Case SignatureOrFileCorrupt
                ret = srCorrupt
                
            Case Else ' // The UI was disabled in dwUIChoice or the admin policy ' // has disabled user trust. lStatus contains the ' // publisher or time stamp chain error.
                'MsgBox "Error is: 0x" & Hex(lStatus) & "."
                lastError = lStatus
                ret = srError
                
        End Select
    
    End If
    
    VerifyFileSignature = ret
    
 

End Function



