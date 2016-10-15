Attribute VB_Name = "mCryptQuery"
'References:
'   http://support.microsoft.com/default.aspx?scid=kb;en-us;323809
'   http://www.vbforums.com/showthread.php?189218-Please-Please-help-with-this-API-CertFindCertificateInStore
'   http://www.vbforums.com/showthread.php?630485-RESOLVED-Having-trouble-with-CryptMsgGetParam-%28vb6%29&p=3903783#post3903783
'   http://microsoft.public.vb.winapi.narkive.com/r3xfPBdP/cryptqueryobject-or-winverifytrust
'   as always thanks to LaVolpe for the excellent help

Option Explicit

Private Declare Function CryptQueryObject Lib "Crypt32.dll" (ByVal dwObjectType As Long, _
                ByVal pvObject As Long, _
                ByVal dwExpectedContentTypeFlags As Long, _
                ByVal dwExpectedFormatTypeFlags As Long, _
                ByVal dwFlags As Long, _
                ByRef pdwMsgAndCertEncodingType As Long, _
                ByRef pdwContentType As Long, _
                ByRef pdwFormatType As Long, _
                ByRef phCertStore As Long, _
                ByRef phMsg As Long, _
                ByRef ppvContext As Long _
) As Long
   
Private Declare Function CertGetNameStringA Lib "Crypt32.dll" ( _
            ByVal pCertContext As Long, _
            ByVal dwType As Long, _
            ByVal dwFlags As Long, _
            ByVal pvTypePara As Long, _
            ByVal pszNameString As String, _
            ByVal cchNameString As Long _
) As Long

Private Declare Function CertFindCertificateInStore2 Lib "Crypt32.dll" _
    Alias "CertFindCertificateInStore" ( _
                                ByVal hCertStore As Long, _
                                ByVal dwCertEncodingType As Long, _
                                ByVal dwFindFlags As Long, _
                                ByVal dwFindType As Long, _
                                      pvFindPara As Any, _
                                ByVal pPrevCertContext As Long _
) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function CryptMsgGetParam Lib "Crypt32.dll" (ByRef hCryptMsg As Long, ByVal dwParamType As Long, ByVal dwIndex As Long, pvData As Any, ByRef pcbData As Long) As Long
Private Declare Function CertFreeCertificateContext Lib "Crypt32.dll" (ByVal pCertContext As Long) As Long
Private Declare Function CertCloseStore Lib "Crypt32.dll" (ByVal hCertStore As Long, ByVal flags As Long) As Long
Private Declare Function CryptMsgClose Lib "Crypt32.dll" (ByVal hCryptMsg As Long) As Long

Private Const CMSG_SIGNER_INFO_PARAM As Long = 6
Private Const CMSG_CERT_PARAM = 12
Private Const CMSG_CMS_SIGNER_INFO_PARAM = 39
Private Const CMSG_ENCODED_SIGNER = 28
Private Const X509_ASN_ENCODING = 1
Private Const PKCS_7_ASN_ENCODING = &H10000
Private Const CERT_QUERY_OBJECT_FILE As Long = &H1
Private Const CERT_QUERY_CONTENT_PKCS7_SIGNED_EMBED As Long = 10
Private Const CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED As Long = 2 ^ CERT_QUERY_CONTENT_PKCS7_SIGNED_EMBED
Private Const CERT_QUERY_FORMAT_BINARY As Long = &H1
Private Const CERT_QUERY_FORMAT_FLAG_BINARY As Long = 2 ^ CERT_QUERY_FORMAT_BINARY
Private Const CERT_FIND_SUBJECT_CERT = &HB0000
Private Const CERT_NAME_SIMPLE_DISPLAY_TYPE = 4
Private Const CERT_NAME_ISSUER_FLAG = 1

Private Type CRYPT_INTEGER_BLOB
    cbData As Long
    pbData As Long
End Type

Private Type CERT_NAME_BLOB
    cbData As Long
    pbData As Long
End Type

Private Type CRYPT_OBJID_BLOB
    cbData As Long
    pbData As Long
End Type

Private Type CRYPT_BIT_BLOB
    cbData As Long
    pbData As Long
    cUnusedBits As Long
End Type

Private Type CERT_EXTENSION
    pszObjId As String
    fCritical As Boolean
    value As CRYPT_OBJID_BLOB
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId As Long 'String <--crashity crash in copymemory rtcconvertunicode automatic handling pfft
    Parameters As CRYPT_OBJID_BLOB
End Type

Private Type CERT_PUBLIC_KEY_INFO
    Algorithm As CRYPT_ALGORITHM_IDENTIFIER
    PublicKey As CRYPT_BIT_BLOB
End Type

Private Type CERT_INFO
    dwVersion As Long
    SerialNumber As CRYPT_INTEGER_BLOB
    SignatureAlgorithm As CRYPT_ALGORITHM_IDENTIFIER
    issuer As CERT_NAME_BLOB
    NotBefore As Date
    NotAfter As Date
    subject As CERT_NAME_BLOB
    SubjectPublicKeyInfo As CERT_PUBLIC_KEY_INFO
    IssuerUniqueId As CRYPT_BIT_BLOB
    SubjectUniqueId As CRYPT_BIT_BLOB
    cExtension As Long
    rgExtension As CERT_EXTENSION
End Type

Private Type CRYPT_ATTRIBUTES
    pszObjId As Long
    rgValue As CRYPT_INTEGER_BLOB
End Type

Private Type CRYPT_DATA_BLOB
    cbData As Long
    pbData As Long
End Type

Private Type CMSG_SIGNER_INFO
    dwVersion As Long
    issuer As CERT_NAME_BLOB
    SerialNumber As CRYPT_INTEGER_BLOB
    HashAlgorithm As CRYPT_ALGORITHM_IDENTIFIER
    HashEncryptionAlgorithm As CRYPT_ALGORITHM_IDENTIFIER
    EncryptedHash As CRYPT_DATA_BLOB
    AuthAttrs As CRYPT_ATTRIBUTES
    UnauthAttrs As CRYPT_ATTRIBUTES
End Type

Public ErrMsg As String

Public Function GetSigner(fPath As String, ByRef out_Issuer As String, ByRef out_Subject As String) As Boolean

    Dim fResult As Long
    Dim dwEncoding As Long
    Dim dwContentType As Long
    Dim dwFormatType As Long
    Dim hStore As Long
    Dim hMsg As Long
    Dim bufSz As Long
    Dim ci As CERT_INFO
    Dim signerInfo As CMSG_SIGNER_INFO
    Dim pCertContext As Long
    Dim dwData  As Long
    Dim sz As Long
    Dim szName As String
    Dim b() As Byte
     
    Const my_encoding = X509_ASN_ENCODING Or PKCS_7_ASN_ENCODING
     
    out_Issuer = Empty
    out_Subject = Empty
    ErrMsg = Empty
    
    fResult = CryptQueryObject( _
                CERT_QUERY_OBJECT_FILE, _
                ByVal StrPtr(fPath), _
                CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED, _
                CERT_QUERY_FORMAT_FLAG_BINARY, 0&, dwEncoding, _
                dwContentType, dwFormatType, hStore, hMsg, ByVal 0& _
            )
    
    If fResult = 1 Then
        ' Determine buffer size for the next call to CryptMsgGetParam
        fResult = CryptMsgGetParam(ByVal hMsg, CMSG_SIGNER_INFO_PARAM, 0&, ByVal 0&, bufSz)
        If fResult = 0 Then GoTo cleanup
    Else
        ErrMsg = fPath & " is not signed!"
        GoTo cleanup
    End If
    
    'should now contain the buffer size required for the next call to CryptMsgGetParam
    If bufSz = 0 Then
        ErrMsg = "Unable to determine buffer size for CryptMsgGetParam!"
        GoTo cleanup
    End If
        
    ReDim b(bufSz)
    
    fResult = CryptMsgGetParam(ByVal hMsg, CMSG_SIGNER_INFO_PARAM, 0&, b(0), bufSz)
    
    If fResult = 0 Then GoTo cleanup
    
    CopyMemory signerInfo, b(0), LenB(signerInfo) 'bug: can crash here with call stack unicode convert sometimes?
                                                  '     problem starts with __vbaRecAnsiToUni - fixed see def above 9.3.15
    
    ci.issuer = signerInfo.issuer
    ci.SerialNumber = signerInfo.SerialNumber
    
    pCertContext = CertFindCertificateInStore2(hStore, my_encoding, 0, CERT_FIND_SUBJECT_CERT, ci, 0)
    
    If pCertContext = 0 Then
        If Err.LastDllError = &H80092004 Then
            ErrMsg = "Failed to find client signing certificate."
        Else
            ErrMsg = "CertFindCertificateInStore failed with " & Hex(Err.LastDllError)
        End If
        GoTo cleanup
    End If
        
    dwData = 0
    dwData = CertGetNameStringA(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE, CERT_NAME_ISSUER_FLAG, 0, szName, dwData)
    
    If dwData = 0 Then
        ErrMsg = "CertGetNameString(ISSUER_FLAG) failed"
    Else
        szName = String(dwData + 1, " ")
        
        dwData = CertGetNameStringA(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE, CERT_NAME_ISSUER_FLAG, 0, szName, dwData)
        
        If dwData = 0 Then
            ErrMsg = "CertGetNameString(ISSUER_FLAG) failed (2)"
        Else
            out_Issuer = Trim(Mid(szName, 1, dwData - 1))
            'MsgBox "Issuer: " & out_Issuer
        End If
    End If
    
    dwData = 0
    dwData = CertGetNameStringA(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE, 0, 0, szName, dwData)
    
    If dwData = 0 Then
        ErrMsg = "CertGetNameString(SubjectName) failed"
    Else
        szName = String(dwData + 1, " ")
        
        dwData = CertGetNameStringA(pCertContext, CERT_NAME_SIMPLE_DISPLAY_TYPE, 0, 0, szName, dwData)
        
        If dwData = 0 Then
            ErrMsg = "CertGetNameString(SubjectName) failed (2)"
        Else
            out_Subject = Trim(Mid(szName, 1, dwData - 1))
            'MsgBox "Issuer: " & out_Issuer
        End If
    End If
    
    If Len(out_Issuer) > 0 Or Len(out_Subject) > 0 Then GetSigner = True
    
cleanup:
 
    If pCertContext <> 0 Then CertFreeCertificateContext pCertContext
    If hStore <> 0 Then CertCloseStore hStore, 0
    If hMsg <> 0 Then CryptMsgClose hMsg


End Function

