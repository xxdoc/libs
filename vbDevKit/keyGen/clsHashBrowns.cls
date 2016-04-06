VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHashBrowns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwFlags As Long) As Long
    
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0" & vbNullChar
Private Const PROV_RSA_FULL = 1
Private Const HP_HASHVAL = &H2
Private Const CRYPT_NEWKEYSET = &H8
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_SHA = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)

Private cryptContext As Long

Private Sub Class_Initialize()
    AcquireContext
End Sub

Private Sub Class_Terminate()
    ReleaseContext
End Sub

Private Sub AcquireContext()
        If Not CBool(CryptAcquireContext(cryptContext, vbNullChar, MS_DEF_PROV, PROV_RSA_FULL, 0)) Then
            If Not CBool(CryptAcquireContext(cryptContext, vbNullChar, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
                Exit Sub
            End If
        End If
End Sub

Private Sub ReleaseContext()
    If (cryptContext <> 0) Then Call CryptReleaseContext(cryptContext, 0)
End Sub

Private Function Hash(data As String, useSHA As Boolean) As String
    Dim ht As Long, hLen As Long, h As String, hl As Long
     
    If (cryptContext = 0) Then MsgBox "Error: Crypto Context Was not Acquired", vbCritical: Exit Function
    
    hLen = IIf(useSHA, 20, 16)
    ht = IIf(useSHA, CALG_SHA, CALG_MD5)
    h = String(hLen, vbNull)
    
    If Not CBool(CryptCreateHash(cryptContext, ht, 0, 0, hl)) Then GoTo failed
    If Not CBool(CryptHashData(hl, data, Len(data), 0)) Then GoTo failed
    If Not CBool(CryptGetHashParam(hl, HP_HASHVAL, h, hLen, 0)) Then GoTo failed
    
    Hash = h
Exit Function
failed: Hash = "Error: Crypto Context Was not Acquired"
End Function

Private Function ToHex(sIn As String, Optional pure As Boolean = True) As String
    '{73BABCE8-129A-470F-935B-BA35A28C8E4A}
    '{      8   4     4    4         12   }
    Dim c As String, i As Integer, j As Integer
    If Not pure And Len(sIn) > 16 Then sIn = Mid(sIn, 1, 16)
    For i = 1 To Len(sIn)
        c = Hex(Asc(Mid(sIn, i, 1)))
        j = Len(ToHex) + 2
        If Not pure Then
            If j = 8 Or j = 13 Or j = 18 Or j = 23 Then c = c & "-"
        End If
        ToHex = ToHex & c
    Next
    If Not pure Then ToHex = "{" & ToHex & "}"
End Function

Function MD5(strIn As String, Optional asHex As Boolean = True) As String
    MD5 = Hash(strIn, False)
    If asHex Then MD5 = ToHex(MD5)
End Function

Function SHA1(strIn As String, Optional asHex As Boolean = True) As String
    SHA1 = Hash(strIn, True)
    If asHex Then SHA1 = ToHex(SHA1)
End Function

