VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : CRegistry
' Date      : 1/4/2000
' Modified  : 8/9/2014
' Purpose   : A class module for Registry access
'---------------------------------------------------------------------------------------

'Dim objReg As New CRegistry, subkeys as Collection
'set subkeys = objReg.EnumerateSubKeys("SOFTWARE\Developer\Software Name\Folders")


Option Explicit

'Constants used for Registry Access
Private Const KEY_ALL_ACCESS = &HF003F
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = &H20019
Private Const KEY_WRITE = &H20006
Private Const KEY_WOW64_32KEY = &H200
Private Const KEY_WOW64_64KEY = &H100
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&

'API Function Declarations used for Registry
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
                                       Alias "RegSetValueExA" _
                                       (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal Reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpData As Any, _
                                        ByVal cbData As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
                                         Alias "RegQueryValueExA" _
                                         (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, lpData As Any, _
                                          lpcbData As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                                      Alias "RegOpenKeyExA" _
                                      (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" _
                                        Alias "RegDeleteValueA" _
                                        (ByVal hKey As Long, _
                                         ByVal lpValueName As String) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" _
                                      Alias "RegDeleteKeyA" _
                                      (ByVal hKey As Long, _
                                       ByVal lpSubKey As String) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
                                        Alias "RegCreateKeyExA" _
                                        (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal Reserved As Long, _
                                         ByVal lpClass As String, _
                                         ByVal dwOptions As Long, _
                                         ByVal samDesired As Long, _
                                         lpSecurityAttributes As Any, _
                                         phkResult As Long, _
                                         lpdwDisposition As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
                                     (ByVal hKey As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" _
                                      Alias "RegEnumKeyExA" _
                                      (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpName As String, _
                                       lpcbName As Long, _
                                       ByVal lpReserved As Long, _
                                       ByVal lpClass As String, _
                                       lpcbClass As Long, _
                                       lpftLastWriteTime As FILETIME) As Long

'User defined types
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Boolean
End Type
Private Type FILETIME
   intLow As Long
   intHigh As Long
End Type

'Variables & Properties
Private mlngRootKey As Long
Private mstrKeyPath As String
Private mstrValueName As String
Private mlngResult As Long
Private mlngSAM As Long
Private mstrStringValue As String
Private mstrExStringValue As String
Private mlngDWordValue As Long
Private mstrErrorMsg As String
Private marrBinValue() As Byte
Private mlngBinValueSize As Long

Property Let BinaryValue(ByVal var As Variant)
   Dim i As Integer
   ReDim marrBinValue(UBound(var))
   For i = 0 To UBound(var)
      marrBinValue(i) = var(i)
   Next
End Property

Property Let BinaryValueSize(lngSize As Long)
   mlngBinValueSize = lngSize
End Property

Property Let StringValue(ByVal strString As String)
   mstrStringValue = strString
End Property

Property Get StringValue() As String
   StringValue = mstrStringValue
End Property

Property Let ExpandedStringValue(ByVal strExStringValue As String)
   mstrExStringValue = strExStringValue
End Property

Property Get ExpandedStringValue() As String
   ExpandedStringValue = mstrExStringValue
End Property

Property Let DWordValue(ByVal lngDWord As Long)
   mlngDWordValue = lngDWord
End Property

Property Get DWordValue() As Long
   DWordValue = mlngDWordValue
End Property

Property Get Result() As Long
   Result = mlngResult
End Property

Property Let RootKey(ByVal strRootKey As String)
   strRootKey = UCase(strRootKey)
   mlngRootKey = 0
   Select Case strRootKey
      Case "HKEY_LOCAL_MACHINE", "HKLM"
         mlngRootKey = &H80000002
      Case "HKEY_CURRENT_USER", "HKCU"
         mlngRootKey = &H80000001
      Case "HKEY_USERS", "HU"
         mlngRootKey = &H80000003
      Case "HKEY_CLASSES_ROOT", "HKCR"
         mlngRootKey = &H80000000
   End Select
   If mlngRootKey = 0 Then
        MsgBox "Unknown RootKey: " & strRootKey
   End If
End Property

Property Let KeyPath(ByVal strKeyPath As String)
   mstrKeyPath = strKeyPath
End Property

Property Let ValueName(ByVal sValueName As String)
   mstrValueName = sValueName
End Property

Property Let is32Bit(ByVal x As Boolean)
   If x Then
         mlngSAM = KEY_ALL_ACCESS Or KEY_WOW64_32KEY
   Else
         mlngSAM = KEY_ALL_ACCESS Or KEY_WOW64_64KEY
   End If
End Property

Property Get ErrorMsg() As String
   ErrorMsg = mstrErrorMsg
End Property

'Method Functions
Public Function GetStringValue() As String
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim strBuffer As String
   Dim cb As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegOpenKeyEx failed"
      Exit Function
   End If
   strBuffer = Space(1024)
   cb = Len(strBuffer)
   lngResult = RegQueryValueEx(lngKeyHandle, _
                               mstrValueName, _
                               0&, _
                               REG_SZ, _
                               ByVal strBuffer, _
                               cb)
   If lngResult = 0 Then
      mstrStringValue = Left(strBuffer, cb - 1)
      mlngResult = lngResult
   Else
      mlngResult = lngResult
      mstrErrorMsg = "RegQueryValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function GetExpandedStringValue() As String
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim strBuffer As String
   Dim cb As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegOpenKey failed"
      Exit Function
   End If
   strBuffer = Space(1024)
   cb = Len(strBuffer)
   lngResult = RegQueryValueEx(lngKeyHandle, _
                               mstrValueName, _
                               0&, _
                               REG_EXPAND_SZ, _
                               ByVal strBuffer, _
                               cb)
   If lngResult = 0 Then
      mstrExStringValue = Left(strBuffer, cb - 1)
      mlngResult = lngResult
   Else
      mlngResult = lngResult
      mstrErrorMsg = "RegQueryValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function SetStringValue() As String
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim lngDisposition As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegCreateKeyEx(mlngRootKey, _
                              mstrKeyPath, _
                              0&, _
                              "", _
                              REG_OPTION_NON_VOLATILE, _
                              mlngSAM, _
                              vbNullString, _
                              lngKeyHandle, _
                              lngDisposition)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegCreateKeyEx failed"
      Exit Function
   End If
   lngResult = RegSetValueEx(lngKeyHandle, _
                             mstrValueName, _
                             0&, _
                             REG_SZ, _
                             ByVal mstrStringValue, _
                             Len(mstrStringValue))
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegSetValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function SetExpandedStringValue() As String
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim lngDisposition As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegCreateKeyEx(mlngRootKey, _
                              mstrKeyPath, _
                              0&, _
                              "", _
                              REG_OPTION_NON_VOLATILE, _
                              mlngSAM, _
                              vbNullString, _
                              lngKeyHandle, _
                              lngDisposition)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegCreateKeyEx failed"
      Exit Function
   End If
   lngResult = RegSetValueEx(lngKeyHandle, _
                             mstrValueName, _
                             0&, _
                             REG_EXPAND_SZ, _
                             ByVal mstrExStringValue, _
                             Len(mstrExStringValue))
   If lngResult > 0 Then
      mstrErrorMsg = "RegSetValueEx failed"
      mlngResult = lngResult
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function GetDWordValue() As Long
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim lngBuffer As Long
   Dim cb As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegOpenKeyEx failed"
      Exit Function
   End If
   cb = Len(lngBuffer)
   lngResult = RegQueryValueEx(lngKeyHandle, _
                               mstrValueName, _
                               0&, _
                               REG_DWORD, _
                               lngBuffer, _
                               cb)
   If lngResult = 0 Then
      mlngDWordValue = lngBuffer
      mlngResult = lngResult
   Else
      mlngResult = lngResult
      mstrErrorMsg = "RegQueryValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function SetDWordValue() As Long
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim lngDisposition As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegCreateKeyEx(mlngRootKey, _
                              mstrKeyPath, _
                              0&, _
                              "", _
                              REG_OPTION_NON_VOLATILE, _
                              mlngSAM, _
                              vbNullString, _
                              lngKeyHandle, _
                              lngDisposition)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegCreateKeyEx failed"
      Exit Function
   End If
   lngResult = RegSetValueEx(lngKeyHandle, _
                             mstrValueName, _
                             0&, _
                             REG_DWORD, _
                             mlngDWordValue, _
                             4)
   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegSetValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)
End Function

Public Function CheckKey() As Long
   Dim lngResult As Long
   Dim lngKeyHandle As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   mlngResult = lngResult
   If lngResult = 0 Then
      lngResult = RegCloseKey(lngKeyHandle)
   End If
End Function

Public Function DeleteKey() As Long
   Dim lngResult As Long
   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegDeleteKey(mlngRootKey, mstrKeyPath)
   mlngResult = lngResult
End Function

Public Function DeleteValue() As Long
   Dim lngKeyHandle As Long
   Dim lngResult As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   If lngResult = 0 Then
      lngResult = RegDeleteValue(lngKeyHandle, _
                                 mstrValueName)
      mlngResult = lngResult
      lngResult = RegCloseKey(lngKeyHandle)
   End If
End Function

Public Function EnumerateSubKeys(Optional path As String) As Collection
   Dim lngKeyHandle As Long
   Dim lngResult As Long
   Dim lngNumSubKeys As Long
   Dim lngIndex As Long
   Dim lngTime As FILETIME
   Dim strValue As String
   Dim strClass As String
   Dim lngValueLen As Long
   Dim lngClassLen As Long
   Dim marrSubKeys As New Collection
    
   If Len(path) > 0 Then KeyPath = path
   
   mlngResult = 0
   mstrErrorMsg = vbNullString
   lngIndex = 0
   lngResult = RegOpenKeyEx(mlngRootKey, _
                            mstrKeyPath, _
                            0&, _
                            mlngSAM, _
                            lngKeyHandle)
   If lngResult <> ERROR_SUCCESS Then
      mlngResult = lngResult
      Exit Function
   End If
   Do
      lngValueLen = 2000
      lngClassLen = 2000
      strValue = String(lngValueLen, 0)
      strClass = String(lngClassLen, 0)
      lngResult = RegEnumKeyEx(lngKeyHandle, _
                               lngIndex, _
                               ByVal strValue, _
                               lngValueLen, _
                               0&, _
                               ByVal strClass, _
                               lngClassLen, _
                               lngTime)

      If lngResult = ERROR_SUCCESS Then
         strValue = Left(strValue, lngValueLen)
         marrSubKeys.Add strValue
         lngIndex = lngIndex + 1
      End If
      
   Loop While lngResult = ERROR_SUCCESS
   
   lngResult = RegCloseKey(lngKeyHandle)
   
   Set EnumerateSubKeys = marrSubKeys

End Function

Public Function SetBinaryValue() As Long
   Dim lngResult As Long
   Dim lngKeyHandle As Long
   Dim lngDisposition As Long
   Dim lngArrSize As Long

   mlngResult = 0
   mstrErrorMsg = vbNullString

   lngResult = RegCreateKeyEx(mlngRootKey, _
                              mstrKeyPath, _
                              0&, _
                              "", _
                              REG_OPTION_NON_VOLATILE, _
                              mlngSAM, _
                              vbNullString, _
                              lngKeyHandle, _
                              lngDisposition)

   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegCreateKeyEx failed"
      Exit Function
   End If

   lngResult = RegSetValueEx(lngKeyHandle, _
                             mstrValueName, _
                             0&, _
                             REG_BINARY, _
                             marrBinValue(0), _
                             mlngBinValueSize)

   If lngResult > 0 Then
      mlngResult = lngResult
      mstrErrorMsg = "RegSetValueEx failed"
   End If
   lngResult = RegCloseKey(lngKeyHandle)

End Function


Private Sub Class_Initialize()
     mlngSAM = KEY_ALL_ACCESS Or KEY_WOW64_32KEY
     RootKey = "HKLM"
End Sub
