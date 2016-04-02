Attribute VB_Name = "mdlEnumControls"
'*********************************************************************************************
'
' Enumerating Installed OCX controls
'
' Main Module
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 07/21/1999
' Last Updated: 07/21/1999
'
'*********************************************************************************************

Option Explicit

Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function SHLoadInProc Lib "shell32" (rclsid As IID) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpOLEStr As String, pclsid As IID) As Long

Const HKEY_CLASSES_ROOT = &H80000000

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Any, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

' Reg Key Security Options
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = READ_CONTROL Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = &H1F0000 Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK

Const REG_SZ = 1

'-----------------------------------------------------------------------------------------

Type ControlInfo
    Description As String
    File As String
    PROGID As String
    CLSID As String
    TYPELIB As String
End Type

Const IMPLEMENTEDCATEGORIES_CONTROL = "\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}"
Const INPROC_SERVER = "\InprocServer"
Const INPROC_SERVER32 = "\InprocServer32"
Const PROGID = "\ProgID"
Const TYPELIB = "\TypeLib"
Const CONTROL = "\Control"

Private Function ExistsFile(ByVal FileName As String) As Boolean

    If Dir$(FileName) <> "" Then
        ExistsFile = True
    End If
    
End Function

Private Function GetDefaultValue(ByVal hKey As Long, ByVal SubKey As String) As String
Dim Data As String, DataL As Long, hCLSIDKey As Long

    If RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, hCLSIDKey) = 0 Then
        
        Data = String$(512, 0)
        DataL = Len(Data)
    
        If RegQueryValueEx(hCLSIDKey, vbNullString, 0, REG_SZ, ByVal Data, DataL) = 0 Then
            GetDefaultValue = Left$(Data, DataL - 1)
        End If
        
        RegCloseKey hCLSIDKey
        
    End If
    
End Function



Private Function ExistsTypeLib(ByVal TYPELIB As String) As Boolean
Dim Data As String, DataL As Long, hKey As Long

    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib\" & TYPELIB, 0, KEY_READ, hKey) = 0 Then
        
        ExistsTypeLib = True
         
        RegCloseKey hKey
        
    End If
    
End Function




Public Function EnumControls(Ctrls() As ControlInfo) As Long
Dim CLSID As Long, MaxKeyLen As Long, Index As Long, R As Long
Dim KeyName As String, KeyNameL As Long, hControlKey As Long
Dim CtrlInfo As ControlInfo, Max As Long, IID As IID

    ReDim Ctrls(0 To 0)
    
    ' Open HKCR\CLSID
    
    If RegOpenKeyEx(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ, CLSID) = 0 Then
    
        ' Get max subkeys lenght
        RegQueryInfoKey CLSID, vbNullString, ByVal 0&, 0&, ByVal 0&, MaxKeyLen, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&
            
        ' Initialize buffer
        KeyName = String$(MaxKeyLen + 1, 0)
        KeyNameL = Len(KeyName)
        
        ' Enum HKCR\CLSID subkeys
        Do While RegEnumKeyEx(CLSID, Index, KeyName, KeyNameL, 0, 0&, ByVal 0&, ByVal 0&) = 0
        
            ' Try to open the subkey HKCR\CLSID\KeyName\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}
            ' {40FC6ED4-2438-11CF-A3DB-080036F12502} is control category
            R = RegOpenKeyEx(CLSID, Left$(KeyName, KeyNameL) & IMPLEMENTEDCATEGORIES_CONTROL, 0, KEY_READ, hControlKey)
            
            ' If the subkey does not exist try to open
            ' the subkey HKCR\CLSID\KeyName\Control
            If R <> 0 Then
                R = RegOpenKeyEx(CLSID, Left$(KeyName, KeyNameL) & CONTROL, 0, KEY_READ, hControlKey)
            End If
            
            If R = 0 Then
                
                ' Get the Info
                With CtrlInfo

                    .CLSID = Left$(KeyName, KeyNameL)
                    .Description = GetDefaultValue(CLSID, .CLSID)
                    .File = GetDefaultValue(CLSID, .CLSID & INPROC_SERVER32)

                    ' If there's no INPROC_SERVER32
                    ' try with INPROC_SERVER, maybe
                    ' the control is a 16bits one.
                    If .File = "" Then
                        .File = GetDefaultValue(CLSID, .CLSID & INPROC_SERVER)
                    End If
                    
                    .PROGID = GetDefaultValue(CLSID, .CLSID & PROGID)
                    .TYPELIB = GetDefaultValue(CLSID, .CLSID & TYPELIB)
                    
                    ' Check if the file is not empty and exist
                    ' and the typelib guid is valid.
                    If .File <> "" And ExistsFile(.File) And ExistsTypeLib(.TYPELIB) Then
                                        
                        ReDim Preserve Ctrls(0 To Max)
                        
                        Ctrls(Max) = CtrlInfo
                        Max = Max + 1
                    
                    Else
                        
                        Debug.Print "Invalid class: "; .Description
                        
                    End If
                    
                End With
                
                ' Close the subkey
                RegCloseKey hControlKey
            
            End If
            
            KeyName = String$(MaxKeyLen + 1, 0)
            KeyNameL = Len(KeyName)

            Index = Index + 1
            
        Loop
        
        ' Close HKCR\CLSID
        RegCloseKey CLSID
    
    End If
    
    EnumControls = Max

End Function


