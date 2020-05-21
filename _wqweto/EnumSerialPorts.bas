    Option Explicit
     
    Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As Long, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
     
    Public Function EnumSerialPorts() As Variant
        Dim sBuffer         As String
        Dim lIdx            As Long
        Dim vRetVal         As Variant
        Dim lCount          As Long
        
        ReDim vRetVal(0 To 255) As Variant
        sBuffer = String$(100000, 1)
        Call QueryDosDevice(0, sBuffer, Len(sBuffer))
        sBuffer = vbNullChar & sBuffer
        For lIdx = 1 To 255
            If InStr(1, sBuffer, vbNullChar & "COM" & lIdx & vbNullChar, vbTextCompare) > 0 Then
                vRetVal(lCount) = "COM" & lIdx
                lCount = lCount + 1
            End If
        Next
        If lCount = 0 Then
            vRetVal = Array()
        Else
            ReDim Preserve vRetVal(0 To lCount - 1) As Variant
        End If
        EnumSerialPorts = vRetVal
    End Function
     
    Private Sub Form_Load()
        Dim vElem           As Variant
        
        For Each vElem In EnumSerialPorts
            Debug.Print vElem
        Next
    End Sub
	
	
	
For new USB serial ports arrival you can use RegisterDeviceNotification to get WM_DEVICECHANGE message like this:
This uses the Modern Subclassing Thunk for the IDE-safe subclassing.
http://www.vbforums.com/showthread.php?872819-VB6-The-Modern-Subclassing-Thunk-(MST)

	Option Explicit
 
'--- Windows Messages
Private Const WM_DEVICECHANGE               As Long = &H219
'--- for RegisterDeviceNotification
Private Const DEVICE_NOTIFY_WINDOW_HANDLE   As Long = &H0
Private Const DBT_DEVTYP_DEVICEINTERFACE    As Long = &H5
Private Const DBT_DEVICEARRIVAL             As Long = &H8000&
Private Const DBT_DEVICEREMOVECOMPLETE      As Long = &H8004&
Private Const GUID_DEVINTERFACE_USB_DEVICE  As String = "{A5DCBF10-6530-11D2-901F-00C04FB951ED}"
 
Private Declare Function RegisterDeviceNotification Lib "user32" Alias "RegisterDeviceNotificationA" (ByVal hRecipient As Long, ByRef NotificationFilter As Any, ByVal Flags As Long) As Long
Private Declare Function UnregisterDeviceNotification Lib "user32" (ByVal Handle As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, pclsid As Any) As Long
 
Private Type DEV_BROADCAST_DEVICEINTERFACE
    dbcc_size           As Long
    dbcc_devicetype     As Long
    dbcc_reserved       As Long
    dbcc_classguid(0 To 3) As Long
    dbcc_name           As Long
End Type
 
Private m_hDevNotify                As Long
Private m_pSubclass                 As IUnknown
 
Private Property Get pvAddressOfSubclassProc() As Form1
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Property
 
Private Sub Form_Load()
    Dim uFilter         As DEV_BROADCAST_DEVICEINTERFACE
 
    '--- on device insert/eject notify w/ WM_DEVICECHANGE
    uFilter.dbcc_size = Len(uFilter)
    uFilter.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE
    Call CLSIDFromString(StrPtr(GUID_DEVINTERFACE_USB_DEVICE), uFilter.dbcc_classguid(0))
    m_hDevNotify = RegisterDeviceNotification(hWnd, uFilter, DEVICE_NOTIFY_WINDOW_HANDLE)
    Set m_pSubclass = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    If m_hDevNotify <> 0 Then
        Call UnregisterDeviceNotification(m_hDevNotify)
        m_hDevNotify = 0
    End If
End Sub
 
Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    Select Case wMsg
    Case WM_DEVICECHANGE
        Select Case wParam
        Case DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE
            Debug.Print "wParam=&H" & Hex(wParam), Timer
        End Select
        Exit Function
        Handled = True
    End Select
end function