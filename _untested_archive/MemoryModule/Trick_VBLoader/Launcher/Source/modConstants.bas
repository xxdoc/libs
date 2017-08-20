Attribute VB_Name = "modConstants"
' // modConstants.bas - main module for loading constants
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

Public Enum MessagesID
    MID_ERRORLOADINGCONST = 100     ' // Errors
    MID_ERRORREADINGPROJECT = 101   '
    MID_ERRORCOPYINGFILE = 102      '
    MID_ERRORWIN32 = 103            '
    MID_ERROREXECUTELINE = 104      '
    MID_ERRORSTARTUPEXE = 105       '
    PROJECT = 200                   ' // Project resource ID
    API_LIB_KERNEL32 = 300          ' // Library names
    API_LIB_NTDLL = 350             '
    API_LIB_USER32 = 400            '
    MSG_LOADER_ERROR = 500
End Enum

' // Paths

Public pAppPath  As Long            ' // Path to application
Public pSysPath  As Long            ' // Path to System32
Public pTmpPath  As Long            ' // Path to Temp
Public pWinPath  As Long            ' // Path to Windows
Public pDrvPath  As Long            ' // Path to system drive
Public pDtpPath  As Long            ' // Path to desktop

' // Substitution constants

Public pAppRepl  As Long
Public pSysRepl  As Long
Public pTmpRepl  As Long
Public pWinRepl  As Long
Public pDrvRepl  As Long
Public pDtpRepl  As Long
Public pStrNull  As Long            ' // \0

Public hInstance    As Long         ' // Base address
Public lpCmdLine    As Long         ' // Command line
Public SI           As STARTUPINFO  ' // Startup parameters
Public LCID         As Long         ' // LCID

' // Load constants
Public Function LoadConstants() As Boolean
    Dim lSize   As Long
    Dim pBuf    As Long
    Dim index   As Long
    Dim ctl     As tagINITCOMMONCONTROLSEX
    
    ' // Load windows classes
    ctl.dwSize = Len(ctl)
    ctl.dwICC = &H3FFF&
    InitCommonControlsEx ctl
    
    ' // Get startup parameters
    GetStartupInfo SI
    
    ' // Get command line
    lpCmdLine = GetCommandLine()
    
    ' // Get base address
    hInstance = GetModuleHandle(ByVal 0&)
    
    ' // Get LCID
    LCID = GetUserDefaultLCID()
    
    ' // Alloc memory for strings
    pBuf = SysAllocStringLen(0, MAX_PATH)
    If pBuf = 0 Then Exit Function
    
    ' // Get path to process file name
    If GetModuleFileName(hInstance, pBuf, MAX_PATH) = 0 Then GoTo CleanUp
    
    ' // Leave only directory
    PathRemoveFileSpec pBuf
    
    ' // Save path
    pAppPath = SysAllocString(pBuf)
    
    ' // Get Windows folder
    If GetWindowsDirectory(pBuf, MAX_PATH) = 0 Then GoTo CleanUp
    pWinPath = SysAllocString(pBuf)
    
    ' // Get System32 folder
    If GetSystemDirectory(pBuf, MAX_PATH) = 0 Then GoTo CleanUp
    pSysPath = SysAllocString(pBuf)
    
    ' // Get Temp directory
    If GetTempPath(MAX_PATH, pBuf) = 0 Then GoTo CleanUp
    pTmpPath = SysAllocString(pBuf)
    
    ' // Get system drive
    PathStripToRoot pBuf
    pDrvPath = SysAllocString(pBuf)
    
    ' // Get desktop path
    If SHGetFolderPath(0, CSIDL_DESKTOPDIRECTORY, 0, SHGFP_TYPE_CURRENT, pBuf) Then GoTo CleanUp
    pDtpPath = SysAllocString(pBuf)
    
    ' // Load wildcards
    For index = 1 To 6
        If LoadString(hInstance, index, pBuf, MAX_PATH) = 0 Then GoTo CleanUp
        Select Case index
        Case 1: pAppRepl = SysAllocString(pBuf)
        Case 2: pSysRepl = SysAllocString(pBuf)
        Case 3: pTmpRepl = SysAllocString(pBuf)
        Case 4: pWinRepl = SysAllocString(pBuf)
        Case 5: pDrvRepl = SysAllocString(pBuf)
        Case 6: pDtpRepl = SysAllocString(pBuf)
        End Select
    Next
    
    ' // vbNullChar
    pStrNull = SysAllocStringLen(0, 0)

    ' // Success
    LoadConstants = True
    
CleanUp:
    
    If pBuf Then SysFreeString pBuf
    
End Function

' // Obtain string from resource (it should be less or equal MAX_PATH)
Public Function GetString( _
                ByVal ID As MessagesID) As Long
                
    GetString = SysAllocStringLen(0, MAX_PATH)
    
    If GetString Then
    
        If LoadString(hInstance, ID, GetString, MAX_PATH) = 0 Then SysFreeString GetString: GetString = 0: Exit Function
        If SysReAllocString(GetString, GetString) = 0 Then SysFreeString GetString: GetString = 0: Exit Function
        
    End If
    
End Function


