Attribute VB_Name = "Module1"
Option Explicit
Public lpData As Long
Public fileSize As Long
Public hInst As Long

Public Const MMIO_INSTALLPROC = &H10000    'mmioInstallIOProc:install
'MMIOProc

Public Const MMIO_GLOBALPROC = &H10000000    'mmioInstallIOProc: install
'globally

Public Const MMIO_READ = &H0
Public Const MMIOM_CLOSE = 4
Public Const MMIOM_OPEN = 3
Public Const MMIOM_READ = MMIO_READ
Public Const MMIO_REMOVEPROC = &H20000
Public Const MMIOM_SEEK = 2
Public Const SEEK_CUR = 1
Public Const SEEK_END = 2
Public Const SEEK_SET = 0
Public Const MEY = &H2059454D    'This is the value of "MEY " run
'through FOURCC

'Create a user defined variable for the API function calls
Public Type MMIOINFO
    dwFlags As Long
    fccIOProc As Long
    pIOProc As Long
    wErrorRet As Long
    htask As Long
    cchBuffer As Long
    pchBuffer As String
    pchNext As String
    pchEndRead As String
    pchEndWrite As String
    lBufOffset As Long
    lDiskOffset As Long
    adwInfo(4) As Long
    dwReserved1 As Long
    dwReserved2 As Long
    hmmio As Long
End Type

'Finds the specified resource in an executable file. The function
'returns a resource handle that can be used by other functions used
'to load the resource.
Public Declare Function FindResource Lib "kernel32" Alias "FindResourceA" ( _
        ByVal hInstance As Long, _
        ByVal lpName As String, _
        ByVal lpType As String) As Long

'Returns a global memory handle to a resource in the specified
'module. The resource is only loaded after calling the LockResource
'function to get a pointer to the resource data.
Public Declare Function LoadResource Lib "kernel32" ( _
        ByVal hInstance As Long, _
        ByVal hResInfo As Long) As Long

'Locks the specified resource. The function returns a 32-bit pointer
'to the data for the resource.
Public Declare Function LockResource Lib "kernel32" ( _
        ByVal hResData As Long) As Long

'Loads the specified dll file and maps the address space for the
'current process.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpLibFileName As String) As Long

'Frees the specified dll file loaded with the LoadLibrary function.
Public Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long

'Copies a block of memory from one location to another.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        hpvDest As Any, _
        hpvSource As Any, _
        ByVal cbCopy As Long)

'Installs or removes a custom I/O procedure. This function also
'locates an installed I/O procedure, using its corresponding
'four-character code.
Public Declare Function mmioInstallIOProc Lib "winmm" Alias "mmioInstallIOProcA" ( _
        ByVal fccIOProc As Long, _
        ByVal pIOProc As Long, _
        ByVal dwFlags As Long) As Long

'Sends the specified command string to an MCI device.
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
        ByVal lpstrCommand As String, _
        ByVal lpstrReturnString As Long, _
        ByVal uReturnLength As Long, _
        ByVal hwndCallback As Long) As Long

'Get MCI error description from return code MCISendString
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" ( _
        ByVal ErrorNumber As Long, _
        ByVal ReturnBuffer As String, _
        ByVal ReturnBufferSize As Long) As Long    'BOOL

'Returns the size, in bytes, of the specified resource
Public Declare Function SizeofResource Lib "kernel32" ( _
        ByVal hInstance As Long, _
        ByVal hResInfo As Long) As Long

'Accesses a unique storage system, such as a database or file
'archive. Install or remove this callback function with the
'mmioInstallIOProc function.
Public Function IOProc(ByRef lpMMIOInfo As MMIOINFO, ByVal uMessage As Long, _
    ByVal lParam1 As Long, _
    ByVal lParam2 As Long) As Long

    Static alreadyOpened As Boolean

    Select Case uMessage
        Case MMIOM_OPEN
            If Not alreadyOpened Then
                alreadyOpened = True
                lpMMIOInfo.lDiskOffset = 0
            End If
            IOProc = 0

        Case MMIOM_CLOSE
            IOProc = 0

        Case MMIOM_READ:
            Call CopyMemory(ByVal lParam1, ByVal _
                    lpData + lpMMIOInfo.lDiskOffset, lParam2)
            lpMMIOInfo.lDiskOffset = lpMMIOInfo.lDiskOffset + lParam2
            IOProc = lParam2

        Case MMIOM_SEEK

            Select Case lParam2
                Case SEEK_SET
                    lpMMIOInfo.lDiskOffset = lParam1

                Case SEEK_CUR
                    lpMMIOInfo.lDiskOffset = lpMMIOInfo.lDiskOffset + lParam1
                    lpMMIOInfo.lDiskOffset = fileSize - 1 - lParam1

                Case SEEK_END
                    lpMMIOInfo.lDiskOffset = fileSize - 1 - lParam1
            End Select

            IOProc = lpMMIOInfo.lDiskOffset

        Case Else
            IOProc = -1    ' Unexpected msgs.  For instance, we do not
            ' process MMIOM_WRITE in this sample
    End Select

End Function    ' IOProc

' Get the error that just occured (if one did occur)
Public Function ShowMCIError(Optional ByVal errNum As Long)
    Dim strMsg As String

    errNum = LOWORD(errNum)    '//We only need low word, low word is actual error number
    ' Get the error description based on the error number

    strMsg = String(260, Chr(0))
    If mciGetErrorString(errNum, strMsg, 260) <> 0 Then
        MsgBox Left(strMsg, InStr(strMsg, Chr(0)) - 1), vbCritical, "Error: " & errNum
    Else
        MsgBox "Unknown Error", vbCritical, "Error: " & errNum
    End If
End Function

' Function that extracts the "Low Order" (lower 16bits) of a 32bit number
Private Function LOWORD(ByVal dwValue As Long) As Integer
    LOWORD = Val("&H" & Right("0000" & Hex(dwValue), 4))
End Function

