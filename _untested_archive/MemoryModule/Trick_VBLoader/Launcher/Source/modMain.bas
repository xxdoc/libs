Attribute VB_Name = "modMain"
' // modMain.bas - main module of launcher
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

' // Storage item flags
Public Enum FileFlags
    FF_REPLACEONEXIST = 1                  ' // Replace file if exists
    FF_IGNOREERROR = 2                     ' // Ignore errors
End Enum

' // Execute item flags
Public Enum ExeFlags
    EF_IGNOREERROR = 1                     ' // Ignore errors
End Enum

' // Storage list item
Private Type BinStorageListItem
    ofstFileName        As Long            ' // Offset of file name
    ofstDestPath        As Long            ' // Offset of file path
    dwSizeOfFile        As Long            ' // Size of file
    ofstBeginOfData     As Long            ' // Offset of beginning data
    dwFlags             As FileFlags       ' // Flags
End Type

' // Execute list item
Private Type BinExecListItem
    ofstFileName        As Long            ' // Offset of file name
    ofstParameters      As Long            ' // Offset of parameters
    dwFlags             As ExeFlags        ' // Flags
End Type

' // Storage descriptor
Private Type BinStorageList
    dwSizeOfStructure   As Long            ' // Size of structure
    iExecutableIndex    As Long            ' // Index of main executable
    dwSizeOfItem        As Long            ' // Size of BinaryStorageItem structure
    dwNumberOfItems     As Long            ' // Number of files in storage
End Type

' // Execute list descriptor
Private Type BinExecList
    dwSizeOfStructure   As Long            ' // Size of structure
    dwSizeOfItem        As Long            ' // Size of BinaryExecuteItem structure
    dwNumberOfItems     As Long            ' // Number of items
End Type

' // Base information about project
Private Type BinProject
    dwSizeOfStructure   As Long            ' // Size of structure
    storageDescriptor   As BinStorageList  ' // Storage descriptor
    execListDescriptor  As BinExecList     ' // Command descriptor
    dwStringsTableLen   As Long            ' // Size of strings table
    dwFileTableLen      As Long            ' // Size of data table
End Type

Public pProjectData     As Long            ' // Decompressed project data
Public pStoragesTable   As Long            ' // Storage list address
Public pExecutesTable   As Long            ' // Execute list address
Public pFilesTable      As Long            ' // File table address
Public pStringsTable    As Long            ' // String table address
Public ProjectDesc      As BinProject      ' // Project descriptor

' // Startup subroutine
Public Sub Main()

    ' // Load constants
    If Not LoadConstants Then
        MessageBox 0, GetString(MID_ERRORLOADINGCONST), 0, MB_ICONERROR Or MB_SYSTEMMODAL
        GoTo EndOfProcess
    End If
    
    ' // Load project
    If Not ReadProject Then
        MessageBox 0, GetString(MID_ERRORREADINGPROJECT), 0, MB_ICONERROR Or MB_SYSTEMMODAL
        GoTo EndOfProcess
    End If
    
    ' // Copying from storage
    If Not CopyProcess Then GoTo EndOfProcess
    
    ' // Execution process
    If Not ExecuteProcess Then GoTo EndOfProcess
    
    ' // If main executable is not presented exit
    If ProjectDesc.storageDescriptor.iExecutableIndex = -1 Then GoTo EndOfProcess
    
    ' // Run exe from memory
    If Not RunProcess Then
        ' // Error occrurs
        MessageBox 0, GetString(MID_ERRORSTARTUPEXE), 0, MB_ICONERROR Or MB_SYSTEMMODAL
    End If
    
EndOfProcess:
    
    If pProjectData Then
        HeapFree GetProcessHeap(), HEAP_NO_SERIALIZE, pProjectData
    End If
    
    ExitProcess 0
    
End Sub

' // Load project
Public Function ReadProject() As Boolean
    Dim hResource       As Long:                Dim hMememory       As Long
    Dim lResSize        As Long:                Dim pRawData        As Long
    Dim status          As Long:                Dim pUncompressed   As Long
    Dim lUncompressSize As Long:                Dim lResultSize     As Long
    Dim tmpStorageItem  As BinStorageListItem:  Dim tmpExecuteItem  As BinExecListItem
    Dim pLocalBuffer    As Long
    
    ' // Load resource
    hResource = FindResource(hInstance, GetString(PROJECT), RT_RCDATA)
    If hResource = 0 Then GoTo CleanUp
    
    hMememory = LoadResource(hInstance, hResource)
    If hMememory = 0 Then GoTo CleanUp
    
    lResSize = SizeofResource(hInstance, hResource)
    If lResSize = 0 Then GoTo CleanUp
    
    pRawData = LockResource(hMememory)
    If pRawData = 0 Then GoTo CleanUp
    
    pLocalBuffer = HeapAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, lResSize)
    If pLocalBuffer = 0 Then GoTo CleanUp
    
    ' // Copy to local buffer
    CopyMemory ByVal pLocalBuffer, ByVal pRawData, lResSize
    
    ' // Set default size
    lUncompressSize = lResSize * 2
    
    ' // Do decompress...
    Do
        
        If pUncompressed Then
            pUncompressed = HeapReAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, ByVal pUncompressed, lUncompressSize)
        Else
            pUncompressed = HeapAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE, lUncompressSize)
        End If
        
        status = RtlDecompressBuffer(COMPRESSION_FORMAT_LZNT1, _
                                     ByVal pUncompressed, lUncompressSize, _
                                     ByVal pLocalBuffer, lResSize, lResultSize)
        
        lUncompressSize = lUncompressSize * 2
        
    Loop While status = STATUS_BAD_COMPRESSION_BUFFER
    
    pProjectData = pUncompressed
    
    If status Then GoTo CleanUp

    ' // Validation check
    If lResultSize < LenB(ProjectDesc) Then GoTo CleanUp
    
    ' // Copy descriptor
    CopyMemory ProjectDesc, ByVal pProjectData, LenB(ProjectDesc)
    
    ' // Check all members
    If ProjectDesc.dwSizeOfStructure <> Len(ProjectDesc) Then GoTo CleanUp
    If ProjectDesc.storageDescriptor.dwSizeOfStructure <> Len(ProjectDesc.storageDescriptor) Then GoTo CleanUp
    If ProjectDesc.storageDescriptor.dwSizeOfItem <> Len(tmpStorageItem) Then GoTo CleanUp
    If ProjectDesc.execListDescriptor.dwSizeOfStructure <> Len(ProjectDesc.execListDescriptor) Then GoTo CleanUp
    If ProjectDesc.execListDescriptor.dwSizeOfItem <> Len(tmpExecuteItem) Then GoTo CleanUp
    
    ' // Initialize pointers
    pStoragesTable = pProjectData + ProjectDesc.dwSizeOfStructure
    pExecutesTable = pStoragesTable + ProjectDesc.storageDescriptor.dwSizeOfItem * ProjectDesc.storageDescriptor.dwNumberOfItems
    pFilesTable = pExecutesTable + ProjectDesc.execListDescriptor.dwSizeOfItem * ProjectDesc.execListDescriptor.dwNumberOfItems
    pStringsTable = pFilesTable + ProjectDesc.dwFileTableLen
    
    ' // Check size
    If (pStringsTable + ProjectDesc.dwStringsTableLen - pProjectData) <> lResultSize Then GoTo CleanUp
    
    ' // Success
    ReadProject = True
    
CleanUp:
    
    If pLocalBuffer Then HeapFree GetProcessHeap(), HEAP_NO_SERIALIZE, pLocalBuffer
    
    If Not ReadProject And pProjectData Then
        HeapFree GetProcessHeap(), HEAP_NO_SERIALIZE, pProjectData
    End If
    
End Function

' // Copying process
Public Function CopyProcess() As Boolean
    Dim bItem       As BinStorageListItem:  Dim index       As Long
    Dim pPath       As Long:                Dim dwWritten   As Long
    Dim msg         As Long:                Dim lStep       As Long
    Dim isError     As Boolean:             Dim pItem       As Long
    Dim pErrMsg     As Long:                Dim pTempString As Long
    
    ' // Set pointer
    pItem = pStoragesTable
    
    ' // Go thru file list
    For index = 0 To ProjectDesc.storageDescriptor.dwNumberOfItems - 1

        ' // Copy file descriptor
        CopyMemory bItem, ByVal pItem, Len(bItem)
        
        ' // Next item
        pItem = pItem + ProjectDesc.storageDescriptor.dwSizeOfItem
        
        ' // If it is not main executable
        If index <> ProjectDesc.storageDescriptor.iExecutableIndex Then
        
            ' // Normalize path
            pPath = NormalizePath(pStringsTable + bItem.ofstDestPath, pStringsTable + bItem.ofstFileName)
            
            ' // Error occurs
            If pPath = 0 Then
            
                pErrMsg = GetString(MID_ERRORWIN32)
                MessageBox 0, pErrMsg, 0, MB_ICONERROR Or MB_SYSTEMMODAL
                GoTo CleanUp
                
            Else
                Dim hFile   As Long
                Dim disp    As CREATIONDISPOSITION
                
                ' // Set overwrite flags
                If bItem.dwFlags And FF_REPLACEONEXIST Then disp = CREATE_ALWAYS Else disp = CREATE_NEW
                
                ' // Set number of subroutine
                lStep = 0
                
                ' // Run subroutines
                Do
                    ' // Disable error flag
                    isError = False
                    
                    ' // Free string
                    If pErrMsg Then SysFreeString pErrMsg: pErrMsg = 0
                    
                    ' // Choose subroutine
                    Select Case lStep
                    Case 0  ' // 0. Create folder
                    
                        If Not CreateSubdirectories(pPath) Then isError = True
                        
                    Case 1  ' // 1. Create file
                    
                        hFile = CreateFile(pPath, FILE_GENERIC_WRITE, 0, ByVal 0&, disp, FILE_ATTRIBUTE_NORMAL, 0)
                        If hFile = INVALID_HANDLE_VALUE Then
                            If GetLastError = ERROR_FILE_EXISTS Then Exit Do
                            isError = True
                        End If
                        
                    Case 2  ' // 2. Copy data to file
                    
                        If WriteFile(hFile, ByVal pFilesTable + bItem.ofstBeginOfData, _
                                     bItem.dwSizeOfFile, dwWritten, ByVal 0&) = 0 Then isError = True
                                     
                        If dwWritten <> bItem.dwSizeOfFile Then
                            isError = True
                        Else
                            CloseHandle hFile: hFile = INVALID_HANDLE_VALUE
                        End If
                        
                    End Select
                    
                    ' // If error occurs show notification (retry, abort, ignore)
                    If isError Then
                    
                        ' // Ignore error
                        If bItem.dwFlags And FF_IGNOREERROR Then Exit Do

                        pTempString = GetString(MID_ERRORCOPYINGFILE)
                        pErrMsg = StrCat(pTempString, pPath)
                        
                        ' // Cleaning
                        SysFreeString pTempString: pTempString = 0
                        
                        Select Case MessageBox(0, pErrMsg, 0, MB_ICONERROR Or MB_SYSTEMMODAL Or MB_CANCELTRYCONTINUE)
                        Case MESSAGEBOXRETURN.IDCONTINUE: Exit Do
                        Case MESSAGEBOXRETURN.IDTRYAGAIN
                        Case Else:  GoTo CleanUp
                        End Select
                        
                    Else: lStep = lStep + 1
                    End If
                    
                Loop While lStep <= 2
                        
                If hFile <> INVALID_HANDLE_VALUE Then
                    CloseHandle hFile: hFile = INVALID_HANDLE_VALUE
                End If
                
                ' // Cleaning
                SysFreeString pPath: pPath = 0
                
            End If
            
        End If
        
    Next
    
    ' // Success
    CopyProcess = True
    
CleanUp:
    
    If pTempString Then SysFreeString pTempString
    If pErrMsg Then SysFreeString pErrMsg
    If pPath Then SysFreeString pPath
    
    If hFile <> INVALID_HANDLE_VALUE Then
        CloseHandle hFile
        hFile = INVALID_HANDLE_VALUE
    End If
    
End Function

' // Execution command process
Public Function ExecuteProcess() As Boolean
    Dim index       As Long:                Dim bItem       As BinExecListItem
    Dim pPath       As Long:                Dim pErrMsg     As Long
    Dim shInfo      As SHELLEXECUTEINFO:    Dim pTempString As Long
    Dim pItem       As Long:                Dim status      As Long

    ' // Set pointer and size
    shInfo.cbSize = Len(shInfo)
    pItem = pExecutesTable
    
    ' // Go thru all items
    For index = 0 To ProjectDesc.execListDescriptor.dwNumberOfItems - 1
    
        ' // Copy item
        CopyMemory bItem, ByVal pItem, ProjectDesc.execListDescriptor.dwSizeOfItem
        
        ' // Set pointer to next item
        pItem = pItem + ProjectDesc.execListDescriptor.dwSizeOfItem
        
        ' // Normalize path
        pPath = NormalizePath(pStringsTable + bItem.ofstFileName, 0)
        
        ' // Fill SHELLEXECUTEINFO
        shInfo.lpFile = pPath
        shInfo.lpParameters = pStringsTable + bItem.ofstParameters
        shInfo.fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI
        shInfo.nShow = SW_SHOWDEFAULT
        
        ' // Performing...
        status = ShellExecuteEx(shInfo)
        
        ' // If error occurs show notification (retry, abort, ignore)
        Do Until status
            
            If pErrMsg Then SysFreeString pErrMsg: pErrMsg = 0
            
            ' // Ignore error
            If bItem.dwFlags And EF_IGNOREERROR Then
                Exit Do
            End If
                        
            pTempString = GetString(MID_ERROREXECUTELINE)
            pErrMsg = StrCat(pTempString, pPath)
            
            SysFreeString pTempString: pTempString = 0
            
            Select Case MessageBox(0, pErrMsg, 0, MB_ICONERROR Or MB_SYSTEMMODAL Or MB_CANCELTRYCONTINUE)
            Case MESSAGEBOXRETURN.IDCONTINUE: Exit Do
            Case MESSAGEBOXRETURN.IDTRYAGAIN
            Case Else: GoTo CleanUp
            End Select

            status = ShellExecuteEx(shInfo)
            
        Loop
        
        ' // Wait for process terminaton
        WaitForSingleObject shInfo.hProcess, INFINITE
        CloseHandle shInfo.hProcess
        
    Next
    
    ' // Success
    ExecuteProcess = True
    
CleanUp:

    If pTempString Then SysFreeString pTempString
    If pErrMsg Then SysFreeString pErrMsg
    If pPath Then SysFreeString pPath
    
End Function

' // Run exe from project in memory
Public Function RunProcess() As Boolean
    Dim bItem       As BinStorageListItem:  Dim Length      As Long
    Dim pFileData   As Long
    
    ' // Get descriptor of executable file
    CopyMemory bItem, ByVal pStoragesTable + ProjectDesc.storageDescriptor.dwSizeOfItem * _
                      ProjectDesc.storageDescriptor.iExecutableIndex, Len(bItem)
    

    ' // Alloc memory within top memory addresses
    pFileData = VirtualAlloc(ByVal 0&, bItem.dwSizeOfFile, MEM_TOP_DOWN Or MEM_COMMIT, PAGE_READWRITE)
    If pFileData = 0 Then Exit Function
    
    ' // Copy raw exe file to this memory
    CopyMemory ByVal pFileData, ByVal pFilesTable + bItem.ofstBeginOfData, bItem.dwSizeOfFile
    
    ' // Free decompressed project data
    HeapFree GetProcessHeap(), HEAP_NO_SERIALIZE, pProjectData
    pProjectData = 0
    
    ' // Run exe from memory
    RunExeFromMemory pFileData, bItem.dwFlags And FF_IGNOREERROR
    
    ' ----------------------------------------------------
    ' // An error occurs
    ' // Clean memory
    
    VirtualFree ByVal pFileData, 0, MEM_RELEASE
    
    ' // If ignore error then success
    If bItem.dwFlags And FF_IGNOREERROR Then RunProcess = True
    
End Function

' // Create all subdirectories by path
Public Function CreateSubdirectories( _
                ByVal pPath As Long) As Boolean
    Dim pComponent As Long
    Dim tChar      As Integer
    
    ' // Pointer to first char
    pComponent = pPath
    
    ' // Go thru path components
    Do
    
        ' // Get next component
        pComponent = PathFindNextComponent(pComponent)
        
        ' // Check if end of line
        CopyMemory tChar, ByVal pComponent, 2
        If tChar = 0 Then Exit Do
        
        ' // Write null-terminator
        CopyMemory ByVal pComponent - 2, 0, 2
        
        ' // Check if path exists
        If PathIsDirectory(pPath) = 0 Then
        
            ' // Create folder
            If CreateDirectory(pPath, ByVal 0&) = 0 Then
                ' // Error
                CopyMemory ByVal pComponent - 2, &H5C, 2
                Exit Function
            End If
            
        End If
        
        ' // Restore path delimiter
        CopyMemory ByVal pComponent - 2, &H5C, 2
        
    Loop
    
    ' // Success
    CreateSubdirectories = True
    
End Function

' // Get normalize path (replace wildcards, append file name)
Public Function NormalizePath( _
                ByVal pPath As Long, _
                ByVal pTitle As Long) As Long
    Dim lPathLen    As Long:    Dim lRelacerLen As Long
    Dim lTitleLen   As Long:    Dim pRelacer    As Long
    Dim lTotalLen   As Long:    Dim lPtr        As Long
    Dim pTempString As Long:    Dim pRetString  As Long
    
    ' // Determine wildcard
    Select Case True
    Case IntlStrEqWorker(0, pPath, pAppRepl, 5): pRelacer = pAppPath
    Case IntlStrEqWorker(0, pPath, pSysRepl, 5): pRelacer = pSysPath
    Case IntlStrEqWorker(0, pPath, pTmpRepl, 5): pRelacer = pTmpPath
    Case IntlStrEqWorker(0, pPath, pWinRepl, 5): pRelacer = pWinPath
    Case IntlStrEqWorker(0, pPath, pDrvRepl, 5): pRelacer = pDrvPath
    Case IntlStrEqWorker(0, pPath, pDtpRepl, 5): pRelacer = pDtpPath
    Case Else: pRelacer = pStrNull
    End Select
    
    ' // Get string size
    lPathLen = lstrlen(ByVal pPath)
    lRelacerLen = lstrlen(ByVal pRelacer)
    
    ' // Skip wildcard
    If lRelacerLen Then
        pPath = pPath + 5 * 2
        lPathLen = lPathLen - 5
    End If
    
    If pTitle Then lTitleLen = lstrlen(ByVal pTitle)
    
    ' // Get length all strings
    lTotalLen = lPathLen + lRelacerLen + lTitleLen
    
    ' // Check overflow (it should be les or equal MAX_PATH)
    If lTotalLen > MAX_PATH Then Exit Function
    
    ' // Create string
    pTempString = SysAllocStringLen(0, MAX_PATH)
    If pTempString = 0 Then Exit Function
    
    ' // Copy
    lstrcpyn ByVal pTempString, ByVal pRelacer, lRelacerLen + 1
    lstrcat ByVal pTempString, ByVal pPath

    ' // If title is presented append
    If pTitle Then

        ' // Error
        If PathAddBackslash(pTempString) = 0 Then GoTo CleanUp

        ' // Copy file name
        lstrcat ByVal pTempString, ByVal pTitle
        
    End If
    
    ' // Alloc memory for translation relative path to absolute
    pRetString = SysAllocStringLen(0, MAX_PATH)
    If pRetString = 0 Then GoTo CleanUp
    
    ' // Normalize
    If PathCanonicalize(pRetString, pTempString) = 0 Then GoTo CleanUp
    
    NormalizePath = pRetString
    
CleanUp:
    
    If pTempString Then SysFreeString pTempString
    If pRetString <> 0 And NormalizePath = 0 Then SysFreeString pRetString
    
End Function

' // Concatenation strings
Public Function StrCat( _
                ByVal pStringDest As Long, _
                ByVal pStringAppended As Long) As Long
    Dim l1 As Long, l2 As Long
    
    l1 = lstrlen(ByVal pStringDest): l2 = lstrlen(ByVal pStringAppended)
    StrCat = SysAllocStringLen(0, l1 + l2)
    
    If StrCat = 0 Then Exit Function
    
    lstrcpyn ByVal StrCat, ByVal pStringDest, l1 + 1
    lstrcat ByVal StrCat, ByVal pStringAppended
    
End Function


