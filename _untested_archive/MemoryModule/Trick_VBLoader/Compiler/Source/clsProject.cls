VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // clsProject.cls - project class of TrickVB6Installer application
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit

' // Error codes and messages
Private Const ER_INVALID_PE_FORMAT          As Long = vbObjectError + 1
Private Const ER_INVALID_PE_FORMAT_STRING   As String = "Invalid executable file"

' // +----------------------------------------------------------------------------------------------+
' // |             Project that placed in resources consists of set of structures.                  |
' // |    There is BinaryProject structure at the beginning that contains base information about    |
' // | project. List items are placed behind this structure: list of storage items precedes list of |
' // |      execution command. Binary data of files loaded to resource follow all list items.       |
' // |          List of string (file names, parameters, etc.) is placed after binary data.          |
' // |                      All object structures are declared below.                               |
' // +----------------------------------------------------------------------------------------------+

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

' // Resource data
Private Type ResourceData
    dwSize              As Long            ' // Size of raw data
    LangID              As Long            ' // Language ID
    binData()           As Byte            ' // Raw data
End Type

' // Resource entry
Private Type ResourceEntry
    strName             As String          ' // Name of resource
    lngData()           As ResourceData    ' // List of data
End Type

' // Resource type
Private Type ResourceType
    strType             As String          ' // Type id
    entData()           As ResourceEntry   ' // List of entries
End Type

' // Resources list
Private Type ResourcesList
    dwNumberOfTypes     As Long
    ResourceTypes()     As ResourceType
End Type

Private mStorage        As clsStorage      ' // Storage list object
Private mExecute        As clsExecute      ' // Execute list object
Private mModify         As Boolean         ' // Determines whether project is modified or not
Private mFileName       As String          ' // Project file name
Private mManifest       As String          ' // Manifest text
Private mUseManifest    As Boolean         ' // Manifest usage flag

' // Manifest usage
Public Property Get UseManifest() As Boolean
    UseManifest = mUseManifest
End Property
Public Property Let UseManifest( _
                    ByRef Value As Boolean)
    mUseManifest = Value
    mModify = True
End Property

' // Text of manifest
Public Property Get Manifest() As String
    Manifest = mManifest
End Property
Public Property Let Manifest( _
                    ByRef Value As String)
    mManifest = Value
    mModify = True
End Property

' // Project file name
Public Property Get FileName() As String
    FileName = mFileName
End Property
Public Property Let FileName( _
                    ByRef Value As String)
    mFileName = Value
End Property

' // Storage list
Public Property Get Storage() As clsStorage
    Set Storage = mStorage
End Property

' // Execute list
Public Property Get Execute() As clsExecute
    Set Execute = mExecute
End Property

' // Has project been modified
Public Property Get Modify() As Boolean
    Modify = mModify
End Property
Public Property Let Modify(ByVal Value As Boolean)
    mModify = Value
End Property

' // Project base directory
Public Property Get BaseDirectory() As String
    BaseDirectory = IIf(Len(mFileName), GetFilePath(mFileName), App.Path)
End Property

' // Compile project
Public Function Compile( _
                ByRef sFileName As String) As Boolean
    Dim stringsTable    As clsStream           ' // String table
    Dim filesTable      As clsStream           ' // File table
    Dim itemsTable      As clsStream           ' // Elements table
    Dim entireFile      As clsStream           ' // All project
    Dim loaderStream    As clsStream           ' // Stream with raw loader data
    Dim binStorageItem  As BinStorageListItem
    Dim binExecuteItem  As BinExecListItem
    Dim fileExtension   As String              ' // Output extension
    Dim loaderPath      As String              ' // Path to loader app
    Dim resList         As ResourcesList       ' // List of all resources from main executable
    Dim hUpdate         As Long                ' // Resource update handle
    Dim rawProject()    As Byte                ' // Raw project data (to save to bin)
        
    On Error GoTo error_handler
    
    If mStorage.Count = 0 And mExecute.Count = 0 Then Exit Function
    
    ' // Create binary streams
    Set stringsTable = New clsStream
    Set filesTable = New clsStream
    Set itemsTable = New clsStream
    Set entireFile = New clsStream
    
    ' // Save all storage items
    AppendStorageItemsInformation filesTable, stringsTable, itemsTable
    
    ' // Save all execute items
    AppendExecuteItemsInformation stringsTable, itemsTable
    
    Dim binProjectHeader    As BinProject

    ' // Set project information
    With binProjectHeader
    
        .dwSizeOfStructure = Len(binProjectHeader)
        .dwStringsTableLen = stringsTable.Size
        .dwFileTableLen = filesTable.Size
        
        .storageDescriptor.dwSizeOfStructure = Len(.storageDescriptor)
        .storageDescriptor.dwSizeOfItem = Len(binStorageItem)
        .storageDescriptor.dwNumberOfItems = mStorage.Count
        .storageDescriptor.iExecutableIndex = mStorage.MainExecutable
        
        .execListDescriptor.dwSizeOfStructure = Len(.execListDescriptor)
        .execListDescriptor.dwSizeOfItem = Len(binExecuteItem)
        .execListDescriptor.dwNumberOfItems = mExecute.Count
        
    End With
    
    ' // Save all data
    entireFile.WriteData VarPtr(binProjectHeader), Len(binProjectHeader)
    entireFile.WriteStream itemsTable
    entireFile.WriteStream filesTable
    entireFile.WriteStream stringsTable
    
    ' // Compress stream
    entireFile.CompressStream
    
    ' // Get file extension
    fileExtension = GetFileExtension(sFileName)

    Select Case 0
    Case StrComp(fileExtension, ".exe", vbTextCompare)    ' // Save to EXE

        ' // Extract resources from main exe
        ExtractResources resList
        
        ' // Get loader path
        loaderPath = GetLoaderPath()
        
        If Len(loaderPath) = 0 Then
            MsgBox "Loader not found", vbExclamation
            GoTo ERROR_COMPILING
        End If
        
        Set loaderStream = New clsStream
        
        LoadFileToStream loaderPath, loaderStream

        If loaderStream.Size <= 0 Then
            MsgBox "Invalid loader file", vbExclamation
            GoTo ERROR_COMPILING
        End If
        
        SaveStreamToFile sFileName, loaderStream
        
        ' // Start of updating resources
        hUpdate = BeginUpdateResource(StrPtr(sFileName), False)
        
        ' // Store resources from Main executable to loader
        If resList.dwNumberOfTypes > 0 Then
            Dim resType As Long:    Dim resID   As Long
            Dim types   As Long:    Dim names   As Long
            Dim langs   As Long

            For types = 0 To resList.dwNumberOfTypes - 1
                For names = 0 To UBound(resList.ResourceTypes(types).entData)
                    For langs = 0 To UBound(resList.ResourceTypes(types).entData(names).lngData)
                        
                        If Left$(resList.ResourceTypes(types).strType, 1) = "#" Then
                            resType = CLng(Mid$(resList.ResourceTypes(types).strType, 2))
                        Else
                            resType = -1
                        End If
                        
                        Select Case resList.ResourceTypes(types).strType
                        Case "#" & CStr(RT_VERSION), _
                             "#" & CStr(RT_GROUP_ICON), _
                             "#" & CStr(RT_ICON)
                            
                            If Left$(resList.ResourceTypes(types).entData(names).strName, 1) = "#" Then
                                resID = CLng(Mid$(resList.ResourceTypes(types).entData(names).strName, 2))
                            Else
                                resID = StrPtr(resList.ResourceTypes(types).entData(names).strName)
                            End If
                            
                            ' // Store numerical resource
                            If UpdateResource(hUpdate, resType, resID, _
                                              resList.ResourceTypes(types).entData(names).lngData(langs).LangID, _
                                              resList.ResourceTypes(types).entData(names).lngData(langs).binData(0), _
                                              resList.ResourceTypes(types).entData(names).lngData(langs).dwSize) = 0 Then
                                MsgBox "Unable to update resources", vbExclamation
                                GoTo ERROR_COMPILING
                            End If

                        End Select

                    Next
                Next
            Next
        End If
        
        ' // Extract project
        ReDim rawProject(entireFile.Size - 1)
        entireFile.BufferPointer = 0
        entireFile.ReadData VarPtr(rawProject(0)), entireFile.Size
        
        ' // Store project data
        If UpdateResource(hUpdate, RT_RCDATA, StrPtr("PROJECT"), 0, rawProject(0), entireFile.Size) = 0 Then
            MsgBox "Unable to save project to resources", vbExclamation
            GoTo ERROR_COMPILING
        End If
        
        ' // Store manifest
        If mUseManifest And Len(mManifest) > 0 Then
            If UpdateResource(hUpdate, 24, 1, 0, ByVal StrPtr(mManifest), LenB(mManifest)) = 0 Then
                MsgBox "Unable to save manifest to resources", vbExclamation
                GoTo ERROR_COMPILING
            End If
        End If

    Case Else  ' // Save to BIN (raw project data)
        
        ' // Save to file
        SaveStreamToFile sFileName, entireFile

    End Select
    
    Compile = True
    
ERROR_COMPILING:

    If hUpdate Then
        EndUpdateResource hUpdate, False
    End If
    
    Exit Function
    
error_handler:
    
    MsgBox "An error occured" & vbNewLine & Err.Description
    
End Function

' // Save project
Public Function Save( _
                ByRef sFileName As String) As Boolean
    Dim pBag        As PropertyBag:     Dim buf()       As Byte
    Dim Ret         As Long:            Dim stor        As clsStorageItem
    Dim fullPath    As String:          Dim relPath     As String
    Dim oldMod      As Boolean:         Dim oStream     As clsStream
    
    On Error GoTo ERROR_SAVING
    
    oldMod = mModify
    
    ' // Translate all path to new base directory
    If sFileName <> mFileName Then
    
        For Each stor In mStorage
        
            fullPath = ToAbsolute(stor.FileName, Me.BaseDirectory)
            relPath = ToRelative(fullPath, GetFilePath(sFileName))
            If Len(relPath) = 0 Then relPath = fullPath
            stor.FileName = relPath
            
        Next
        
    End If
    
    ' // Create property bag and stream
    Set pBag = New PropertyBag
    Set oStream = New clsStream
    
    ' // Save all members
    pBag.WriteProperty "Storage", mStorage.Save.Contents
    pBag.WriteProperty "Execute", mExecute.Save.Contents
    pBag.WriteProperty "Manifest", mManifest
    pBag.WriteProperty "UseManifest", mUseManifest
    
    ' // Create project file
    buf = pBag.Contents
    oStream.WriteData VarPtr(buf(0)), UBound(buf) + 1
    
    SaveStreamToFile sFileName, oStream

    mFileName = sFileName
    mModify = False
    Save = True
    
    Exit Function
    
ERROR_SAVING:

    ' // Translate all pathes to old base directory
    If sFileName <> mFileName Then
    
        For Each stor In mStorage
            fullPath = ToAbsolute(stor.FileName, GetFilePath(sFileName))
            stor.FileName = ToRelative(fullPath, Me.BaseDirectory)
        Next
        
    End If
    
    mModify = oldMod
    
End Function

' // Load project
Public Function Load( _
                ByRef sFileName As String) As Boolean
    Dim Size    As Long:    Dim buf()   As Byte
    Dim Ret     As Long:    Dim oStream As clsStream
    
    On Error GoTo ERROR_LOADING
    
    Set oStream = New clsStream
    
    ' // Try to open file
    LoadFileToStream sFileName, oStream
    
    If oStream.Size <= 0 Then
        MsgBox "Unable to open file '" & sFileName & "'", vbExclamation
        Exit Function
    End If
    
    ReDim buf(oStream.Size - 1)
    oStream.BufferPointer = 0
    oStream.ReadData VarPtr(buf(0)), oStream.Size
    
    Dim pBag        As PropertyBag: Dim lProp       As PropertyBag
    Dim lbuf()      As Byte:        Dim mStor       As clsStorage
    Dim mExe        As clsExecute:  Dim sManifest   As String
    
    ' // Create property bag
    Set pBag = New PropertyBag
    ' // Set data
    pBag.Contents = buf()
    
    ' // Load members
    lbuf = pBag.ReadProperty("Storage")
    Set lProp = New PropertyBag
    lProp.Contents = lbuf
    Set mStor = New clsStorage
    Set mStor.Project = Me
    mStor.Load lProp
    Set mExe = New clsExecute
    Set mExe.Project = Me
    lbuf = pBag.ReadProperty("Execute")
    lProp.Contents = lbuf
    mExe.Load lProp
    
    ' // Load manifest
    sManifest = pBag.ReadProperty("Manifest")
    mUseManifest = pBag.ReadProperty("UseManifest")
    
    mManifest = sManifest
    
    Set mStorage = mStor
    Set mExecute = mExe
    mFileName = sFileName
    
    mModify = False
    Load = True
    
ERROR_LOADING:

End Function

' // Save all execute items
Private Sub AppendExecuteItemsInformation( _
            ByVal stringsStream As clsStream, _
            ByVal itemsStream As clsStream)
    Dim eItem   As clsExecuteItem
    Dim binItem As BinExecListItem
    Dim pString As String
    
    ' // Go thru execute items
    For Each eItem In mExecute
        
        ' // Set properties
        binItem.dwFlags = eItem.Flags
        
        ' // Save file name to strings table
        binItem.ofstFileName = stringsStream.BufferPointer
        pString = eItem.FileName & vbNullChar
        stringsStream.WriteData StrPtr(pString), LenB(pString)
        binItem.ofstParameters = stringsStream.BufferPointer
        pString = eItem.Parameters & vbNullChar
        stringsStream.WriteData StrPtr(pString), LenB(pString)

        ' // Save item
        itemsStream.WriteData VarPtr(binItem), Len(binItem)
        
    Next
    
End Sub

' // Save all storage items
Private Sub AppendStorageItemsInformation( _
            ByVal filesStream As clsStream, _
            ByVal stringsStream As clsStream, _
            ByVal itemsStream As clsStream)
    Dim sItem       As clsStorageItem
    Dim fullPath    As String
    Dim binItem     As BinStorageListItem
    Dim fileTitle   As String
    
    ' // Go thru storage items
    For Each sItem In mStorage

        ' // Get absolute path to file
        fullPath = ToAbsolute(sItem.FileName, Me.BaseDirectory)
        
        ' // Save pointer to data
        binItem.ofstBeginOfData = filesStream.BufferPointer
        
        ' // Load file
        binItem.dwSizeOfFile = LoadFileToStream(fullPath, filesStream)
        
        fileTitle = GetFileTitle(sItem.FileName, True)
        
        ' // Save string info
        binItem.ofstFileName = stringsStream.BufferPointer
        stringsStream.WriteData StrPtr(fileTitle), LenB(fileTitle) + 2
        binItem.ofstDestPath = stringsStream.BufferPointer
        stringsStream.WriteData StrPtr(sItem.DestinationPath), LenB(sItem.DestinationPath) + 2
        
        ' // Save other info
        binItem.dwFlags = sItem.Flags
        
        ' // Save list item to stream
        itemsStream.WriteData VarPtr(binItem), Len(binItem)
        
    Next
    
End Sub

' // Load resources (icons and version)
Private Sub ExtractResources( _
            ByRef resList As ResourcesList)
    Dim dosHdr      As IMAGE_DOS_HEADER:                Dim ntHdr       As IMAGE_NT_HEADERS
    Dim sec()       As IMAGE_SECTION_HEADER:            Dim mainDir     As IMAGE_RESOURCE_DIRECTORY
    Dim entry       As IMAGE_RESOURCE_DIRECTORY_ENTRY:  Dim index       As Long
    Dim resIdx      As Long:                            Dim resSize     As Long
    Dim strSize     As Long:                            Dim oStream     As clsStream
    Dim fullPath    As String
    
    If mStorage.MainExecutable = -1 Then
        resList.dwNumberOfTypes = 0
        Exit Sub
    End If
    
    ' // Get absolute path to file
    fullPath = ToAbsolute(mStorage(mStorage.MainExecutable).FileName, Me.BaseDirectory)
        
    Set oStream = New clsStream
    
    LoadFileToStream fullPath, oStream
    
    ' // Get Dos header
    oStream.BufferPointer = 0
    oStream.ReadData VarPtr(dosHdr), Len(dosHdr)

    ' // Check
    If dosHdr.e_magic <> IMAGE_DOS_SIGNATURE Then
        MsgBox "Main executable is not proper executable file."
        Err.Raise ER_INVALID_PE_FORMAT, ER_INVALID_PE_FORMAT_STRING
    End If
    
    oStream.BufferPointer = dosHdr.e_lfanew
    
    ' // Get nt headers
    oStream.ReadData VarPtr(ntHdr), Len(ntHdr)

    ' // Check
    If ntHdr.Signature <> IMAGE_NT_SIGNATURE Or _
       ntHdr.OptionalHeader.Magic <> IMAGE_NT_OPTIONAL_HDR32_MAGIC Then
        MsgBox "Main executable is not proper executable file."
        Err.Raise ER_INVALID_PE_FORMAT, ER_INVALID_PE_FORMAT_STRING
    End If

    ' // Check if resources exist
    resIdx = ntHdr.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_RESOURCE).VirtualAddress
    resSize = ntHdr.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_RESOURCE).Size
    
    If resIdx = 0 Then
        resList.dwNumberOfTypes = 0
        Exit Sub
    End If

    ' // Get sections
    ReDim sec(ntHdr.FileHeader.NumberOfSections - 1)
    
    oStream.BufferPointer = dosHdr.e_lfanew + _
                            Len(ntHdr.Signature) + _
                            Len(ntHdr.FileHeader) + _
                            ntHdr.FileHeader.SizeOfOptionalHeader
    oStream.ReadData VarPtr(sec(0)), Len(sec(0)) * ntHdr.FileHeader.NumberOfSections

    ' // Get index of resource table in array
    resIdx = RVA2RAW(resIdx, sec())

    Dim types   As Long:    Dim idx1 As Long
    Dim names   As Long:    Dim idx2 As Long
    Dim langs   As Long:    Dim idx3 As Long
    Dim namDir  As IMAGE_RESOURCE_DIRECTORY
    Dim lngDir  As IMAGE_RESOURCE_DIRECTORY
    Dim resDat  As IMAGE_RESOURCE_DATA_ENTRY
    
    ' // Get main directory of resources
    oStream.BufferPointer = resIdx
    oStream.ReadData VarPtr(mainDir), Len(mainDir)
       
    resList.dwNumberOfTypes = mainDir.NumberOfIdEntries + mainDir.NumberOfNamedEntries
    
    If resList.dwNumberOfTypes <= 0 Then
        resList.dwNumberOfTypes = 0
        Exit Sub
    End If
    
    ReDim resList.ResourceTypes(resList.dwNumberOfTypes - 1)
    
    For types = 0 To resList.dwNumberOfTypes - 1
        
        With resList.ResourceTypes(types)
        
        ' // 1. Types of resources
        
        idx1 = resIdx + Len(mainDir) + types * Len(entry)
        
        oStream.BufferPointer = idx1
        oStream.ReadData VarPtr(entry), Len(entry)

        If entry.NameId < 0 Then
        
            index = (entry.NameId And &H7FFFFFFF) + resIdx
            oStream.BufferPointer = index
            oStream.ReadData VarPtr(strSize), 2
            
            .strType = Space$(strSize)
            oStream.BufferPointer = index + 2
            oStream.ReadData StrPtr(.strType), strSize * 2
            
        Else
        
            .strType = "#" & CStr(entry.NameId)
            
        End If
        
        If entry.OffsetToData < 0 Then
            ' // Is directory
            
            idx1 = resIdx + (entry.OffsetToData And &H7FFFFFFF)
            
            oStream.BufferPointer = idx1
            oStream.ReadData VarPtr(namDir), Len(namDir)
            
            ReDim .entData(namDir.NumberOfIdEntries + namDir.NumberOfNamedEntries - 1)
            
            For names = 0 To namDir.NumberOfIdEntries + namDir.NumberOfNamedEntries - 1
                
                With .entData(names)
                
                ' // 2. Identifiers of resources
                
                idx2 = idx1 + Len(namDir) + names * Len(entry)
                
                oStream.BufferPointer = idx2
                oStream.ReadData VarPtr(entry), Len(entry)

                If entry.NameId < 0 Then
                
                    index = (entry.NameId And &H7FFFFFFF) + resIdx
                    
                    oStream.BufferPointer = index
                    oStream.ReadData VarPtr(strSize), 2
            
                    .strName = Space$(strSize)
                    oStream.BufferPointer = index + 2
                    oStream.ReadData StrPtr(.strName), strSize * 2
                    
                Else
                
                    .strName = "#" & CStr(entry.NameId)
                    
                End If
        
                If entry.OffsetToData < 0 Then
                    ' // Is directory
                    
                    idx2 = resIdx + (entry.OffsetToData And &H7FFFFFFF)

                    oStream.BufferPointer = idx2
                    oStream.ReadData VarPtr(lngDir), Len(lngDir)
                
                    ReDim .lngData(lngDir.NumberOfIdEntries + lngDir.NumberOfNamedEntries - 1)
                    
                    For langs = 0 To lngDir.NumberOfIdEntries + lngDir.NumberOfNamedEntries - 1
                        
                        With .lngData(langs)
                        
                        ' // 3. Languages
                        
                        idx3 = idx2 + Len(lngDir) + langs * Len(entry)

                        oStream.BufferPointer = idx3
                        oStream.ReadData VarPtr(entry), Len(entry)
                        
                        If entry.OffsetToData > 0 Then
        
                            idx3 = resIdx + entry.OffsetToData

                            oStream.BufferPointer = idx3
                            oStream.ReadData VarPtr(resDat), Len(resDat)
                            
                            .LangID = entry.NameId
                            .dwSize = resDat.Size
                            
                            If .dwSize Then
                            
                                ' // Copy resource data
                                ReDim .binData(.dwSize - 1)
                                
                                index = RVA2RAW(resDat.OffsetToData, sec())
                                
                                oStream.BufferPointer = index
                                oStream.ReadData VarPtr(.binData(0)), .dwSize
                                
                            End If
                            
                        End If
                        
                        End With
                        
                    Next
                    
                End If
                
                End With
                
            Next
            
        End If
        
        End With
        
    Next

End Sub

' // RVA to RAW
Private Function RVA2RAW( _
                 ByVal rva As Long, _
                 ByRef sec() As IMAGE_SECTION_HEADER) As Long
    Dim index As Long
    
    For index = 0 To UBound(sec)
        
        If rva >= sec(index).VirtualAddress And _
           rva < sec(index).VirtualAddress + sec(index).VirtualSize Then
            RVA2RAW = sec(index).PointerToRawData + (rva - sec(index).VirtualAddress)
            Exit Function
        End If
        
    Next
    
    RVA2RAW = rva
    
End Function

' // Save stream data to file
Private Function SaveStreamToFile( _
                 ByRef sFileName As String, _
                 ByVal oStream As clsStream) As Long
    Dim hFile       As Long
    Dim binData()   As Byte
    Dim dwWritten   As Long
    Dim Ret         As Long
    
    On Error GoTo error_handler
    
    Err.Clear
    
    Ret = -1
    
    hFile = CreateFile(StrPtr(sFileName), GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 75
    End If
    
    If oStream.Size > 0 Then
        
        oStream.BufferPointer = 0
        
        ' // Read data to local buffer
        ReDim binData(oStream.Size - 1)
        
        oStream.ReadData VarPtr(binData(0)), oStream.Size
        
        If WriteFile(hFile, binData(0), oStream.Size, dwWritten, ByVal 0&) = 0 Then
            Err.Raise 3040
        End If
        
    End If
    
    Ret = dwWritten
    
error_handler:
    
    ' // Clean up
    CloseHandle hFile

    SaveStreamToFile = Ret
    
    If Err.Number Then Err.Raise Err.Number
    
End Function

' // Load file to byte stream. Returns size of written data
Private Function LoadFileToStream( _
                 ByRef sFileName As String, _
                 ByVal oStream As clsStream) As Long
    Dim hFile       As Long
    Dim sizeOfFile  As LARGE_INTEGER
    Dim binData()   As Byte
    Dim dwRead      As Long
    Dim Ret         As Long
    
    On Error GoTo error_handler
    
    Err.Clear
    
    Ret = -1
    
    hFile = CreateFile(StrPtr(sFileName), GENERIC_READ, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise 75
    End If

    ' // Check size
    GetFileSizeEx hFile, sizeOfFile
    If sizeOfFile.highpart <> 0 Or sizeOfFile.lowpart > 100000000 Or sizeOfFile.lowpart <= 0 Then
        Err.Raise 7
    End If
    
    ReDim binData(sizeOfFile.lowpart - 1)
    
    If ReadFile(hFile, binData(0), sizeOfFile.lowpart, dwRead, ByVal 0&) = 0 Then
        Err.Raise 3040
    End If

    ' // Save to stream
    oStream.WriteData VarPtr(binData(0)), dwRead
    
    Ret = dwRead
    
error_handler:
    
    ' // Clean up
    CloseHandle hFile
    
    LoadFileToStream = Ret
    
    If Err.Number Then Err.Raise Err.Number
    
End Function

' // Translate to relative path
Public Function ToRelative( _
                ByRef fullPath As String, _
                ByRef BaseDirectory As String) As String
                
    If PathIsRelative(StrPtr(fullPath)) Then ToRelative = fullPath: Exit Function
    
    ToRelative = Space(MAX_PATH)
    If PathRelativePathTo(StrPtr(ToRelative), _
                           StrPtr(BaseDirectory), _
                           vbDirectory, StrPtr(fullPath), 0) = 0 Then ToRelative = fullPath: Exit Function
                           
    ToRelative = Left$(ToRelative, InStr(1, ToRelative, vbNullChar) - 1)
    
End Function

' // Translate to absolute path
Public Function ToAbsolute( _
                ByRef relPath As String, _
                ByRef BaseDirectory As String) As String
    Dim sl As String
    
    If PathIsRelative(StrPtr(relPath)) = 0 Then ToAbsolute = relPath: Exit Function
    
    ToAbsolute = Space(MAX_PATH)
    If Right(BaseDirectory, 1) <> "\" Then sl = "\"
    If PathCanonicalize(StrPtr(ToAbsolute), StrPtr(BaseDirectory & sl & relPath)) = 0 Then ToAbsolute = vbNullString: Exit Function
    
    ToAbsolute = Left$(ToAbsolute, InStr(1, ToAbsolute, vbNullChar) - 1)
    
End Function

' // Get loader path
Private Function GetLoaderPath() As String
    Dim sFileName   As String
    Dim dwAttr      As Long
    
    ' // Get loader path
    sFileName = ToAbsolute("..\..\Launcher\Release\loader.exe", App.Path)
    
    ' // Determine if file exists
    dwAttr = GetFileAttributes(StrPtr(sFileName))
    
    If dwAttr <> -1 And (dwAttr And vbDirectory) = 0 Then
        GetLoaderPath = sFileName
    End If
    
End Function

' -
Private Sub Class_Initialize()
    Set mExecute = New clsExecute
    Set mExecute.Project = Me
    Set mStorage = New clsStorage
    Set mStorage.Project = Me
End Sub
