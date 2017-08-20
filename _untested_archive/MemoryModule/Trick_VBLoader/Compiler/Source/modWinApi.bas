Attribute VB_Name = "modWinApi"
' // modWinApi.bas - needed API function declarations, types and constants
' // © Krivous Anatoly Anatolevich (The trick), 2016

Option Explicit
          
Public Const MAX_PATH                        As Long = 260
Public Const RT_RCDATA                       As Long = 10&
Public Const GENERIC_WRITE                   As Long = &H40000000
Public Const GENERIC_READ                    As Long = &H80000000
Public Const CREATE_ALWAYS                   As Long = 2
Public Const OPEN_EXISTING                   As Long = 3
Public Const FILE_SHARE_READ                 As Long = &H1
Public Const FILE_ATTRIBUTE_NORMAL           As Long = &H80
Public Const INVALID_HANDLE_VALUE            As Long = -1
Public Const IMAGE_DOS_SIGNATURE             As Long = &H5A4D
Public Const IMAGE_NT_SIGNATURE              As Long = &H4550&
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC   As Long = &H10B&
Public Const IMAGE_DIRECTORY_ENTRY_RESOURCE  As Long = 2
Public Const RT_ICON                         As Long = 3
Public Const RT_GROUP_ICON                   As Long = RT_ICON + 11
Public Const RT_VERSION                      As Long = 16
Public Const OFN_ALLOWMULTISELECT            As Long = &H200
Public Const OFN_EXPLORER                    As Long = &H80000
Public Const OFN_OVERWRITEPROMPT             As Long = &H2
Public Const WM_GETMINMAXINFO                As Long = &H24
Public Const COMPRESSION_FORMAT_LZNT1        As Long = 2
Public Const LR_SHARED                       As Long = &H8000&
Public Const SM_CXSMICON                     As Long = 49
Public Const SM_CYSMICON                     As Long = 50
Public Const IMAGE_ICON                      As Long = 1
Public Const WM_SETICON                      As Long = &H80
Public Const ICON_SMALL                      As Long = 0

Public Type IMAGE_DOS_HEADER
    e_magic                     As Integer
    e_cblp                      As Integer
    e_cp                        As Integer
    e_crlc                      As Integer
    e_cparhdr                   As Integer
    e_minalloc                  As Integer
    e_maxalloc                  As Integer
    e_ss                        As Integer
    e_sp                        As Integer
    e_csum                      As Integer
    e_ip                        As Integer
    e_cs                        As Integer
    e_lfarlc                    As Integer
    e_ovno                      As Integer
    e_res(0 To 3)               As Integer
    e_oemid                     As Integer
    e_oeminfo                   As Integer
    e_res2(0 To 9)              As Integer
    e_lfanew                    As Long
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress              As Long
    Size                        As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUnitializedData       As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    BaseOfData                  As Long
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    W32VersionValue             As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    SubSystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(15)           As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_FILE_HEADER
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Public Type IMAGE_NT_HEADERS
    Signature                   As Long
    FileHeader                  As IMAGE_FILE_HEADER
    OptionalHeader              As IMAGE_OPTIONAL_HEADER
End Type

Public Type IMAGE_SECTION_HEADER
    SectionName(7)              As Byte
    VirtualSize                 As Long
    VirtualAddress              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Public Type IMAGE_RESOURCE_DIRECTORY
    Characteristics             As Long
    TimeDateStamp               As Long
    MajorVersion                As Integer
    MinorVersion                As Integer
    NumberOfNamedEntries        As Integer
    NumberOfIdEntries           As Integer
End Type

Public Type IMAGE_RESOURCE_DIRECTORY_ENTRY
    NameId                      As Long
    OffsetToData                As Long
End Type

Public Type IMAGE_RESOURCE_DATA_ENTRY
    OffsetToData                As Long
    Size                        As Long
    CodePage                    As Long
    Reserved                    As Long
End Type

Public Type LARGE_INTEGER
    lowpart                     As Long
    highpart                    As Long
End Type

Public Type OPENFILENAME
    lStructSize                 As Long
    hwndOwner                   As Long
    hInstance                   As Long
    lpstrFilter                 As Long
    lpstrCustomFilter           As Long
    nMaxCustFilter              As Long
    nFilterIndex                As Long
    lpstrFile                   As Long
    nMaxFile                    As Long
    lpstrFileTitle              As Long
    nMaxFileTitle               As Long
    lpstrInitialDir             As Long
    lpstrTitle                  As Long
    Flags                       As Long
    nFileOffset                 As Integer
    nFileExtension              As Integer
    lpstrDefExt                 As Long
    lCustData                   As Long
    lpfnHook                    As Long
    lpTemplateName              As Long
End Type

Public Type POINTAPI
    x                           As Long
    y                           As Long
End Type

Public Type MINMAXINFO
    ptReserved                  As POINTAPI
    ptMaxSize                   As POINTAPI
    ptMaxPosition               As POINTAPI
    ptMinTrackSize              As POINTAPI
    ptMaxTrackSize              As POINTAPI
End Type

Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
                        Alias "GetOpenFileNameW" ( _
                        ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" _
                        Alias "GetSaveFileNameW" ( _
                        ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetFileAttributes Lib "kernel32" _
                        Alias "GetFileAttributesW" ( _
                        ByVal lpFileName As Long) As Long
Public Declare Function GetFileSizeEx Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpFileSize As LARGE_INTEGER) As Long
Public Declare Function CreateFile Lib "kernel32" _
                        Alias "CreateFileW" ( _
                        ByVal lpFileName As Long, _
                        ByVal dwDesiredAccess As Long, _
                        ByVal dwShareMode As Long, _
                        ByRef lpSecurityAttributes As Any, _
                        ByVal dwCreationDisposition As Long, _
                        ByVal dwFlagsAndAttributes As Long, _
                        ByVal hTemplateFile As Long) As Long
Public Declare Function WriteFile Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal nNumberOfBytesToWrite As Long, _
                        ByRef lpNumberOfBytesWritten As Long, _
                        ByRef lpOverlapped As Any) As Long
Public Declare Function ReadFile Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        ByRef lpBuffer As Any, _
                        ByVal nNumberOfBytesToRead As Long, _
                        ByRef lpNumberOfBytesRead As Long, _
                        ByRef lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" ( _
                        ByVal hObject As Long) As Long
Public Declare Function BeginUpdateResource Lib "kernel32" _
                        Alias "BeginUpdateResourceW" ( _
                        ByVal pFileName As Long, _
                        ByVal bDeleteExistingResources As Long) As Long
Public Declare Function UpdateResource Lib "kernel32" _
                        Alias "UpdateResourceW" ( _
                        ByVal hUpdate As Long, _
                        ByVal lpType As Long, _
                        ByVal lpName As Long, _
                        ByVal wLanguage As Long, _
                        ByRef lpData As Any, _
                        ByVal cbData As Long) As Long
Public Declare Function EndUpdateResource Lib "kernel32" _
                        Alias "EndUpdateResourceW" ( _
                        ByVal hUpdate As Long, _
                        ByVal fDiscard As Long) As Long
Public Declare Function PathRelativePathTo Lib "Shlwapi.dll" _
                        Alias "PathRelativePathToW" ( _
                        ByVal pszPath As Long, _
                        ByVal pszFrom As Long, _
                        ByVal dwAttrFrom As Long, _
                        ByVal pszTo As Long, _
                        ByVal dwAttrTo As Long) As Long
Public Declare Function PathCanonicalize Lib "Shlwapi.dll" _
                        Alias "PathCanonicalizeW" ( _
                        ByVal lpszDst As Long, _
                        ByVal lpszSrc As Long) As Long
Public Declare Function PathIsRelative Lib "Shlwapi.dll" _
                        Alias "PathIsRelativeW" ( _
                        ByVal lpszPath As Long) As Long
Public Declare Function lstrcpyn Lib "kernel32" _
                        Alias "lstrcpynW" ( _
                        ByRef lpString1 As Any, _
                        ByRef lpString2 As Any, _
                        ByVal Length As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" ( _
                        ByRef lp As Any, _
                        ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" ( _
                        ByRef lp As Any, _
                        ByVal ucb As Long) As Long
Public Declare Function RtlGetCompressionWorkSpaceSize Lib "ntdll" ( _
                        ByVal CompressionFormatAndEngine As Integer, _
                        ByRef CompressBufferWorkSpaceSize As Long, _
                        ByRef CompressFragmentWorkSpaceSize As Long) As Long
Public Declare Function RtlCompressBuffer Lib "ntdll" ( _
                        ByVal CompressionFormatAndEngine As Integer, _
                        ByRef UncompressedBuffer As Any, _
                        ByVal UncompressedBufferSize As Long, _
                        ByRef CompressedBuffer As Any, _
                        ByVal CompressedBufferSize As Long, _
                        ByVal UncompressedChunkSize As Long, _
                        ByRef FinalCompressedSize As Long, _
                        ByRef WorkSpace As Any) As Long
Public Declare Function GetSystemMetrics Lib "user32" ( _
                        ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" _
                        Alias "SendMessageA" ( _
                        ByVal hWnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
Public Declare Function LoadImage Lib "user32" _
                        Alias "LoadImageA" ( _
                        ByVal hInst As Long, _
                        ByVal lpsz As String, _
                        ByVal uType As Long, _
                        ByVal cxDesired As Long, _
                        ByVal cyDesired As Long, _
                        ByVal fuLoad As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
                   Alias "RtlMoveMemory" ( _
                   ByRef Destination As Any, _
                   ByRef Source As Any, _
                   ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" _
                   Alias "RtlZeroMemory" ( _
                   ByRef dest As Any, _
                   ByVal numBytes As Long)



