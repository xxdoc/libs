VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemMapFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyToMem Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, source As Any, ByVal Length As Long)
Private Declare Sub CopyFromMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal source As Long, ByVal Length As Long)

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const PAGE_READWRITE = &H4
Private Const SECTION_MAP_WRITE = &H2
Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SECTION_QUERY As Long = &H1
Private Const SECTION_MAP_READ As Long = &H4
Private Const SECTION_MAP_EXECUTE As Long = &H8
Private Const SECTION_EXTEND_SIZE As Long = &H10

Private Const SECTION_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or _
                                           SECTION_MAP_WRITE Or SECTION_MAP_READ Or _
                                           SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
                                           
Private Const FILE_MAP_ALL_ACCESS As Long = SECTION_ALL_ACCESS


Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Dim hFile As Long
Dim hFileWnd As Long
Dim Max As Long
Dim vFName As String
Dim gAddr As Long

Public DebugMode As Boolean
Public ErrorMessage As String

Public Property Get FileHandle() As Long
    FileHandle = hFile
End Property

Public Property Get FileName() As String
   FileName = vFName
End Property

Public Property Get MaxSize() As Long
   MaxSize = Max
End Property

Function CreateMemMapFile(fName As String, mSize As Long, Optional initilize As Boolean = False) As Boolean

     
    
    Dim T As SECURITY_ATTRIBUTES
    vFName = UCase$(fName)
    Max = mSize
    T.nLength = Len(T)
    T.bInheritHandle = 1
    
    If hFile <> 0 Then
        ErrorMessage = "Cannot open multiple virtural files with one class"
        Exit Function
    End If
    
    hFile = CreateFileMapping(&HFFFFFFFF, T, PAGE_READWRITE, 0, mSize, vFName)
    
    If hFile = 0 Then
        ErrorMessage = "Unable to create virtual file"
        Exit Function
    End If
    
    If hFile <> 0 Then
        Dim b() As Byte
        ReDim b(mSize)
        gAddr = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, mSize)
        
        If DebugMode Then Debug.Print "Write Address: " & Hex(gAddr) & " b(0) = " & Hex(VarPtr(b(0)))
        
        If initilize Then CopyToMem gAddr, b(0), mSize
        CreateMemMapFile = True
    End If
    
End Function

Function ReadAll(strOut As String) As Boolean
       
    If hFile = 0 Then
        ErrorMessage = "Virtual File or Virtual File Interface not initialized"
        Exit Function
    End If
    
    ReadAll = ReadLength(strOut, Max)
    
End Function

Function ReadLength(strOut As String, nBytes As Long, Optional Offset As Long = 0) As Boolean

    If hFile = 0 Then
        ErrorMessage = "Virtual File or Virtual File Interface not initialized"
        Exit Function
    End If
    
    If (Offset + nBytes) > Max Then
        ErrorMessage = "Invalid size and offset combination"
        Exit Function
    End If
   
    Dim b() As Byte
    ReDim b(nBytes)
    
    If DebugMode Then Debug.Print "Reading Address: " & Hex(gAddr) & " b(0) = " & Hex(VarPtr(b(0)))
    
    CopyFromMem b(0), gAddr + Offset, nBytes
    strOut = StrConv(b(), vbUnicode)
    ReadLength = True
    
End Function

Function WriteFile(bData As String, Optional Offset As Long = 0, Optional ErasePrevData As Boolean = False) As Boolean

    If bData = "" Then Exit Function
    
    If hFile = 0 Then
        ErrorMessage = "Virtual File or Virtual File Interface not initialized"
        Exit Function
    End If

    If (Offset + Len(bData)) > Max Then
        ErrorMessage = "Invalid data chunk size and offset combination"
        Exit Function
    End If
    
    Dim b() As Byte
    b() = StrConv(bData, vbFromUnicode)
    
    If DebugMode Then Debug.Print "Write Address: " & Hex(gAddr + Offset) & " b(0) = " & Hex(VarPtr(b(0)))
    
    CopyToMem gAddr + Offset, b(0), UBound(b)
    
    WriteFile = True
    
End Function

Function ReMap() As Boolean
    If gAddr <> 0 Then UnMap
    gAddr = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, Max)
    If gAddr <> 0 Then ReMap = True
End Function

Function UnMap()
        UnmapViewOfFile gAddr
        gAddr = 0
End Function

Sub CloseFile()

    If hFile = 0 Then
        Exit Sub
    End If
    
    CloseHandle hFile
    hFile = 0
    Max = 0
    vFName = ""
    
End Sub

Private Sub Class_Terminate()
    UnMap
    CloseFile
End Sub

