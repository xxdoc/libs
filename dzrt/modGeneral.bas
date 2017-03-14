Attribute VB_Name = "modGeneral"
Option Explicit

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'just a useful bonus method to add into the C dll might as well...
Declare Function ut_crc32 Lib "utypes.dll" Alias "crc32" (ByRef b As Byte, ByVal sz As Long) As Long
Declare Function ut_crc32w Lib "utypes.dll" Alias "crc32w" (ByVal b As Long, ByVal sz As Long) As Long

'Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal flags As Integer, WorkSpaceSize As Long, UNKNOWN_PARAMETER As Long) As Long
'Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal flags As Long, ByVal ProtectMode As Long) As Long
'Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, ByVal UNKNOWN_PARAMETER As Long, OutputSize As Long, ByVal WorkSpace As Long) As Long
'Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, OutputSize As Long) As Long
'Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal flags As Long) As Long
 
Private Declare Function CryptBinaryToString Lib "Crypt32" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal flags As Integer, WorkSpaceSize As Long, UNKNOWN_PARAMETER As Long) As Long
Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, ByVal NumBits As Long, regionsize As Long, ByVal flags As Long, ByVal ProtectMode As Long) As Long
Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, ByVal UNKNOWN_PARAMETER As Long, OutputSize As Long, ByVal WorkSpace As Long) As Long
Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal flags As Integer, ByVal BuffUnCompressed As Long, ByVal UnCompSize As Long, ByVal BuffCompressed As Long, ByVal CompBuffSize As Long, OutputSize As Long) As Long
Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcHandle As Long, BaseAddress As Long, regionsize As Long, ByVal flags As Long) As Long
 
Public hUTypes As Long
 
Const STATUS_SUCCESS = 0
Const STATUS_BUFFER_ALL_ZEROS = &H117
Const STATUS_INVALID_PARAMETER = &HC000000D
Const STATUS_UNSUPPORTED_COMPRESSION = &HC000025F
Const STATUS_NOT_SUPPORTED_ON_SBS = &HC0000300
Const STATUS_BUFFER_TOO_SMALL = &HC0000023
Const STATUS_BAD_COMPRESSION_BUFFER = &HC0000242

Const COMPRESSION_FORMAT_LZNT1 = &H2
Const COMPRESSION_ENGINE_STANDARD = &H0   '// Standard compression
Const COMPRESSION_ENGINE_MAXIMUM = &H100  '// Maximum compression
 
 
Function ensureUTypes() As Boolean
    
    On Error Resume Next
    
    If hUTypes <> 0 Then
        ensureUTypes = True
        Exit Function
    End If
    
    Dim pth As String, b() As Byte, f As Long
    
    pth = App.path & "\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\UTypes.dll"
    If Not FileExists(pth) Then pth = App.path & "\..\..\..\UTypes.dll"
    
    If Not FileExists(pth) Then

        pth = App.path & "\UTypes.dll"
        b() = LoadResData("UTYPES", "DLLS")
        If AryIsEmpty(b) Then
            MsgBox "Failed to find UTypes.dll in resource?"
            Exit Function
        End If
        
        f = FreeFile
        Open pth For Binary As f
        Put f, , b()
        Close f
        
        'MsgBox "Dropped utypes.dll to: " & pth & " - Err: " & Err.Number
    End If
    
    If Not FileExists(pth) Then
        MsgBox "Failed to write UTypes.dll to disk from resource?"
        Exit Function
    End If
        
    hUTypes = LoadLibrary(pth)
        
    If hUTypes = 0 Then
        MsgBox "Failed to load UTypes.dll library?"
        Exit Function
    End If
    
    ensureUTypes = True
    
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x
  
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

'Public Function RTLCompress(data() As Byte, Out() As Byte) As Long
'   Dim WorkSpaceSize As Long
'   Dim WorkSpace As Long
'   Dim lOutputSize As Long
'
'   ReDim Out(UBound(data) * 1.13 + 4)
'   RtlGetCompressionWorkSpaceSize 2, WorkSpaceSize, 0
'   NtAllocateVirtualMemory -1, WorkSpace, 0, WorkSpaceSize, 4096, 64
'   RtlCompressBuffer 2, VarPtr(data(0)), UBound(data) + 1, VarPtr(Out(0)), (UBound(data) * 1.13 + 4), 0, lOutputSize, WorkSpace
'   NtFreeVirtualMemory -1, WorkSpace, 0, 16384
'   ReDim Preserve Out(lOutputSize)
'   RTLCompress = lOutputSize
'
'End Function
'
'Public Function RTLDeCompress(data() As Byte, dest() As Byte) As Long
'   If UBound(data) Then
'       Dim lBufferSize As Long
'       ReDim dest(UBound(data) * 12.5)
'       RtlDecompressBuffer 2, VarPtr(dest(0)), (UBound(data) * 12.5), VarPtr(data(0)), UBound(data), lBufferSize
'       If lBufferSize Then
'           ReDim Preserve dest(lBufferSize - 1)
'           RTLDeCompress = lBufferSize - 1
'       End If
'   End If
'End Function

Public Function RTLCompress(Data() As Byte, bOut() As Byte, Optional max As Boolean = False) As Boolean
    Dim WorkSpaceSize As Long
    Dim WorkSpace As Long
    Dim lCompress As Long
    Dim flags As Long
    Dim ret As Long
    
    If AryIsEmpty(Data) Then Exit Function

    flags = COMPRESSION_FORMAT_LZNT1
    If max Then flags = flags Or COMPRESSION_ENGINE_MAXIMUM
    
    ReDim bOut(UBound(Data) * 1.13 + 4)
    RtlGetCompressionWorkSpaceSize 2, WorkSpaceSize, 0
    NtAllocateVirtualMemory -1, WorkSpace, 0, WorkSpaceSize, 4096, 64
    ret = RtlCompressBuffer(flags, VarPtr(Data(0)), UBound(Data) + 1, VarPtr(bOut(0)), (UBound(Data) * 1.13 + 4), 0, lCompress, WorkSpace)
    NtFreeVirtualMemory -1, WorkSpace, 0, 16384
    
    If ret = STATUS_SUCCESS Then
        RTLCompress = True
        ReDim Preserve bOut(lCompress + 1)
    Else
        Erase bOut
    End If
    
    
'    If CryptBinaryToString(Out(0), lCompress + 1, &H4, 0&, WorkSpaceSize) <> 0 Then
'        strHexOut = String(WorkSpaceSize - 1, 0)
'        If CryptBinaryToString(Out(0), lCompress + 1, &H4, StrPtr(strHexOut), WorkSpaceSize) <> 0 Then
'            strHexOut = Replace$(Replace$(Replace$(strHexOut, " ", vbNullString), vbNewLine, vbNullString), vbTab, vbNullString)
'            CompressToHex = True
'        End If
'    End If

End Function

Public Function RTLDeCompress(ByRef bytBuf() As Byte, BitOut() As Byte, Optional max As Boolean = False) As Boolean
    Dim lBufferSize As Long
    Dim WorkSpaceSize As Long
    Dim dwActualUsed As Long
    Dim ret As Long
    Dim flags As Long
    
    If AryIsEmpty(bytBuf) Then Exit Function
    
    flags = COMPRESSION_FORMAT_LZNT1
    If max Then flags = flags Or COMPRESSION_ENGINE_MAXIMUM
    
    ReDim BitOut(UBound(bytBuf) * 12.5)
    ret = RtlDecompressBuffer(flags, VarPtr(BitOut(0)), (UBound(bytBuf) * 12.5), VarPtr(bytBuf(0)), UBound(bytBuf), lBufferSize)
    
    If ret = STATUS_SUCCESS And lBufferSize Then
        ReDim Preserve BitOut(lBufferSize - 1)
        RTLDeCompress = True
    Else
        Erase BitOut
        Exit Function
    End If
    
End Function


