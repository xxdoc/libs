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
Public Declare Sub CopyMemory_ Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
 
Public hUTypes As Long
 
Global Const LANG_US = &H409
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
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
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

 Sub CopyMemory(Destination As Long, Source As Long, ByVal length As Long)
     CopyMemory_ ByVal Destination, ByVal Source, length
 End Sub
 
'this function will convert any of the following to a byte array:
'   read a file if path supplied and allowFilePaths = true
'   byte(), integer() or long() arrays
'   all other data types it will attempt to convert them to string, then to byte array
'   if the data type you pass can not be converted with cstr() it will throw an error.
'   no other types make sense to support explicitly
'   this assumes all arrays are 0 based..
Function LoadData(fileStringOrByte, Optional allowFilePaths As Boolean = True) As Byte()
    
    Dim f As Long
    Dim size As Long
    Dim b() As Byte
    Dim l() As Long    ' must cast to specific array type or
    Dim i() As Integer ' else you are reading part of the variant structure..
    
    If allowFilePaths And FileExists(fileStringOrByte) Then
         f = FreeFile
         Open fileStringOrByte For Binary As f
         ReDim b(LOF(f) - 1)
         Get f, , b()
         Close f
    ElseIf TypeName(fileStringOrByte) = "Byte()" Then
        b() = fileStringOrByte
    ElseIf TypeName(fileStringOrByte) = "Integer()" Then
        i() = fileStringOrByte
        ReDim b((UBound(i) * 2) - 1)
        CopyMemory VarPtr(b(0)), VarPtr(i(0)), UBound(b) + 1
    ElseIf TypeName(fileStringOrByte) = "Long()" Then
        l() = fileStringOrByte
        ReDim b((UBound(l) * 4) - 1)
        CopyMemory VarPtr(b(0)), VarPtr(l(0)), UBound(b) + 1
    Else
        b() = StrConv(CStr(fileStringOrByte), vbFromUnicode, LANG_US)
    End If
    
    LoadData = b()
    
End Function


'ported from Detect It Easy - Binary::calculateEntropy
'   https://github.com/horsicq/DIE-engine/blob/master/binary.cpp#L2319
Function fileEntropy(pth As String, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Dim nSize As Long
    Dim nTemp As Long
    Const BUFFER_SIZE = &H1000
    Dim buf() As Byte
    Dim f As Long
    Dim i As Long
    
    On Error Resume Next
    
    f = FreeFile
    Open pth For Binary Access Read As f
    If Err.Number <> 0 Then GoTo ret0
    
    sz = LOF(f) - 1
    
    If leng = 0 Then GoTo ret0
    
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    Seek f, offset
    nSize = leng
    fEntropy = 1.44269504088896
    ReDim buf(BUFFER_SIZE)
    
    'read the file in chunks and count how many times each byte value occurs
    While (nSize > 0)
        nTemp = IIf(nSize < BUFFER_SIZE, nSize, BUFFER_SIZE)
        If nTemp <> BUFFER_SIZE Then ReDim buf(nTemp) 'last chunk, partial buffer
        Get f, , buf()
        For i = 0 To UBound(buf)
            bytes(buf(i)) = bytes(buf(i)) + 1
        Next
        nSize = nSize - nTemp
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    Close f
    fileEntropy = Round(fEntropy / CSng(leng), 3)
    
Exit Function
ret0:
    Close f
End Function


Function memEntropy(buf() As Byte, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    Dim sz As Long
    Dim fEntropy As Single
    Dim bytes(255) As Single
    Dim temp As Single
    Const BUFFER_SIZE = &H1000
    Dim i As Long
    
    sz = UBound(buf)
    
    If leng = 0 Then GoTo ret0
    If leng = -1 Then
        leng = sz - offset
        If leng = 0 Then GoTo ret0
    End If
    
    If offset >= sz Then GoTo ret0
    If offset + leng > sz Then GoTo ret0
    
    fEntropy = 1.44269504088896
    
    While (offset < sz)
        'count each byte value occurance
        bytes(buf(offset)) = bytes(buf(offset)) + 1
        offset = offset + 1
    Wend
    
    For i = 0 To UBound(bytes)
        temp = bytes(i) / CSng(leng)
        If temp <> 0 Then
            fEntropy = fEntropy + (-Log(temp) / Log(2)) * bytes(i)
        End If
    Next
    
    memEntropy = Round(fEntropy / CSng(leng), 3)
    
Exit Function
ret0:
End Function




