Attribute VB_Name = "modPEStructs"
Option Explicit
'these struct definitions were taken from VB debugger sample by VF-fCRO
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42422&lngWId=1

Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Public Const IMAGE_SIZEOF_SHORT_NAME = 8
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B

Public Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public Type IMAGE_IMPORT_DIRECTORY
    pFuncAry As Long
    timestamp As Long
    forwarder As Long
    pDllName As Long
    pThunk As Long
End Type

Public Type IMAGE_SECTION_HEADER 'https://msdn.microsoft.com/en-us/library/windows/desktop/ms680341(v=vs.85).aspx
    nameSec As String * 6
    PhisicalAddress As Integer
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long 'rva
    size As Long
End Type


Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_OPTIONAL_HEADER_64
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    'BaseOfData As Long                        'this was removed for pe32+
    ImageBase As Currency                        'changed
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Currency                         'changed
    SizeOfStackCommit As Currency                         'changed
    SizeOfHeapReserve As Currency                         'changed
    SizeOfHeapCommit As Currency                        'changed
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

'Enum eDATA_DIRECTORY
'    Export_Table = 0
'    Import_Table = 1
'    Resource_Table = 2
'    Exception_Table = 3
'    Certificate_Table = 4
'    Relocation_Table = 5
'    Debug_Data = 6
'    Architecture_Data = 7
'    Machine_Value = 8        '(MIPS_GP)
'    TLS_Table = 9
'    Load_Configuration_Table = 10
'    Bound_Import_Table = 11
'    Import_Address_Table = 12
'    Delay_Import_Descriptor = 13
'    COM_Runtime_Header = 14
'    Reserved = 15
'End Enum


Public Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
    'OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public Type RESDIRECTORY
   Characteristics As Long
   TimeDateStamp As Long
   MajorVersion As Integer
   MinorVersion As Integer
   NumberOfNamedEntries As Integer
   NumberOfIdEntries As Integer
End Type

Public Type RESOURCE_DATAENTRY
   Data_RVA As Long
   size As Long
   CodePage As Long
   Reserved As Long
End Type

Public Type RESOURCE_DIRECTORY_ENTRY
    NameOffset_or_ID As Long          'which is based on if loaded from named entry or id entry list
    DataEntry_orSubDir_Offset As Long 'if highbit=1 then its SubDir offset else direct link to a dataentry
End Type


Private c_ole As Collection
Private c_ws2 As Collection


Function toHex(ParamArray elems())
    On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(elems)
        elems(i).Text = Hex(elems(i).Text)
    Next
End Function



Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub

Function GetHextxt(t As TextBox, v As Long) As Boolean
    
    On Error Resume Next
    v = CLng("&h" & t)
    If Err.Number > 0 Then
        MsgBox "Error " & t.Text & " is not valid hex number", vbInformation
        Exit Function
    End If
    
    GetHextxt = True
    
End Function

Function timeStampToDate(timestamp As Long) As String

    On Error Resume Next
    Dim base As Date
    Dim compiled As Date
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", timestamp, base)
    timeStampToDate = Format(compiled, "mmm d yyyy h:nn:ss") 'compatiable with cdate()
    '"GMT: " & Format(compiled, "ddd mmm d h:nn:ss yyyy")

End Function


Sub Enable(t As TextBox, Optional enabled = True)
    t.BackColor = IIf(enabled, vbWhite, &H80000004)
    t.enabled = enabled
    t.Text = Empty
End Sub

Function Align(ByVal valu) As Long
    While valu Mod 16 <> 0
        valu = valu + 1
    Wend
    Align = valu
End Function

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function


Sub ConfigureListView(lv As Object)
        
        Dim i As Integer
        
        lv.FullRowSelect = True
        lv.GridLines = True
        lv.HideColumnHeaders = False
        lv.View = 3 'lvwReport
    
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "Section Name"
        lv.ColumnHeaders.Add , , "VirtualAddr"
        lv.ColumnHeaders.Add , , "VirtualSize"
        lv.ColumnHeaders.Add , , "RawOffset"
        lv.ColumnHeaders.Add , , "RawSize"
        lv.ColumnHeaders.Add , , "Characteristics"
        
        lv.Width = (1250 * 6) + 250
        lv.Height = 1800
        
        For i = 1 To 6
            lv.ColumnHeaders(i).Width = 1250
        Next
        
End Sub

Sub FilloutListView(lv As Object, Sections As Collection)
        
    If Sections.Count = 0 Then
        MsgBox "Sections not loaded yet"
        Exit Sub
    End If
    
    Dim cs As CSection, li As Object 'ListItem
    lv.ListItems.Clear
    
    For Each cs In Sections
        Set li = lv.ListItems.Add(, , cs.nameSec)
        li.SubItems(1) = Hex(cs.VirtualAddress)
        li.SubItems(2) = Hex(cs.VirtualSize)
        li.SubItems(3) = Hex(cs.PointerToRawData)
        li.SubItems(4) = Hex(cs.SizeOfRawData)
        li.SubItems(5) = Hex(cs.Characteristics)
    Next
    
    Dim i As Integer
    For i = 1 To lv.ColumnHeaders.Count
        lv.ColumnHeaders(i).Width = 1000
    Next
    With lv.ColumnHeaders(i - 1)
        .Width = lv.Width - .Left - 100
    End With
    
    
End Sub

Function HexDump2(b() As Byte, Optional hexOnly = 0) As String
    Dim tmp As String
    tmp = StrConv(b, vbUnicode)
    HexDump2 = HexDump(tmp, hexOnly)
End Function

Function HexDump(ByVal str, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, x
    
    offset = 0
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Function rpad(v, Optional l As Long = 10)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        rpad = v & String(l - x, " ")
    Else
hell:
        rpad = v
    End If
End Function

'https://github.com/erocarrera/pefile/blob/8d60469de3b70109ac603c68c48fb3e7b84261e8/ordlookup/__init__.py
Function ordLookup(dll, ord)
    
    On Error Resume Next
    Dim Name As String
    
    'our ordinals are in hex..we need dec
    ord = CLng("&h" & Replace(ord, "@", Empty))
    If Err.Number <> 0 Then Err.Raise 1, "ordLookup", "Could not convert ord to long: " & ord
    
    If dll = "ws2_32" Or dll = "wsock32" Or dll = "oleaut32" Then
        If c_ole Is Nothing Then ordInit
        If dll = "oleaut32" Then
            Name = c_ole("ord:" & ord)
        Else
            Name = c_ws2("ord:" & ord)
        End If
    End If
    
    If Len(Name) = 0 Or Err.Number <> 0 Then
         Name = "ord" & ord
    End If
        
     ordLookup = LCase(Name)
    
End Function

Private Sub ordInit()
    
    Set c_ws2 = New Collection
    Set c_ole = New Collection
    
    c_ws2.Add "accept", "ord:1"
    c_ws2.Add "bind", "ord:2"
    c_ws2.Add "closesocket", "ord:3"
    c_ws2.Add "connect", "ord:4"
    c_ws2.Add "getpeername", "ord:5"
    c_ws2.Add "getsockname", "ord:6"
    c_ws2.Add "getsockopt", "ord:7"
    c_ws2.Add "htonl", "ord:8"
    c_ws2.Add "htons", "ord:9"
    c_ws2.Add "ioctlsocket", "ord:10"
    c_ws2.Add "inet_addr", "ord:11"
    c_ws2.Add "inet_ntoa", "ord:12"
    c_ws2.Add "listen", "ord:13"
    c_ws2.Add "ntohl", "ord:14"
    c_ws2.Add "ntohs", "ord:15"
    c_ws2.Add "recv", "ord:16"
    c_ws2.Add "recvfrom", "ord:17"
    c_ws2.Add "select", "ord:18"
    c_ws2.Add "send", "ord:19"
    c_ws2.Add "sendto", "ord:20"
    c_ws2.Add "setsockopt", "ord:21"
    c_ws2.Add "shutdown", "ord:22"
    c_ws2.Add "socket", "ord:23"
    c_ws2.Add "GetAddrInfoW", "ord:24"
    c_ws2.Add "GetNameInfoW", "ord:25"
    c_ws2.Add "WSApSetPostRoutine", "ord:26"
    c_ws2.Add "FreeAddrInfoW", "ord:27"
    c_ws2.Add "WPUCompleteOverlappedRequest", "ord:28"
    c_ws2.Add "WSAAccept", "ord:29"
    c_ws2.Add "WSAAddressToStringA", "ord:30"
    c_ws2.Add "WSAAddressToStringW", "ord:31"
    c_ws2.Add "WSACloseEvent", "ord:32"
    c_ws2.Add "WSAConnect", "ord:33"
    c_ws2.Add "WSACreateEvent", "ord:34"
    c_ws2.Add "WSADuplicateSocketA", "ord:35"
    c_ws2.Add "WSADuplicateSocketW", "ord:36"
    c_ws2.Add "WSAEnumNameSpaceProvidersA", "ord:37"
    c_ws2.Add "WSAEnumNameSpaceProvidersW", "ord:38"
    c_ws2.Add "WSAEnumNetworkEvents", "ord:39"
    c_ws2.Add "WSAEnumProtocolsA", "ord:40"
    c_ws2.Add "WSAEnumProtocolsW", "ord:41"
    c_ws2.Add "WSAEventSelect", "ord:42"
    c_ws2.Add "WSAGetOverlappedResult", "ord:43"
    c_ws2.Add "WSAGetQOSByName", "ord:44"
    c_ws2.Add "WSAGetServiceClassInfoA", "ord:45"
    c_ws2.Add "WSAGetServiceClassInfoW", "ord:46"
    c_ws2.Add "WSAGetServiceClassNameByClassIdA", "ord:47"
    c_ws2.Add "WSAGetServiceClassNameByClassIdW", "ord:48"
    c_ws2.Add "WSAHtonl", "ord:49"
    c_ws2.Add "WSAHtons", "ord:50"
    c_ws2.Add "gethostbyaddr", "ord:51"
    c_ws2.Add "gethostbyname", "ord:52"
    c_ws2.Add "getprotobyname", "ord:53"
    c_ws2.Add "getprotobynumber", "ord:54"
    c_ws2.Add "getservbyname", "ord:55"
    c_ws2.Add "getservbyport", "ord:56"
    c_ws2.Add "gethostname", "ord:57"
    c_ws2.Add "WSAInstallServiceClassA", "ord:58"
    c_ws2.Add "WSAInstallServiceClassW", "ord:59"
    c_ws2.Add "WSAIoctl", "ord:60"
    c_ws2.Add "WSAJoinLeaf", "ord:61"
    c_ws2.Add "WSALookupServiceBeginA", "ord:62"
    c_ws2.Add "WSALookupServiceBeginW", "ord:63"
    c_ws2.Add "WSALookupServiceEnd", "ord:64"
    c_ws2.Add "WSALookupServiceNextA", "ord:65"
    c_ws2.Add "WSALookupServiceNextW", "ord:66"
    c_ws2.Add "WSANSPIoctl", "ord:67"
    c_ws2.Add "WSANtohl", "ord:68"
    c_ws2.Add "WSANtohs", "ord:69"
    c_ws2.Add "WSAProviderConfigChange", "ord:70"
    c_ws2.Add "WSARecv", "ord:71"
    c_ws2.Add "WSARecvDisconnect", "ord:72"
    c_ws2.Add "WSARecvFrom", "ord:73"
    c_ws2.Add "WSARemoveServiceClass", "ord:74"
    c_ws2.Add "WSAResetEvent", "ord:75"
    c_ws2.Add "WSASend", "ord:76"
    c_ws2.Add "WSASendDisconnect", "ord:77"
    c_ws2.Add "WSASendTo", "ord:78"
    c_ws2.Add "WSASetEvent", "ord:79"
    c_ws2.Add "WSASetServiceA", "ord:80"
    c_ws2.Add "WSASetServiceW", "ord:81"
    c_ws2.Add "WSASocketA", "ord:82"
    c_ws2.Add "WSASocketW", "ord:83"
    c_ws2.Add "WSAStringToAddressA", "ord:84"
    c_ws2.Add "WSAStringToAddressW", "ord:85"
    c_ws2.Add "WSAWaitForMultipleEvents", "ord:86"
    c_ws2.Add "WSCDeinstallProvider", "ord:87"
    c_ws2.Add "WSCEnableNSProvider", "ord:88"
    c_ws2.Add "WSCEnumProtocols", "ord:89"
    c_ws2.Add "WSCGetProviderPath", "ord:90"
    c_ws2.Add "WSCInstallNameSpace", "ord:91"
    c_ws2.Add "WSCInstallProvider", "ord:92"
    c_ws2.Add "WSCUnInstallNameSpace", "ord:93"
    c_ws2.Add "WSCUpdateProvider", "ord:94"
    c_ws2.Add "WSCWriteNameSpaceOrder", "ord:95"
    c_ws2.Add "WSCWriteProviderOrder", "ord:96"
    c_ws2.Add "freeaddrinfo", "ord:97"
    c_ws2.Add "getaddrinfo", "ord:98"
    c_ws2.Add "getnameinfo", "ord:99"
    c_ws2.Add "WSAAsyncSelect", "ord:101"
    c_ws2.Add "WSAAsyncGetHostByAddr", "ord:102"
    c_ws2.Add "WSAAsyncGetHostByName", "ord:103"
    c_ws2.Add "WSAAsyncGetProtoByNumber", "ord:104"
    c_ws2.Add "WSAAsyncGetProtoByName", "ord:105"
    c_ws2.Add "WSAAsyncGetServByPort", "ord:106"
    c_ws2.Add "WSAAsyncGetServByName", "ord:107"
    c_ws2.Add "WSACancelAsyncRequest", "ord:108"
    c_ws2.Add "WSASetBlockingHook", "ord:109"
    c_ws2.Add "WSAUnhookBlockingHook", "ord:110"
    c_ws2.Add "WSAGetLastError", "ord:111"
    c_ws2.Add "WSASetLastError", "ord:112"
    c_ws2.Add "WSACancelBlockingCall", "ord:113"
    c_ws2.Add "WSAIsBlocking", "ord:114"
    c_ws2.Add "WSAStartup", "ord:115"
    c_ws2.Add "WSACleanup", "ord:116"
    c_ws2.Add "__WSAFDIsSet", "ord:151"
    c_ws2.Add "WEP", "ord:500"
    
    c_ole.Add "SysAllocString", "ord:2"
    c_ole.Add "SysReAllocString", "ord:3"
    c_ole.Add "SysAllocStringLen", "ord:4"
    c_ole.Add "SysReAllocStringLen", "ord:5"
    c_ole.Add "SysFreeString", "ord:6"
    c_ole.Add "SysStringLen", "ord:7"
    c_ole.Add "VariantInit", "ord:8"
    c_ole.Add "VariantClear", "ord:9"
    c_ole.Add "VariantCopy", "ord:10"
    c_ole.Add "VariantCopyInd", "ord:11"
    c_ole.Add "VariantChangeType", "ord:12"
    c_ole.Add "VariantTimeToDosDateTime", "ord:13"
    c_ole.Add "DosDateTimeToVariantTime", "ord:14"
    c_ole.Add "SafeArrayCreate", "ord:15"
    c_ole.Add "SafeArrayDestroy", "ord:16"
    c_ole.Add "SafeArrayGetDim", "ord:17"
    c_ole.Add "SafeArrayGetElemsize", "ord:18"
    c_ole.Add "SafeArrayGetUBound", "ord:19"
    c_ole.Add "SafeArrayGetLBound", "ord:20"
    c_ole.Add "SafeArrayLock", "ord:21"
    c_ole.Add "SafeArrayUnlock", "ord:22"
    c_ole.Add "SafeArrayAccessData", "ord:23"
    c_ole.Add "SafeArrayUnaccessData", "ord:24"
    c_ole.Add "SafeArrayGetElement", "ord:25"
    c_ole.Add "SafeArrayPutElement", "ord:26"
    c_ole.Add "SafeArrayCopy", "ord:27"
    c_ole.Add "DispGetParam", "ord:28"
    c_ole.Add "DispGetIDsOfNames", "ord:29"
    c_ole.Add "DispInvoke", "ord:30"
    c_ole.Add "CreateDispTypeInfo", "ord:31"
    c_ole.Add "CreateStdDispatch", "ord:32"
    c_ole.Add "RegisterActiveObject", "ord:33"
    c_ole.Add "RevokeActiveObject", "ord:34"
    c_ole.Add "GetActiveObject", "ord:35"
    c_ole.Add "SafeArrayAllocDescriptor", "ord:36"
    c_ole.Add "SafeArrayAllocData", "ord:37"
    c_ole.Add "SafeArrayDestroyDescriptor", "ord:38"
    c_ole.Add "SafeArrayDestroyData", "ord:39"
    c_ole.Add "SafeArrayRedim", "ord:40"
    c_ole.Add "SafeArrayAllocDescriptorEx", "ord:41"
    c_ole.Add "SafeArrayCreateEx", "ord:42"
    c_ole.Add "SafeArrayCreateVectorEx", "ord:43"
    c_ole.Add "SafeArraySetRecordInfo", "ord:44"
    c_ole.Add "SafeArrayGetRecordInfo", "ord:45"
    c_ole.Add "VarParseNumFromStr", "ord:46"
    c_ole.Add "VarNumFromParseNum", "ord:47"
    c_ole.Add "VarI2FromUI1", "ord:48"
    c_ole.Add "VarI2FromI4", "ord:49"
    c_ole.Add "VarI2FromR4", "ord:50"
    c_ole.Add "VarI2FromR8", "ord:51"
    c_ole.Add "VarI2FromCy", "ord:52"
    c_ole.Add "VarI2FromDate", "ord:53"
    c_ole.Add "VarI2FromStr", "ord:54"
    c_ole.Add "VarI2FromDisp", "ord:55"
    c_ole.Add "VarI2FromBool", "ord:56"
    c_ole.Add "SafeArraySetIID", "ord:57"
    c_ole.Add "VarI4FromUI1", "ord:58"
    c_ole.Add "VarI4FromI2", "ord:59"
    c_ole.Add "VarI4FromR4", "ord:60"
    c_ole.Add "VarI4FromR8", "ord:61"
    c_ole.Add "VarI4FromCy", "ord:62"
    c_ole.Add "VarI4FromDate", "ord:63"
    c_ole.Add "VarI4FromStr", "ord:64"
    c_ole.Add "VarI4FromDisp", "ord:65"
    c_ole.Add "VarI4FromBool", "ord:66"
    c_ole.Add "SafeArrayGetIID", "ord:67"
    c_ole.Add "VarR4FromUI1", "ord:68"
    c_ole.Add "VarR4FromI2", "ord:69"
    c_ole.Add "VarR4FromI4", "ord:70"
    c_ole.Add "VarR4FromR8", "ord:71"
    c_ole.Add "VarR4FromCy", "ord:72"
    c_ole.Add "VarR4FromDate", "ord:73"
    c_ole.Add "VarR4FromStr", "ord:74"
    c_ole.Add "VarR4FromDisp", "ord:75"
    c_ole.Add "VarR4FromBool", "ord:76"
    c_ole.Add "SafeArrayGetVartype", "ord:77"
    c_ole.Add "VarR8FromUI1", "ord:78"
    c_ole.Add "VarR8FromI2", "ord:79"
    c_ole.Add "VarR8FromI4", "ord:80"
    c_ole.Add "VarR8FromR4", "ord:81"
    c_ole.Add "VarR8FromCy", "ord:82"
    c_ole.Add "VarR8FromDate", "ord:83"
    c_ole.Add "VarR8FromStr", "ord:84"
    c_ole.Add "VarR8FromDisp", "ord:85"
    c_ole.Add "VarR8FromBool", "ord:86"
    c_ole.Add "VarFormat", "ord:87"
    c_ole.Add "VarDateFromUI1", "ord:88"
    c_ole.Add "VarDateFromI2", "ord:89"
    c_ole.Add "VarDateFromI4", "ord:90"
    c_ole.Add "VarDateFromR4", "ord:91"
    c_ole.Add "VarDateFromR8", "ord:92"
    c_ole.Add "VarDateFromCy", "ord:93"
    c_ole.Add "VarDateFromStr", "ord:94"
    c_ole.Add "VarDateFromDisp", "ord:95"
    c_ole.Add "VarDateFromBool", "ord:96"
    c_ole.Add "VarFormatDateTime", "ord:97"
    c_ole.Add "VarCyFromUI1", "ord:98"
    c_ole.Add "VarCyFromI2", "ord:99"
    c_ole.Add "VarCyFromI4", "ord:100"
    c_ole.Add "VarCyFromR4", "ord:101"
    c_ole.Add "VarCyFromR8", "ord:102"
    c_ole.Add "VarCyFromDate", "ord:103"
    c_ole.Add "VarCyFromStr", "ord:104"
    c_ole.Add "VarCyFromDisp", "ord:105"
    c_ole.Add "VarCyFromBool", "ord:106"
    c_ole.Add "VarFormatNumber", "ord:107"
    c_ole.Add "VarBstrFromUI1", "ord:108"
    c_ole.Add "VarBstrFromI2", "ord:109"
    c_ole.Add "VarBstrFromI4", "ord:110"
    c_ole.Add "VarBstrFromR4", "ord:111"
    c_ole.Add "VarBstrFromR8", "ord:112"
    c_ole.Add "VarBstrFromCy", "ord:113"
    c_ole.Add "VarBstrFromDate", "ord:114"
    c_ole.Add "VarBstrFromDisp", "ord:115"
    c_ole.Add "VarBstrFromBool", "ord:116"
    c_ole.Add "VarFormatPercent", "ord:117"
    c_ole.Add "VarBoolFromUI1", "ord:118"
    c_ole.Add "VarBoolFromI2", "ord:119"
    c_ole.Add "VarBoolFromI4", "ord:120"
    c_ole.Add "VarBoolFromR4", "ord:121"
    c_ole.Add "VarBoolFromR8", "ord:122"
    c_ole.Add "VarBoolFromDate", "ord:123"
    c_ole.Add "VarBoolFromCy", "ord:124"
    c_ole.Add "VarBoolFromStr", "ord:125"
    c_ole.Add "VarBoolFromDisp", "ord:126"
    c_ole.Add "VarFormatCurrency", "ord:127"
    c_ole.Add "VarWeekdayName", "ord:128"
    c_ole.Add "VarMonthName", "ord:129"
    c_ole.Add "VarUI1FromI2", "ord:130"
    c_ole.Add "VarUI1FromI4", "ord:131"
    c_ole.Add "VarUI1FromR4", "ord:132"
    c_ole.Add "VarUI1FromR8", "ord:133"
    c_ole.Add "VarUI1FromCy", "ord:134"
    c_ole.Add "VarUI1FromDate", "ord:135"
    c_ole.Add "VarUI1FromStr", "ord:136"
    c_ole.Add "VarUI1FromDisp", "ord:137"
    c_ole.Add "VarUI1FromBool", "ord:138"
    c_ole.Add "VarFormatFromTokens", "ord:139"
    c_ole.Add "VarTokenizeFormatString", "ord:140"
    c_ole.Add "VarAdd", "ord:141"
    c_ole.Add "VarAnd", "ord:142"
    c_ole.Add "VarDiv", "ord:143"
    c_ole.Add "DllCanUnloadNow", "ord:144"
    c_ole.Add "DllGetClassObject", "ord:145"
    c_ole.Add "DispCallFunc", "ord:146"
    c_ole.Add "VariantChangeTypeEx", "ord:147"
    c_ole.Add "SafeArrayPtrOfIndex", "ord:148"
    c_ole.Add "SysStringByteLen", "ord:149"
    c_ole.Add "SysAllocStringByteLen", "ord:150"
    c_ole.Add "DllRegisterServer", "ord:151"
    c_ole.Add "VarEqv", "ord:152"
    c_ole.Add "VarIdiv", "ord:153"
    c_ole.Add "VarImp", "ord:154"
    c_ole.Add "VarMod", "ord:155"
    c_ole.Add "VarMul", "ord:156"
    c_ole.Add "VarOr", "ord:157"
    c_ole.Add "VarPow", "ord:158"
    c_ole.Add "VarSub", "ord:159"
    c_ole.Add "CreateTypeLib", "ord:160"
    c_ole.Add "LoadTypeLib", "ord:161"
    c_ole.Add "LoadRegTypeLib", "ord:162"
    c_ole.Add "RegisterTypeLib", "ord:163"
    c_ole.Add "QueryPathOfRegTypeLib", "ord:164"
    c_ole.Add "LHashValOfNameSys", "ord:165"
    c_ole.Add "LHashValOfNameSysA", "ord:166"
    c_ole.Add "VarXor", "ord:167"
    c_ole.Add "VarAbs", "ord:168"
    c_ole.Add "VarFix", "ord:169"
    c_ole.Add "OaBuildVersion", "ord:170"
    c_ole.Add "ClearCustData", "ord:171"
    c_ole.Add "VarInt", "ord:172"
    c_ole.Add "VarNeg", "ord:173"
    c_ole.Add "VarNot", "ord:174"
    c_ole.Add "VarRound", "ord:175"
    c_ole.Add "VarCmp", "ord:176"
    c_ole.Add "VarDecAdd", "ord:177"
    c_ole.Add "VarDecDiv", "ord:178"
    c_ole.Add "VarDecMul", "ord:179"
    c_ole.Add "CreateTypeLib2", "ord:180"
    c_ole.Add "VarDecSub", "ord:181"
    c_ole.Add "VarDecAbs", "ord:182"
    c_ole.Add "LoadTypeLibEx", "ord:183"
    c_ole.Add "SystemTimeToVariantTime", "ord:184"
    c_ole.Add "VariantTimeToSystemTime", "ord:185"
    c_ole.Add "UnRegisterTypeLib", "ord:186"
    c_ole.Add "VarDecFix", "ord:187"
    c_ole.Add "VarDecInt", "ord:188"
    c_ole.Add "VarDecNeg", "ord:189"
    c_ole.Add "VarDecFromUI1", "ord:190"
    c_ole.Add "VarDecFromI2", "ord:191"
    c_ole.Add "VarDecFromI4", "ord:192"
    c_ole.Add "VarDecFromR4", "ord:193"
    c_ole.Add "VarDecFromR8", "ord:194"
    c_ole.Add "VarDecFromDate", "ord:195"
    c_ole.Add "VarDecFromCy", "ord:196"
    c_ole.Add "VarDecFromStr", "ord:197"
    c_ole.Add "VarDecFromDisp", "ord:198"
    c_ole.Add "VarDecFromBool", "ord:199"
    c_ole.Add "GetErrorInfo", "ord:200"
    c_ole.Add "SetErrorInfo", "ord:201"
    c_ole.Add "CreateErrorInfo", "ord:202"
    c_ole.Add "VarDecRound", "ord:203"
    c_ole.Add "VarDecCmp", "ord:204"
    c_ole.Add "VarI2FromI1", "ord:205"
    c_ole.Add "VarI2FromUI2", "ord:206"
    c_ole.Add "VarI2FromUI4", "ord:207"
    c_ole.Add "VarI2FromDec", "ord:208"
    c_ole.Add "VarI4FromI1", "ord:209"
    c_ole.Add "VarI4FromUI2", "ord:210"
    c_ole.Add "VarI4FromUI4", "ord:211"
    c_ole.Add "VarI4FromDec", "ord:212"
    c_ole.Add "VarR4FromI1", "ord:213"
    c_ole.Add "VarR4FromUI2", "ord:214"
    c_ole.Add "VarR4FromUI4", "ord:215"
    c_ole.Add "VarR4FromDec", "ord:216"
    c_ole.Add "VarR8FromI1", "ord:217"
    c_ole.Add "VarR8FromUI2", "ord:218"
    c_ole.Add "VarR8FromUI4", "ord:219"
    c_ole.Add "VarR8FromDec", "ord:220"
    c_ole.Add "VarDateFromI1", "ord:221"
    c_ole.Add "VarDateFromUI2", "ord:222"
    c_ole.Add "VarDateFromUI4", "ord:223"
    c_ole.Add "VarDateFromDec", "ord:224"
    c_ole.Add "VarCyFromI1", "ord:225"
    c_ole.Add "VarCyFromUI2", "ord:226"
    c_ole.Add "VarCyFromUI4", "ord:227"
    c_ole.Add "VarCyFromDec", "ord:228"
    c_ole.Add "VarBstrFromI1", "ord:229"
    c_ole.Add "VarBstrFromUI2", "ord:230"
    c_ole.Add "VarBstrFromUI4", "ord:231"
    c_ole.Add "VarBstrFromDec", "ord:232"
    c_ole.Add "VarBoolFromI1", "ord:233"
    c_ole.Add "VarBoolFromUI2", "ord:234"
    c_ole.Add "VarBoolFromUI4", "ord:235"
    c_ole.Add "VarBoolFromDec", "ord:236"
    c_ole.Add "VarUI1FromI1", "ord:237"
    c_ole.Add "VarUI1FromUI2", "ord:238"
    c_ole.Add "VarUI1FromUI4", "ord:239"
    c_ole.Add "VarUI1FromDec", "ord:240"
    c_ole.Add "VarDecFromI1", "ord:241"
    c_ole.Add "VarDecFromUI2", "ord:242"
    c_ole.Add "VarDecFromUI4", "ord:243"
    c_ole.Add "VarI1FromUI1", "ord:244"
    c_ole.Add "VarI1FromI2", "ord:245"
    c_ole.Add "VarI1FromI4", "ord:246"
    c_ole.Add "VarI1FromR4", "ord:247"
    c_ole.Add "VarI1FromR8", "ord:248"
    c_ole.Add "VarI1FromDate", "ord:249"
    c_ole.Add "VarI1FromCy", "ord:250"
    c_ole.Add "VarI1FromStr", "ord:251"
    c_ole.Add "VarI1FromDisp", "ord:252"
    c_ole.Add "VarI1FromBool", "ord:253"
    c_ole.Add "VarI1FromUI2", "ord:254"
    c_ole.Add "VarI1FromUI4", "ord:255"
    c_ole.Add "VarI1FromDec", "ord:256"
    c_ole.Add "VarUI2FromUI1", "ord:257"
    c_ole.Add "VarUI2FromI2", "ord:258"
    c_ole.Add "VarUI2FromI4", "ord:259"
    c_ole.Add "VarUI2FromR4", "ord:260"
    c_ole.Add "VarUI2FromR8", "ord:261"
    c_ole.Add "VarUI2FromDate", "ord:262"
    c_ole.Add "VarUI2FromCy", "ord:263"
    c_ole.Add "VarUI2FromStr", "ord:264"
    c_ole.Add "VarUI2FromDisp", "ord:265"
    c_ole.Add "VarUI2FromBool", "ord:266"
    c_ole.Add "VarUI2FromI1", "ord:267"
    c_ole.Add "VarUI2FromUI4", "ord:268"
    c_ole.Add "VarUI2FromDec", "ord:269"
    c_ole.Add "VarUI4FromUI1", "ord:270"
    c_ole.Add "VarUI4FromI2", "ord:271"
    c_ole.Add "VarUI4FromI4", "ord:272"
    c_ole.Add "VarUI4FromR4", "ord:273"
    c_ole.Add "VarUI4FromR8", "ord:274"
    c_ole.Add "VarUI4FromDate", "ord:275"
    c_ole.Add "VarUI4FromCy", "ord:276"
    c_ole.Add "VarUI4FromStr", "ord:277"
    c_ole.Add "VarUI4FromDisp", "ord:278"
    c_ole.Add "VarUI4FromBool", "ord:279"
    c_ole.Add "VarUI4FromI1", "ord:280"
    c_ole.Add "VarUI4FromUI2", "ord:281"
    c_ole.Add "VarUI4FromDec", "ord:282"
    c_ole.Add "BSTR_UserSize", "ord:283"
    c_ole.Add "BSTR_UserMarshal", "ord:284"
    c_ole.Add "BSTR_UserUnmarshal", "ord:285"
    c_ole.Add "BSTR_UserFree", "ord:286"
    c_ole.Add "VARIANT_UserSize", "ord:287"
    c_ole.Add "VARIANT_UserMarshal", "ord:288"
    c_ole.Add "VARIANT_UserUnmarshal", "ord:289"
    c_ole.Add "VARIANT_UserFree", "ord:290"
    c_ole.Add "LPSAFEARRAY_UserSize", "ord:291"
    c_ole.Add "LPSAFEARRAY_UserMarshal", "ord:292"
    c_ole.Add "LPSAFEARRAY_UserUnmarshal", "ord:293"
    c_ole.Add "LPSAFEARRAY_UserFree", "ord:294"
    c_ole.Add "LPSAFEARRAY_Size", "ord:295"
    c_ole.Add "LPSAFEARRAY_Marshal", "ord:296"
    c_ole.Add "LPSAFEARRAY_Unmarshal", "ord:297"
    c_ole.Add "VarDecCmpR8", "ord:298"
    c_ole.Add "VarCyAdd", "ord:299"
    c_ole.Add "DllUnregisterServer", "ord:300"
    c_ole.Add "OACreateTypeLib2", "ord:301"
    c_ole.Add "VarCyMul", "ord:303"
    c_ole.Add "VarCyMulI4", "ord:304"
    c_ole.Add "VarCySub", "ord:305"
    c_ole.Add "VarCyAbs", "ord:306"
    c_ole.Add "VarCyFix", "ord:307"
    c_ole.Add "VarCyInt", "ord:308"
    c_ole.Add "VarCyNeg", "ord:309"
    c_ole.Add "VarCyRound", "ord:310"
    c_ole.Add "VarCyCmp", "ord:311"
    c_ole.Add "VarCyCmpR8", "ord:312"
    c_ole.Add "VarBstrCat", "ord:313"
    c_ole.Add "VarBstrCmp", "ord:314"
    c_ole.Add "VarR8Pow", "ord:315"
    c_ole.Add "VarR4CmpR8", "ord:316"
    c_ole.Add "VarR8Round", "ord:317"
    c_ole.Add "VarCat", "ord:318"
    c_ole.Add "VarDateFromUdateEx", "ord:319"
    c_ole.Add "GetRecordInfoFromGuids", "ord:322"
    c_ole.Add "GetRecordInfoFromTypeInfo", "ord:323"
    c_ole.Add "SetVarConversionLocaleSetting", "ord:325"
    c_ole.Add "GetVarConversionLocaleSetting", "ord:326"
    c_ole.Add "SetOaNoCache", "ord:327"
    c_ole.Add "VarCyMulI8", "ord:329"
    c_ole.Add "VarDateFromUdate", "ord:330"
    c_ole.Add "VarUdateFromDate", "ord:331"
    c_ole.Add "GetAltMonthNames", "ord:332"
    c_ole.Add "VarI8FromUI1", "ord:333"
    c_ole.Add "VarI8FromI2", "ord:334"
    c_ole.Add "VarI8FromR4", "ord:335"
    c_ole.Add "VarI8FromR8", "ord:336"
    c_ole.Add "VarI8FromCy", "ord:337"
    c_ole.Add "VarI8FromDate", "ord:338"
    c_ole.Add "VarI8FromStr", "ord:339"
    c_ole.Add "VarI8FromDisp", "ord:340"
    c_ole.Add "VarI8FromBool", "ord:341"
    c_ole.Add "VarI8FromI1", "ord:342"
    c_ole.Add "VarI8FromUI2", "ord:343"
    c_ole.Add "VarI8FromUI4", "ord:344"
    c_ole.Add "VarI8FromDec", "ord:345"
    c_ole.Add "VarI2FromI8", "ord:346"
    c_ole.Add "VarI2FromUI8", "ord:347"
    c_ole.Add "VarI4FromI8", "ord:348"
    c_ole.Add "VarI4FromUI8", "ord:349"
    c_ole.Add "VarR4FromI8", "ord:360"
    c_ole.Add "VarR4FromUI8", "ord:361"
    c_ole.Add "VarR8FromI8", "ord:362"
    c_ole.Add "VarR8FromUI8", "ord:363"
    c_ole.Add "VarDateFromI8", "ord:364"
    c_ole.Add "VarDateFromUI8", "ord:365"
    c_ole.Add "VarCyFromI8", "ord:366"
    c_ole.Add "VarCyFromUI8", "ord:367"
    c_ole.Add "VarBstrFromI8", "ord:368"
    c_ole.Add "VarBstrFromUI8", "ord:369"
    c_ole.Add "VarBoolFromI8", "ord:370"
    c_ole.Add "VarBoolFromUI8", "ord:371"
    c_ole.Add "VarUI1FromI8", "ord:372"
    c_ole.Add "VarUI1FromUI8", "ord:373"
    c_ole.Add "VarDecFromI8", "ord:374"
    c_ole.Add "VarDecFromUI8", "ord:375"
    c_ole.Add "VarI1FromI8", "ord:376"
    c_ole.Add "VarI1FromUI8", "ord:377"
    c_ole.Add "VarUI2FromI8", "ord:378"
    c_ole.Add "VarUI2FromUI8", "ord:379"
    c_ole.Add "OleLoadPictureEx", "ord:401"
    c_ole.Add "OleLoadPictureFileEx", "ord:402"
    c_ole.Add "SafeArrayCreateVector", "ord:411"
    c_ole.Add "SafeArrayCopyData", "ord:412"
    c_ole.Add "VectorFromBstr", "ord:413"
    c_ole.Add "BstrFromVector", "ord:414"
    c_ole.Add "OleIconToCursor", "ord:415"
    c_ole.Add "OleCreatePropertyFrameIndirect", "ord:416"
    c_ole.Add "OleCreatePropertyFrame", "ord:417"
    c_ole.Add "OleLoadPicture", "ord:418"
    c_ole.Add "OleCreatePictureIndirect", "ord:419"
    c_ole.Add "OleCreateFontIndirect", "ord:420"
    c_ole.Add "OleTranslateColor", "ord:421"
    c_ole.Add "OleLoadPictureFile", "ord:422"
    c_ole.Add "OleSavePictureFile", "ord:423"
    c_ole.Add "OleLoadPicturePath", "ord:424"
    c_ole.Add "VarUI4FromI8", "ord:425"
    c_ole.Add "VarUI4FromUI8", "ord:426"
    c_ole.Add "VarI8FromUI8", "ord:427"
    c_ole.Add "VarUI8FromI8", "ord:428"
    c_ole.Add "VarUI8FromUI1", "ord:429"
    c_ole.Add "VarUI8FromI2", "ord:430"
    c_ole.Add "VarUI8FromR4", "ord:431"
    c_ole.Add "VarUI8FromR8", "ord:432"
    c_ole.Add "VarUI8FromCy", "ord:433"
    c_ole.Add "VarUI8FromDate", "ord:434"
    c_ole.Add "VarUI8FromStr", "ord:435"
    c_ole.Add "VarUI8FromDisp", "ord:436"
    c_ole.Add "VarUI8FromBool", "ord:437"
    c_ole.Add "VarUI8FromI1", "ord:438"
    c_ole.Add "VarUI8FromUI2", "ord:439"
    c_ole.Add "VarUI8FromUI4", "ord:440"
    c_ole.Add "VarUI8FromDec", "ord:441"
    c_ole.Add "RegisterTypeLibForUser", "ord:442"
    c_ole.Add "UnRegisterTypeLibForUser", "ord:443"

End Sub


