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



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
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
    timeStampToDate = "GMT: " & Format(compiled, "ddd mmm d h:nn:ss yyyy")

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
        lv.ColumnHeaders.add , , "Section Name"
        lv.ColumnHeaders.add , , "VirtualAddr"
        lv.ColumnHeaders.add , , "VirtualSize"
        lv.ColumnHeaders.add , , "RawOffset"
        lv.ColumnHeaders.add , , "RawSize"
        lv.ColumnHeaders.add , , "Characteristics"
        
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
        Set li = lv.ListItems.add(, , cs.nameSec)
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
    
    c_ws2.add "accept", "ord:1"
    c_ws2.add "bind", "ord:2"
    c_ws2.add "closesocket", "ord:3"
    c_ws2.add "connect", "ord:4"
    c_ws2.add "getpeername", "ord:5"
    c_ws2.add "getsockname", "ord:6"
    c_ws2.add "getsockopt", "ord:7"
    c_ws2.add "htonl", "ord:8"
    c_ws2.add "htons", "ord:9"
    c_ws2.add "ioctlsocket", "ord:10"
    c_ws2.add "inet_addr", "ord:11"
    c_ws2.add "inet_ntoa", "ord:12"
    c_ws2.add "listen", "ord:13"
    c_ws2.add "ntohl", "ord:14"
    c_ws2.add "ntohs", "ord:15"
    c_ws2.add "recv", "ord:16"
    c_ws2.add "recvfrom", "ord:17"
    c_ws2.add "select", "ord:18"
    c_ws2.add "send", "ord:19"
    c_ws2.add "sendto", "ord:20"
    c_ws2.add "setsockopt", "ord:21"
    c_ws2.add "shutdown", "ord:22"
    c_ws2.add "socket", "ord:23"
    c_ws2.add "GetAddrInfoW", "ord:24"
    c_ws2.add "GetNameInfoW", "ord:25"
    c_ws2.add "WSApSetPostRoutine", "ord:26"
    c_ws2.add "FreeAddrInfoW", "ord:27"
    c_ws2.add "WPUCompleteOverlappedRequest", "ord:28"
    c_ws2.add "WSAAccept", "ord:29"
    c_ws2.add "WSAAddressToStringA", "ord:30"
    c_ws2.add "WSAAddressToStringW", "ord:31"
    c_ws2.add "WSACloseEvent", "ord:32"
    c_ws2.add "WSAConnect", "ord:33"
    c_ws2.add "WSACreateEvent", "ord:34"
    c_ws2.add "WSADuplicateSocketA", "ord:35"
    c_ws2.add "WSADuplicateSocketW", "ord:36"
    c_ws2.add "WSAEnumNameSpaceProvidersA", "ord:37"
    c_ws2.add "WSAEnumNameSpaceProvidersW", "ord:38"
    c_ws2.add "WSAEnumNetworkEvents", "ord:39"
    c_ws2.add "WSAEnumProtocolsA", "ord:40"
    c_ws2.add "WSAEnumProtocolsW", "ord:41"
    c_ws2.add "WSAEventSelect", "ord:42"
    c_ws2.add "WSAGetOverlappedResult", "ord:43"
    c_ws2.add "WSAGetQOSByName", "ord:44"
    c_ws2.add "WSAGetServiceClassInfoA", "ord:45"
    c_ws2.add "WSAGetServiceClassInfoW", "ord:46"
    c_ws2.add "WSAGetServiceClassNameByClassIdA", "ord:47"
    c_ws2.add "WSAGetServiceClassNameByClassIdW", "ord:48"
    c_ws2.add "WSAHtonl", "ord:49"
    c_ws2.add "WSAHtons", "ord:50"
    c_ws2.add "gethostbyaddr", "ord:51"
    c_ws2.add "gethostbyname", "ord:52"
    c_ws2.add "getprotobyname", "ord:53"
    c_ws2.add "getprotobynumber", "ord:54"
    c_ws2.add "getservbyname", "ord:55"
    c_ws2.add "getservbyport", "ord:56"
    c_ws2.add "gethostname", "ord:57"
    c_ws2.add "WSAInstallServiceClassA", "ord:58"
    c_ws2.add "WSAInstallServiceClassW", "ord:59"
    c_ws2.add "WSAIoctl", "ord:60"
    c_ws2.add "WSAJoinLeaf", "ord:61"
    c_ws2.add "WSALookupServiceBeginA", "ord:62"
    c_ws2.add "WSALookupServiceBeginW", "ord:63"
    c_ws2.add "WSALookupServiceEnd", "ord:64"
    c_ws2.add "WSALookupServiceNextA", "ord:65"
    c_ws2.add "WSALookupServiceNextW", "ord:66"
    c_ws2.add "WSANSPIoctl", "ord:67"
    c_ws2.add "WSANtohl", "ord:68"
    c_ws2.add "WSANtohs", "ord:69"
    c_ws2.add "WSAProviderConfigChange", "ord:70"
    c_ws2.add "WSARecv", "ord:71"
    c_ws2.add "WSARecvDisconnect", "ord:72"
    c_ws2.add "WSARecvFrom", "ord:73"
    c_ws2.add "WSARemoveServiceClass", "ord:74"
    c_ws2.add "WSAResetEvent", "ord:75"
    c_ws2.add "WSASend", "ord:76"
    c_ws2.add "WSASendDisconnect", "ord:77"
    c_ws2.add "WSASendTo", "ord:78"
    c_ws2.add "WSASetEvent", "ord:79"
    c_ws2.add "WSASetServiceA", "ord:80"
    c_ws2.add "WSASetServiceW", "ord:81"
    c_ws2.add "WSASocketA", "ord:82"
    c_ws2.add "WSASocketW", "ord:83"
    c_ws2.add "WSAStringToAddressA", "ord:84"
    c_ws2.add "WSAStringToAddressW", "ord:85"
    c_ws2.add "WSAWaitForMultipleEvents", "ord:86"
    c_ws2.add "WSCDeinstallProvider", "ord:87"
    c_ws2.add "WSCEnableNSProvider", "ord:88"
    c_ws2.add "WSCEnumProtocols", "ord:89"
    c_ws2.add "WSCGetProviderPath", "ord:90"
    c_ws2.add "WSCInstallNameSpace", "ord:91"
    c_ws2.add "WSCInstallProvider", "ord:92"
    c_ws2.add "WSCUnInstallNameSpace", "ord:93"
    c_ws2.add "WSCUpdateProvider", "ord:94"
    c_ws2.add "WSCWriteNameSpaceOrder", "ord:95"
    c_ws2.add "WSCWriteProviderOrder", "ord:96"
    c_ws2.add "freeaddrinfo", "ord:97"
    c_ws2.add "getaddrinfo", "ord:98"
    c_ws2.add "getnameinfo", "ord:99"
    c_ws2.add "WSAAsyncSelect", "ord:101"
    c_ws2.add "WSAAsyncGetHostByAddr", "ord:102"
    c_ws2.add "WSAAsyncGetHostByName", "ord:103"
    c_ws2.add "WSAAsyncGetProtoByNumber", "ord:104"
    c_ws2.add "WSAAsyncGetProtoByName", "ord:105"
    c_ws2.add "WSAAsyncGetServByPort", "ord:106"
    c_ws2.add "WSAAsyncGetServByName", "ord:107"
    c_ws2.add "WSACancelAsyncRequest", "ord:108"
    c_ws2.add "WSASetBlockingHook", "ord:109"
    c_ws2.add "WSAUnhookBlockingHook", "ord:110"
    c_ws2.add "WSAGetLastError", "ord:111"
    c_ws2.add "WSASetLastError", "ord:112"
    c_ws2.add "WSACancelBlockingCall", "ord:113"
    c_ws2.add "WSAIsBlocking", "ord:114"
    c_ws2.add "WSAStartup", "ord:115"
    c_ws2.add "WSACleanup", "ord:116"
    c_ws2.add "__WSAFDIsSet", "ord:151"
    c_ws2.add "WEP", "ord:500"
    
    c_ole.add "SysAllocString", "ord:2"
    c_ole.add "SysReAllocString", "ord:3"
    c_ole.add "SysAllocStringLen", "ord:4"
    c_ole.add "SysReAllocStringLen", "ord:5"
    c_ole.add "SysFreeString", "ord:6"
    c_ole.add "SysStringLen", "ord:7"
    c_ole.add "VariantInit", "ord:8"
    c_ole.add "VariantClear", "ord:9"
    c_ole.add "VariantCopy", "ord:10"
    c_ole.add "VariantCopyInd", "ord:11"
    c_ole.add "VariantChangeType", "ord:12"
    c_ole.add "VariantTimeToDosDateTime", "ord:13"
    c_ole.add "DosDateTimeToVariantTime", "ord:14"
    c_ole.add "SafeArrayCreate", "ord:15"
    c_ole.add "SafeArrayDestroy", "ord:16"
    c_ole.add "SafeArrayGetDim", "ord:17"
    c_ole.add "SafeArrayGetElemsize", "ord:18"
    c_ole.add "SafeArrayGetUBound", "ord:19"
    c_ole.add "SafeArrayGetLBound", "ord:20"
    c_ole.add "SafeArrayLock", "ord:21"
    c_ole.add "SafeArrayUnlock", "ord:22"
    c_ole.add "SafeArrayAccessData", "ord:23"
    c_ole.add "SafeArrayUnaccessData", "ord:24"
    c_ole.add "SafeArrayGetElement", "ord:25"
    c_ole.add "SafeArrayPutElement", "ord:26"
    c_ole.add "SafeArrayCopy", "ord:27"
    c_ole.add "DispGetParam", "ord:28"
    c_ole.add "DispGetIDsOfNames", "ord:29"
    c_ole.add "DispInvoke", "ord:30"
    c_ole.add "CreateDispTypeInfo", "ord:31"
    c_ole.add "CreateStdDispatch", "ord:32"
    c_ole.add "RegisterActiveObject", "ord:33"
    c_ole.add "RevokeActiveObject", "ord:34"
    c_ole.add "GetActiveObject", "ord:35"
    c_ole.add "SafeArrayAllocDescriptor", "ord:36"
    c_ole.add "SafeArrayAllocData", "ord:37"
    c_ole.add "SafeArrayDestroyDescriptor", "ord:38"
    c_ole.add "SafeArrayDestroyData", "ord:39"
    c_ole.add "SafeArrayRedim", "ord:40"
    c_ole.add "SafeArrayAllocDescriptorEx", "ord:41"
    c_ole.add "SafeArrayCreateEx", "ord:42"
    c_ole.add "SafeArrayCreateVectorEx", "ord:43"
    c_ole.add "SafeArraySetRecordInfo", "ord:44"
    c_ole.add "SafeArrayGetRecordInfo", "ord:45"
    c_ole.add "VarParseNumFromStr", "ord:46"
    c_ole.add "VarNumFromParseNum", "ord:47"
    c_ole.add "VarI2FromUI1", "ord:48"
    c_ole.add "VarI2FromI4", "ord:49"
    c_ole.add "VarI2FromR4", "ord:50"
    c_ole.add "VarI2FromR8", "ord:51"
    c_ole.add "VarI2FromCy", "ord:52"
    c_ole.add "VarI2FromDate", "ord:53"
    c_ole.add "VarI2FromStr", "ord:54"
    c_ole.add "VarI2FromDisp", "ord:55"
    c_ole.add "VarI2FromBool", "ord:56"
    c_ole.add "SafeArraySetIID", "ord:57"
    c_ole.add "VarI4FromUI1", "ord:58"
    c_ole.add "VarI4FromI2", "ord:59"
    c_ole.add "VarI4FromR4", "ord:60"
    c_ole.add "VarI4FromR8", "ord:61"
    c_ole.add "VarI4FromCy", "ord:62"
    c_ole.add "VarI4FromDate", "ord:63"
    c_ole.add "VarI4FromStr", "ord:64"
    c_ole.add "VarI4FromDisp", "ord:65"
    c_ole.add "VarI4FromBool", "ord:66"
    c_ole.add "SafeArrayGetIID", "ord:67"
    c_ole.add "VarR4FromUI1", "ord:68"
    c_ole.add "VarR4FromI2", "ord:69"
    c_ole.add "VarR4FromI4", "ord:70"
    c_ole.add "VarR4FromR8", "ord:71"
    c_ole.add "VarR4FromCy", "ord:72"
    c_ole.add "VarR4FromDate", "ord:73"
    c_ole.add "VarR4FromStr", "ord:74"
    c_ole.add "VarR4FromDisp", "ord:75"
    c_ole.add "VarR4FromBool", "ord:76"
    c_ole.add "SafeArrayGetVartype", "ord:77"
    c_ole.add "VarR8FromUI1", "ord:78"
    c_ole.add "VarR8FromI2", "ord:79"
    c_ole.add "VarR8FromI4", "ord:80"
    c_ole.add "VarR8FromR4", "ord:81"
    c_ole.add "VarR8FromCy", "ord:82"
    c_ole.add "VarR8FromDate", "ord:83"
    c_ole.add "VarR8FromStr", "ord:84"
    c_ole.add "VarR8FromDisp", "ord:85"
    c_ole.add "VarR8FromBool", "ord:86"
    c_ole.add "VarFormat", "ord:87"
    c_ole.add "VarDateFromUI1", "ord:88"
    c_ole.add "VarDateFromI2", "ord:89"
    c_ole.add "VarDateFromI4", "ord:90"
    c_ole.add "VarDateFromR4", "ord:91"
    c_ole.add "VarDateFromR8", "ord:92"
    c_ole.add "VarDateFromCy", "ord:93"
    c_ole.add "VarDateFromStr", "ord:94"
    c_ole.add "VarDateFromDisp", "ord:95"
    c_ole.add "VarDateFromBool", "ord:96"
    c_ole.add "VarFormatDateTime", "ord:97"
    c_ole.add "VarCyFromUI1", "ord:98"
    c_ole.add "VarCyFromI2", "ord:99"
    c_ole.add "VarCyFromI4", "ord:100"
    c_ole.add "VarCyFromR4", "ord:101"
    c_ole.add "VarCyFromR8", "ord:102"
    c_ole.add "VarCyFromDate", "ord:103"
    c_ole.add "VarCyFromStr", "ord:104"
    c_ole.add "VarCyFromDisp", "ord:105"
    c_ole.add "VarCyFromBool", "ord:106"
    c_ole.add "VarFormatNumber", "ord:107"
    c_ole.add "VarBstrFromUI1", "ord:108"
    c_ole.add "VarBstrFromI2", "ord:109"
    c_ole.add "VarBstrFromI4", "ord:110"
    c_ole.add "VarBstrFromR4", "ord:111"
    c_ole.add "VarBstrFromR8", "ord:112"
    c_ole.add "VarBstrFromCy", "ord:113"
    c_ole.add "VarBstrFromDate", "ord:114"
    c_ole.add "VarBstrFromDisp", "ord:115"
    c_ole.add "VarBstrFromBool", "ord:116"
    c_ole.add "VarFormatPercent", "ord:117"
    c_ole.add "VarBoolFromUI1", "ord:118"
    c_ole.add "VarBoolFromI2", "ord:119"
    c_ole.add "VarBoolFromI4", "ord:120"
    c_ole.add "VarBoolFromR4", "ord:121"
    c_ole.add "VarBoolFromR8", "ord:122"
    c_ole.add "VarBoolFromDate", "ord:123"
    c_ole.add "VarBoolFromCy", "ord:124"
    c_ole.add "VarBoolFromStr", "ord:125"
    c_ole.add "VarBoolFromDisp", "ord:126"
    c_ole.add "VarFormatCurrency", "ord:127"
    c_ole.add "VarWeekdayName", "ord:128"
    c_ole.add "VarMonthName", "ord:129"
    c_ole.add "VarUI1FromI2", "ord:130"
    c_ole.add "VarUI1FromI4", "ord:131"
    c_ole.add "VarUI1FromR4", "ord:132"
    c_ole.add "VarUI1FromR8", "ord:133"
    c_ole.add "VarUI1FromCy", "ord:134"
    c_ole.add "VarUI1FromDate", "ord:135"
    c_ole.add "VarUI1FromStr", "ord:136"
    c_ole.add "VarUI1FromDisp", "ord:137"
    c_ole.add "VarUI1FromBool", "ord:138"
    c_ole.add "VarFormatFromTokens", "ord:139"
    c_ole.add "VarTokenizeFormatString", "ord:140"
    c_ole.add "VarAdd", "ord:141"
    c_ole.add "VarAnd", "ord:142"
    c_ole.add "VarDiv", "ord:143"
    c_ole.add "DllCanUnloadNow", "ord:144"
    c_ole.add "DllGetClassObject", "ord:145"
    c_ole.add "DispCallFunc", "ord:146"
    c_ole.add "VariantChangeTypeEx", "ord:147"
    c_ole.add "SafeArrayPtrOfIndex", "ord:148"
    c_ole.add "SysStringByteLen", "ord:149"
    c_ole.add "SysAllocStringByteLen", "ord:150"
    c_ole.add "DllRegisterServer", "ord:151"
    c_ole.add "VarEqv", "ord:152"
    c_ole.add "VarIdiv", "ord:153"
    c_ole.add "VarImp", "ord:154"
    c_ole.add "VarMod", "ord:155"
    c_ole.add "VarMul", "ord:156"
    c_ole.add "VarOr", "ord:157"
    c_ole.add "VarPow", "ord:158"
    c_ole.add "VarSub", "ord:159"
    c_ole.add "CreateTypeLib", "ord:160"
    c_ole.add "LoadTypeLib", "ord:161"
    c_ole.add "LoadRegTypeLib", "ord:162"
    c_ole.add "RegisterTypeLib", "ord:163"
    c_ole.add "QueryPathOfRegTypeLib", "ord:164"
    c_ole.add "LHashValOfNameSys", "ord:165"
    c_ole.add "LHashValOfNameSysA", "ord:166"
    c_ole.add "VarXor", "ord:167"
    c_ole.add "VarAbs", "ord:168"
    c_ole.add "VarFix", "ord:169"
    c_ole.add "OaBuildVersion", "ord:170"
    c_ole.add "ClearCustData", "ord:171"
    c_ole.add "VarInt", "ord:172"
    c_ole.add "VarNeg", "ord:173"
    c_ole.add "VarNot", "ord:174"
    c_ole.add "VarRound", "ord:175"
    c_ole.add "VarCmp", "ord:176"
    c_ole.add "VarDecAdd", "ord:177"
    c_ole.add "VarDecDiv", "ord:178"
    c_ole.add "VarDecMul", "ord:179"
    c_ole.add "CreateTypeLib2", "ord:180"
    c_ole.add "VarDecSub", "ord:181"
    c_ole.add "VarDecAbs", "ord:182"
    c_ole.add "LoadTypeLibEx", "ord:183"
    c_ole.add "SystemTimeToVariantTime", "ord:184"
    c_ole.add "VariantTimeToSystemTime", "ord:185"
    c_ole.add "UnRegisterTypeLib", "ord:186"
    c_ole.add "VarDecFix", "ord:187"
    c_ole.add "VarDecInt", "ord:188"
    c_ole.add "VarDecNeg", "ord:189"
    c_ole.add "VarDecFromUI1", "ord:190"
    c_ole.add "VarDecFromI2", "ord:191"
    c_ole.add "VarDecFromI4", "ord:192"
    c_ole.add "VarDecFromR4", "ord:193"
    c_ole.add "VarDecFromR8", "ord:194"
    c_ole.add "VarDecFromDate", "ord:195"
    c_ole.add "VarDecFromCy", "ord:196"
    c_ole.add "VarDecFromStr", "ord:197"
    c_ole.add "VarDecFromDisp", "ord:198"
    c_ole.add "VarDecFromBool", "ord:199"
    c_ole.add "GetErrorInfo", "ord:200"
    c_ole.add "SetErrorInfo", "ord:201"
    c_ole.add "CreateErrorInfo", "ord:202"
    c_ole.add "VarDecRound", "ord:203"
    c_ole.add "VarDecCmp", "ord:204"
    c_ole.add "VarI2FromI1", "ord:205"
    c_ole.add "VarI2FromUI2", "ord:206"
    c_ole.add "VarI2FromUI4", "ord:207"
    c_ole.add "VarI2FromDec", "ord:208"
    c_ole.add "VarI4FromI1", "ord:209"
    c_ole.add "VarI4FromUI2", "ord:210"
    c_ole.add "VarI4FromUI4", "ord:211"
    c_ole.add "VarI4FromDec", "ord:212"
    c_ole.add "VarR4FromI1", "ord:213"
    c_ole.add "VarR4FromUI2", "ord:214"
    c_ole.add "VarR4FromUI4", "ord:215"
    c_ole.add "VarR4FromDec", "ord:216"
    c_ole.add "VarR8FromI1", "ord:217"
    c_ole.add "VarR8FromUI2", "ord:218"
    c_ole.add "VarR8FromUI4", "ord:219"
    c_ole.add "VarR8FromDec", "ord:220"
    c_ole.add "VarDateFromI1", "ord:221"
    c_ole.add "VarDateFromUI2", "ord:222"
    c_ole.add "VarDateFromUI4", "ord:223"
    c_ole.add "VarDateFromDec", "ord:224"
    c_ole.add "VarCyFromI1", "ord:225"
    c_ole.add "VarCyFromUI2", "ord:226"
    c_ole.add "VarCyFromUI4", "ord:227"
    c_ole.add "VarCyFromDec", "ord:228"
    c_ole.add "VarBstrFromI1", "ord:229"
    c_ole.add "VarBstrFromUI2", "ord:230"
    c_ole.add "VarBstrFromUI4", "ord:231"
    c_ole.add "VarBstrFromDec", "ord:232"
    c_ole.add "VarBoolFromI1", "ord:233"
    c_ole.add "VarBoolFromUI2", "ord:234"
    c_ole.add "VarBoolFromUI4", "ord:235"
    c_ole.add "VarBoolFromDec", "ord:236"
    c_ole.add "VarUI1FromI1", "ord:237"
    c_ole.add "VarUI1FromUI2", "ord:238"
    c_ole.add "VarUI1FromUI4", "ord:239"
    c_ole.add "VarUI1FromDec", "ord:240"
    c_ole.add "VarDecFromI1", "ord:241"
    c_ole.add "VarDecFromUI2", "ord:242"
    c_ole.add "VarDecFromUI4", "ord:243"
    c_ole.add "VarI1FromUI1", "ord:244"
    c_ole.add "VarI1FromI2", "ord:245"
    c_ole.add "VarI1FromI4", "ord:246"
    c_ole.add "VarI1FromR4", "ord:247"
    c_ole.add "VarI1FromR8", "ord:248"
    c_ole.add "VarI1FromDate", "ord:249"
    c_ole.add "VarI1FromCy", "ord:250"
    c_ole.add "VarI1FromStr", "ord:251"
    c_ole.add "VarI1FromDisp", "ord:252"
    c_ole.add "VarI1FromBool", "ord:253"
    c_ole.add "VarI1FromUI2", "ord:254"
    c_ole.add "VarI1FromUI4", "ord:255"
    c_ole.add "VarI1FromDec", "ord:256"
    c_ole.add "VarUI2FromUI1", "ord:257"
    c_ole.add "VarUI2FromI2", "ord:258"
    c_ole.add "VarUI2FromI4", "ord:259"
    c_ole.add "VarUI2FromR4", "ord:260"
    c_ole.add "VarUI2FromR8", "ord:261"
    c_ole.add "VarUI2FromDate", "ord:262"
    c_ole.add "VarUI2FromCy", "ord:263"
    c_ole.add "VarUI2FromStr", "ord:264"
    c_ole.add "VarUI2FromDisp", "ord:265"
    c_ole.add "VarUI2FromBool", "ord:266"
    c_ole.add "VarUI2FromI1", "ord:267"
    c_ole.add "VarUI2FromUI4", "ord:268"
    c_ole.add "VarUI2FromDec", "ord:269"
    c_ole.add "VarUI4FromUI1", "ord:270"
    c_ole.add "VarUI4FromI2", "ord:271"
    c_ole.add "VarUI4FromI4", "ord:272"
    c_ole.add "VarUI4FromR4", "ord:273"
    c_ole.add "VarUI4FromR8", "ord:274"
    c_ole.add "VarUI4FromDate", "ord:275"
    c_ole.add "VarUI4FromCy", "ord:276"
    c_ole.add "VarUI4FromStr", "ord:277"
    c_ole.add "VarUI4FromDisp", "ord:278"
    c_ole.add "VarUI4FromBool", "ord:279"
    c_ole.add "VarUI4FromI1", "ord:280"
    c_ole.add "VarUI4FromUI2", "ord:281"
    c_ole.add "VarUI4FromDec", "ord:282"
    c_ole.add "BSTR_UserSize", "ord:283"
    c_ole.add "BSTR_UserMarshal", "ord:284"
    c_ole.add "BSTR_UserUnmarshal", "ord:285"
    c_ole.add "BSTR_UserFree", "ord:286"
    c_ole.add "VARIANT_UserSize", "ord:287"
    c_ole.add "VARIANT_UserMarshal", "ord:288"
    c_ole.add "VARIANT_UserUnmarshal", "ord:289"
    c_ole.add "VARIANT_UserFree", "ord:290"
    c_ole.add "LPSAFEARRAY_UserSize", "ord:291"
    c_ole.add "LPSAFEARRAY_UserMarshal", "ord:292"
    c_ole.add "LPSAFEARRAY_UserUnmarshal", "ord:293"
    c_ole.add "LPSAFEARRAY_UserFree", "ord:294"
    c_ole.add "LPSAFEARRAY_Size", "ord:295"
    c_ole.add "LPSAFEARRAY_Marshal", "ord:296"
    c_ole.add "LPSAFEARRAY_Unmarshal", "ord:297"
    c_ole.add "VarDecCmpR8", "ord:298"
    c_ole.add "VarCyAdd", "ord:299"
    c_ole.add "DllUnregisterServer", "ord:300"
    c_ole.add "OACreateTypeLib2", "ord:301"
    c_ole.add "VarCyMul", "ord:303"
    c_ole.add "VarCyMulI4", "ord:304"
    c_ole.add "VarCySub", "ord:305"
    c_ole.add "VarCyAbs", "ord:306"
    c_ole.add "VarCyFix", "ord:307"
    c_ole.add "VarCyInt", "ord:308"
    c_ole.add "VarCyNeg", "ord:309"
    c_ole.add "VarCyRound", "ord:310"
    c_ole.add "VarCyCmp", "ord:311"
    c_ole.add "VarCyCmpR8", "ord:312"
    c_ole.add "VarBstrCat", "ord:313"
    c_ole.add "VarBstrCmp", "ord:314"
    c_ole.add "VarR8Pow", "ord:315"
    c_ole.add "VarR4CmpR8", "ord:316"
    c_ole.add "VarR8Round", "ord:317"
    c_ole.add "VarCat", "ord:318"
    c_ole.add "VarDateFromUdateEx", "ord:319"
    c_ole.add "GetRecordInfoFromGuids", "ord:322"
    c_ole.add "GetRecordInfoFromTypeInfo", "ord:323"
    c_ole.add "SetVarConversionLocaleSetting", "ord:325"
    c_ole.add "GetVarConversionLocaleSetting", "ord:326"
    c_ole.add "SetOaNoCache", "ord:327"
    c_ole.add "VarCyMulI8", "ord:329"
    c_ole.add "VarDateFromUdate", "ord:330"
    c_ole.add "VarUdateFromDate", "ord:331"
    c_ole.add "GetAltMonthNames", "ord:332"
    c_ole.add "VarI8FromUI1", "ord:333"
    c_ole.add "VarI8FromI2", "ord:334"
    c_ole.add "VarI8FromR4", "ord:335"
    c_ole.add "VarI8FromR8", "ord:336"
    c_ole.add "VarI8FromCy", "ord:337"
    c_ole.add "VarI8FromDate", "ord:338"
    c_ole.add "VarI8FromStr", "ord:339"
    c_ole.add "VarI8FromDisp", "ord:340"
    c_ole.add "VarI8FromBool", "ord:341"
    c_ole.add "VarI8FromI1", "ord:342"
    c_ole.add "VarI8FromUI2", "ord:343"
    c_ole.add "VarI8FromUI4", "ord:344"
    c_ole.add "VarI8FromDec", "ord:345"
    c_ole.add "VarI2FromI8", "ord:346"
    c_ole.add "VarI2FromUI8", "ord:347"
    c_ole.add "VarI4FromI8", "ord:348"
    c_ole.add "VarI4FromUI8", "ord:349"
    c_ole.add "VarR4FromI8", "ord:360"
    c_ole.add "VarR4FromUI8", "ord:361"
    c_ole.add "VarR8FromI8", "ord:362"
    c_ole.add "VarR8FromUI8", "ord:363"
    c_ole.add "VarDateFromI8", "ord:364"
    c_ole.add "VarDateFromUI8", "ord:365"
    c_ole.add "VarCyFromI8", "ord:366"
    c_ole.add "VarCyFromUI8", "ord:367"
    c_ole.add "VarBstrFromI8", "ord:368"
    c_ole.add "VarBstrFromUI8", "ord:369"
    c_ole.add "VarBoolFromI8", "ord:370"
    c_ole.add "VarBoolFromUI8", "ord:371"
    c_ole.add "VarUI1FromI8", "ord:372"
    c_ole.add "VarUI1FromUI8", "ord:373"
    c_ole.add "VarDecFromI8", "ord:374"
    c_ole.add "VarDecFromUI8", "ord:375"
    c_ole.add "VarI1FromI8", "ord:376"
    c_ole.add "VarI1FromUI8", "ord:377"
    c_ole.add "VarUI2FromI8", "ord:378"
    c_ole.add "VarUI2FromUI8", "ord:379"
    c_ole.add "OleLoadPictureEx", "ord:401"
    c_ole.add "OleLoadPictureFileEx", "ord:402"
    c_ole.add "SafeArrayCreateVector", "ord:411"
    c_ole.add "SafeArrayCopyData", "ord:412"
    c_ole.add "VectorFromBstr", "ord:413"
    c_ole.add "BstrFromVector", "ord:414"
    c_ole.add "OleIconToCursor", "ord:415"
    c_ole.add "OleCreatePropertyFrameIndirect", "ord:416"
    c_ole.add "OleCreatePropertyFrame", "ord:417"
    c_ole.add "OleLoadPicture", "ord:418"
    c_ole.add "OleCreatePictureIndirect", "ord:419"
    c_ole.add "OleCreateFontIndirect", "ord:420"
    c_ole.add "OleTranslateColor", "ord:421"
    c_ole.add "OleLoadPictureFile", "ord:422"
    c_ole.add "OleSavePictureFile", "ord:423"
    c_ole.add "OleLoadPicturePath", "ord:424"
    c_ole.add "VarUI4FromI8", "ord:425"
    c_ole.add "VarUI4FromUI8", "ord:426"
    c_ole.add "VarI8FromUI8", "ord:427"
    c_ole.add "VarUI8FromI8", "ord:428"
    c_ole.add "VarUI8FromUI1", "ord:429"
    c_ole.add "VarUI8FromI2", "ord:430"
    c_ole.add "VarUI8FromR4", "ord:431"
    c_ole.add "VarUI8FromR8", "ord:432"
    c_ole.add "VarUI8FromCy", "ord:433"
    c_ole.add "VarUI8FromDate", "ord:434"
    c_ole.add "VarUI8FromStr", "ord:435"
    c_ole.add "VarUI8FromDisp", "ord:436"
    c_ole.add "VarUI8FromBool", "ord:437"
    c_ole.add "VarUI8FromI1", "ord:438"
    c_ole.add "VarUI8FromUI2", "ord:439"
    c_ole.add "VarUI8FromUI4", "ord:440"
    c_ole.add "VarUI8FromDec", "ord:441"
    c_ole.add "RegisterTypeLibForUser", "ord:442"
    c_ole.add "UnRegisterTypeLibForUser", "ord:443"

End Sub


