VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   14310
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox text2 
      Height          =   3075
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   5424
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   3075
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   5424
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":52B7
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox Text3 
      Height          =   1995
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   7020
      Width           =   13815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ole As Collection
Dim ws2 As Collection

Private Declare Function GetTickCount Lib "kernel32" () As Long

Function ordLookup(dll, ord)
    
    On Error Resume Next
    Dim name As String
    
    'pelib ordinals are in hex..we need dec
    ord = CLng("&h" & Replace(ord, "@", Empty))
    If Err.Number <> 0 Then Err.Raise 1, "ordLookup", "Could not convert ord to long: " & ord
    
    If dll = "ws2_32" Or dll = "wsock32" Or dll = "oleaut32" Then
        If ole Is Nothing Then ordInit
        If dll = "oleaut32" Then
            name = ole("ord:" & ord)
        Else
            name = ws2("ord:" & ord)
        End If
    End If
    
    If Len(name) = 0 Or Err.Number <> 0 Then
         name = "ord" & ord
    End If
        
     ordLookup = LCase(name)
    
End Function

Private Sub Command1_Click()

    Dim tmp() As String
    Dim y()
    Dim c As New Collection

    ReadFile "C:\Documents and Settings\david\Desktop\ws2.txt", tmp
    
    For Each x In tmp
        x = Split(x, ":")
        push y, "ws2.add """ & Replace(x(1), ",", "") & """, ""ord:" & x(0) & """"
    Next
    
    WriteFile "C:\out_out", Join(y, vbCrLf)
   
    
End Sub

Private Sub Form_Load()


    Dim pe As New CPEEditor
    Dim e As CImport
    Dim fx() As String
    Dim dll
    Dim f
    Dim ordCount As Long
    
    Dim st As Long
    st = GetTickCount
    
    text2 = Replace(text2, vbCrLf, Empty)
    
    'pe.LoadFile "C:\windows\notepad.exe"          '0.031s in vb IDE, python 3.243553s
    'pe.LoadFile "C:\windows\system32\mshtml.dll"  '0.265s in vb IDE, python 33.191025s
    
    For Each e In pe.Imports.Modules
        dll = LCase(e.DllName)
        y = Split(dll, ".", 2)
        If UBound(y) > 0 Then
            If y(1) = "ocx" Or y(1) = "sys" Or y(1) = "dll" Then
                dll = y(0)
            End If
        Else
            dll = dll
        End If
        
        For Each f In e.functions
            'check is first character is @ for ordinal (note numbers in hex)
            If VBA.Left(f, 1) = "@" Then
               push fx, dll & "." & ordLookup(dll, f)
               ordCount = ordCount + 1
            Else
                push fx, dll & "." & LCase(f)
            End If
        Next
        
    Next
            
    et = GetTickCount
    text1.Text = Join(fx, ",")
    
    If text1.Text = text2.Text Then
        MsgBox "matched!"
        Text3 = "OrdCount: " & ordCount & vbCrLf & md5(text1.Text) & vbCrLf & "time = " & ((et - st) / 1000) & "s"
    Else
        Me.Caption = "fail CalcLen Top:" & Len(text1) & "  Static Len Bottom:" & Len(text2)
        
        Dim a() As Byte, b() As Byte, firstDiff
        a() = StrConv(text1.Text, vbFromUnicode)
        b() = StrConv(text2.Text, vbFromUnicode)
        
        For i = 0 To UBound(a)
            If a(i) <> b(i) Then
                firstDiff = i
                Exit For
            End If
        Next
        
        text1.SelStart = firstDiff + 1
        text1.SelLength = 20
        text1.SelBold = True
        text1.SelColor = vbRed
        
        text2.SelStart = firstDiff + 1
        text2.SelLength = 20
        text2.SelBold = True
        text2.SelColor = vbRed
        
        k = Mid(text1.Text, firstDiff + 1, 20)
        y = Mid(text2.Text, firstDiff + 1, 20)
        Text3 = "First difference at: " & firstDiff & vbCrLf & k & vbCrLf & y & vbCrLf & "OrdCount: " & ordCount & _
                vbCrLf & vbCrLf & "time = " & ((et - st) / 1000) & "s"

    End If
    
End Sub


Sub ordInit()
    
    Set ws2 = New Collection
    Set ole = New Collection
    
    ws2.Add "accept", "ord:1"
    ws2.Add "bind", "ord:2"
    ws2.Add "closesocket", "ord:3"
    ws2.Add "connect", "ord:4"
    ws2.Add "getpeername", "ord:5"
    ws2.Add "getsockname", "ord:6"
    ws2.Add "getsockopt", "ord:7"
    ws2.Add "htonl", "ord:8"
    ws2.Add "htons", "ord:9"
    ws2.Add "ioctlsocket", "ord:10"
    ws2.Add "inet_addr", "ord:11"
    ws2.Add "inet_ntoa", "ord:12"
    ws2.Add "listen", "ord:13"
    ws2.Add "ntohl", "ord:14"
    ws2.Add "ntohs", "ord:15"
    ws2.Add "recv", "ord:16"
    ws2.Add "recvfrom", "ord:17"
    ws2.Add "select", "ord:18"
    ws2.Add "send", "ord:19"
    ws2.Add "sendto", "ord:20"
    ws2.Add "setsockopt", "ord:21"
    ws2.Add "shutdown", "ord:22"
    ws2.Add "socket", "ord:23"
    ws2.Add "GetAddrInfoW", "ord:24"
    ws2.Add "GetNameInfoW", "ord:25"
    ws2.Add "WSApSetPostRoutine", "ord:26"
    ws2.Add "FreeAddrInfoW", "ord:27"
    ws2.Add "WPUCompleteOverlappedRequest", "ord:28"
    ws2.Add "WSAAccept", "ord:29"
    ws2.Add "WSAAddressToStringA", "ord:30"
    ws2.Add "WSAAddressToStringW", "ord:31"
    ws2.Add "WSACloseEvent", "ord:32"
    ws2.Add "WSAConnect", "ord:33"
    ws2.Add "WSACreateEvent", "ord:34"
    ws2.Add "WSADuplicateSocketA", "ord:35"
    ws2.Add "WSADuplicateSocketW", "ord:36"
    ws2.Add "WSAEnumNameSpaceProvidersA", "ord:37"
    ws2.Add "WSAEnumNameSpaceProvidersW", "ord:38"
    ws2.Add "WSAEnumNetworkEvents", "ord:39"
    ws2.Add "WSAEnumProtocolsA", "ord:40"
    ws2.Add "WSAEnumProtocolsW", "ord:41"
    ws2.Add "WSAEventSelect", "ord:42"
    ws2.Add "WSAGetOverlappedResult", "ord:43"
    ws2.Add "WSAGetQOSByName", "ord:44"
    ws2.Add "WSAGetServiceClassInfoA", "ord:45"
    ws2.Add "WSAGetServiceClassInfoW", "ord:46"
    ws2.Add "WSAGetServiceClassNameByClassIdA", "ord:47"
    ws2.Add "WSAGetServiceClassNameByClassIdW", "ord:48"
    ws2.Add "WSAHtonl", "ord:49"
    ws2.Add "WSAHtons", "ord:50"
    ws2.Add "gethostbyaddr", "ord:51"
    ws2.Add "gethostbyname", "ord:52"
    ws2.Add "getprotobyname", "ord:53"
    ws2.Add "getprotobynumber", "ord:54"
    ws2.Add "getservbyname", "ord:55"
    ws2.Add "getservbyport", "ord:56"
    ws2.Add "gethostname", "ord:57"
    ws2.Add "WSAInstallServiceClassA", "ord:58"
    ws2.Add "WSAInstallServiceClassW", "ord:59"
    ws2.Add "WSAIoctl", "ord:60"
    ws2.Add "WSAJoinLeaf", "ord:61"
    ws2.Add "WSALookupServiceBeginA", "ord:62"
    ws2.Add "WSALookupServiceBeginW", "ord:63"
    ws2.Add "WSALookupServiceEnd", "ord:64"
    ws2.Add "WSALookupServiceNextA", "ord:65"
    ws2.Add "WSALookupServiceNextW", "ord:66"
    ws2.Add "WSANSPIoctl", "ord:67"
    ws2.Add "WSANtohl", "ord:68"
    ws2.Add "WSANtohs", "ord:69"
    ws2.Add "WSAProviderConfigChange", "ord:70"
    ws2.Add "WSARecv", "ord:71"
    ws2.Add "WSARecvDisconnect", "ord:72"
    ws2.Add "WSARecvFrom", "ord:73"
    ws2.Add "WSARemoveServiceClass", "ord:74"
    ws2.Add "WSAResetEvent", "ord:75"
    ws2.Add "WSASend", "ord:76"
    ws2.Add "WSASendDisconnect", "ord:77"
    ws2.Add "WSASendTo", "ord:78"
    ws2.Add "WSASetEvent", "ord:79"
    ws2.Add "WSASetServiceA", "ord:80"
    ws2.Add "WSASetServiceW", "ord:81"
    ws2.Add "WSASocketA", "ord:82"
    ws2.Add "WSASocketW", "ord:83"
    ws2.Add "WSAStringToAddressA", "ord:84"
    ws2.Add "WSAStringToAddressW", "ord:85"
    ws2.Add "WSAWaitForMultipleEvents", "ord:86"
    ws2.Add "WSCDeinstallProvider", "ord:87"
    ws2.Add "WSCEnableNSProvider", "ord:88"
    ws2.Add "WSCEnumProtocols", "ord:89"
    ws2.Add "WSCGetProviderPath", "ord:90"
    ws2.Add "WSCInstallNameSpace", "ord:91"
    ws2.Add "WSCInstallProvider", "ord:92"
    ws2.Add "WSCUnInstallNameSpace", "ord:93"
    ws2.Add "WSCUpdateProvider", "ord:94"
    ws2.Add "WSCWriteNameSpaceOrder", "ord:95"
    ws2.Add "WSCWriteProviderOrder", "ord:96"
    ws2.Add "freeaddrinfo", "ord:97"
    ws2.Add "getaddrinfo", "ord:98"
    ws2.Add "getnameinfo", "ord:99"
    ws2.Add "WSAAsyncSelect", "ord:101"
    ws2.Add "WSAAsyncGetHostByAddr", "ord:102"
    ws2.Add "WSAAsyncGetHostByName", "ord:103"
    ws2.Add "WSAAsyncGetProtoByNumber", "ord:104"
    ws2.Add "WSAAsyncGetProtoByName", "ord:105"
    ws2.Add "WSAAsyncGetServByPort", "ord:106"
    ws2.Add "WSAAsyncGetServByName", "ord:107"
    ws2.Add "WSACancelAsyncRequest", "ord:108"
    ws2.Add "WSASetBlockingHook", "ord:109"
    ws2.Add "WSAUnhookBlockingHook", "ord:110"
    ws2.Add "WSAGetLastError", "ord:111"
    ws2.Add "WSASetLastError", "ord:112"
    ws2.Add "WSACancelBlockingCall", "ord:113"
    ws2.Add "WSAIsBlocking", "ord:114"
    ws2.Add "WSAStartup", "ord:115"
    ws2.Add "WSACleanup", "ord:116"
    ws2.Add "__WSAFDIsSet", "ord:151"
    ws2.Add "WEP", "ord:500"
    
    ole.Add "SysAllocString", "ord:2"
    ole.Add "SysReAllocString", "ord:3"
    ole.Add "SysAllocStringLen", "ord:4"
    ole.Add "SysReAllocStringLen", "ord:5"
    ole.Add "SysFreeString", "ord:6"
    ole.Add "SysStringLen", "ord:7"
    ole.Add "VariantInit", "ord:8"
    ole.Add "VariantClear", "ord:9"
    ole.Add "VariantCopy", "ord:10"
    ole.Add "VariantCopyInd", "ord:11"
    ole.Add "VariantChangeType", "ord:12"
    ole.Add "VariantTimeToDosDateTime", "ord:13"
    ole.Add "DosDateTimeToVariantTime", "ord:14"
    ole.Add "SafeArrayCreate", "ord:15"
    ole.Add "SafeArrayDestroy", "ord:16"
    ole.Add "SafeArrayGetDim", "ord:17"
    ole.Add "SafeArrayGetElemsize", "ord:18"
    ole.Add "SafeArrayGetUBound", "ord:19"
    ole.Add "SafeArrayGetLBound", "ord:20"
    ole.Add "SafeArrayLock", "ord:21"
    ole.Add "SafeArrayUnlock", "ord:22"
    ole.Add "SafeArrayAccessData", "ord:23"
    ole.Add "SafeArrayUnaccessData", "ord:24"
    ole.Add "SafeArrayGetElement", "ord:25"
    ole.Add "SafeArrayPutElement", "ord:26"
    ole.Add "SafeArrayCopy", "ord:27"
    ole.Add "DispGetParam", "ord:28"
    ole.Add "DispGetIDsOfNames", "ord:29"
    ole.Add "DispInvoke", "ord:30"
    ole.Add "CreateDispTypeInfo", "ord:31"
    ole.Add "CreateStdDispatch", "ord:32"
    ole.Add "RegisterActiveObject", "ord:33"
    ole.Add "RevokeActiveObject", "ord:34"
    ole.Add "GetActiveObject", "ord:35"
    ole.Add "SafeArrayAllocDescriptor", "ord:36"
    ole.Add "SafeArrayAllocData", "ord:37"
    ole.Add "SafeArrayDestroyDescriptor", "ord:38"
    ole.Add "SafeArrayDestroyData", "ord:39"
    ole.Add "SafeArrayRedim", "ord:40"
    ole.Add "SafeArrayAllocDescriptorEx", "ord:41"
    ole.Add "SafeArrayCreateEx", "ord:42"
    ole.Add "SafeArrayCreateVectorEx", "ord:43"
    ole.Add "SafeArraySetRecordInfo", "ord:44"
    ole.Add "SafeArrayGetRecordInfo", "ord:45"
    ole.Add "VarParseNumFromStr", "ord:46"
    ole.Add "VarNumFromParseNum", "ord:47"
    ole.Add "VarI2FromUI1", "ord:48"
    ole.Add "VarI2FromI4", "ord:49"
    ole.Add "VarI2FromR4", "ord:50"
    ole.Add "VarI2FromR8", "ord:51"
    ole.Add "VarI2FromCy", "ord:52"
    ole.Add "VarI2FromDate", "ord:53"
    ole.Add "VarI2FromStr", "ord:54"
    ole.Add "VarI2FromDisp", "ord:55"
    ole.Add "VarI2FromBool", "ord:56"
    ole.Add "SafeArraySetIID", "ord:57"
    ole.Add "VarI4FromUI1", "ord:58"
    ole.Add "VarI4FromI2", "ord:59"
    ole.Add "VarI4FromR4", "ord:60"
    ole.Add "VarI4FromR8", "ord:61"
    ole.Add "VarI4FromCy", "ord:62"
    ole.Add "VarI4FromDate", "ord:63"
    ole.Add "VarI4FromStr", "ord:64"
    ole.Add "VarI4FromDisp", "ord:65"
    ole.Add "VarI4FromBool", "ord:66"
    ole.Add "SafeArrayGetIID", "ord:67"
    ole.Add "VarR4FromUI1", "ord:68"
    ole.Add "VarR4FromI2", "ord:69"
    ole.Add "VarR4FromI4", "ord:70"
    ole.Add "VarR4FromR8", "ord:71"
    ole.Add "VarR4FromCy", "ord:72"
    ole.Add "VarR4FromDate", "ord:73"
    ole.Add "VarR4FromStr", "ord:74"
    ole.Add "VarR4FromDisp", "ord:75"
    ole.Add "VarR4FromBool", "ord:76"
    ole.Add "SafeArrayGetVartype", "ord:77"
    ole.Add "VarR8FromUI1", "ord:78"
    ole.Add "VarR8FromI2", "ord:79"
    ole.Add "VarR8FromI4", "ord:80"
    ole.Add "VarR8FromR4", "ord:81"
    ole.Add "VarR8FromCy", "ord:82"
    ole.Add "VarR8FromDate", "ord:83"
    ole.Add "VarR8FromStr", "ord:84"
    ole.Add "VarR8FromDisp", "ord:85"
    ole.Add "VarR8FromBool", "ord:86"
    ole.Add "VarFormat", "ord:87"
    ole.Add "VarDateFromUI1", "ord:88"
    ole.Add "VarDateFromI2", "ord:89"
    ole.Add "VarDateFromI4", "ord:90"
    ole.Add "VarDateFromR4", "ord:91"
    ole.Add "VarDateFromR8", "ord:92"
    ole.Add "VarDateFromCy", "ord:93"
    ole.Add "VarDateFromStr", "ord:94"
    ole.Add "VarDateFromDisp", "ord:95"
    ole.Add "VarDateFromBool", "ord:96"
    ole.Add "VarFormatDateTime", "ord:97"
    ole.Add "VarCyFromUI1", "ord:98"
    ole.Add "VarCyFromI2", "ord:99"
    ole.Add "VarCyFromI4", "ord:100"
    ole.Add "VarCyFromR4", "ord:101"
    ole.Add "VarCyFromR8", "ord:102"
    ole.Add "VarCyFromDate", "ord:103"
    ole.Add "VarCyFromStr", "ord:104"
    ole.Add "VarCyFromDisp", "ord:105"
    ole.Add "VarCyFromBool", "ord:106"
    ole.Add "VarFormatNumber", "ord:107"
    ole.Add "VarBstrFromUI1", "ord:108"
    ole.Add "VarBstrFromI2", "ord:109"
    ole.Add "VarBstrFromI4", "ord:110"
    ole.Add "VarBstrFromR4", "ord:111"
    ole.Add "VarBstrFromR8", "ord:112"
    ole.Add "VarBstrFromCy", "ord:113"
    ole.Add "VarBstrFromDate", "ord:114"
    ole.Add "VarBstrFromDisp", "ord:115"
    ole.Add "VarBstrFromBool", "ord:116"
    ole.Add "VarFormatPercent", "ord:117"
    ole.Add "VarBoolFromUI1", "ord:118"
    ole.Add "VarBoolFromI2", "ord:119"
    ole.Add "VarBoolFromI4", "ord:120"
    ole.Add "VarBoolFromR4", "ord:121"
    ole.Add "VarBoolFromR8", "ord:122"
    ole.Add "VarBoolFromDate", "ord:123"
    ole.Add "VarBoolFromCy", "ord:124"
    ole.Add "VarBoolFromStr", "ord:125"
    ole.Add "VarBoolFromDisp", "ord:126"
    ole.Add "VarFormatCurrency", "ord:127"
    ole.Add "VarWeekdayName", "ord:128"
    ole.Add "VarMonthName", "ord:129"
    ole.Add "VarUI1FromI2", "ord:130"
    ole.Add "VarUI1FromI4", "ord:131"
    ole.Add "VarUI1FromR4", "ord:132"
    ole.Add "VarUI1FromR8", "ord:133"
    ole.Add "VarUI1FromCy", "ord:134"
    ole.Add "VarUI1FromDate", "ord:135"
    ole.Add "VarUI1FromStr", "ord:136"
    ole.Add "VarUI1FromDisp", "ord:137"
    ole.Add "VarUI1FromBool", "ord:138"
    ole.Add "VarFormatFromTokens", "ord:139"
    ole.Add "VarTokenizeFormatString", "ord:140"
    ole.Add "VarAdd", "ord:141"
    ole.Add "VarAnd", "ord:142"
    ole.Add "VarDiv", "ord:143"
    ole.Add "DllCanUnloadNow", "ord:144"
    ole.Add "DllGetClassObject", "ord:145"
    ole.Add "DispCallFunc", "ord:146"
    ole.Add "VariantChangeTypeEx", "ord:147"
    ole.Add "SafeArrayPtrOfIndex", "ord:148"
    ole.Add "SysStringByteLen", "ord:149"
    ole.Add "SysAllocStringByteLen", "ord:150"
    ole.Add "DllRegisterServer", "ord:151"
    ole.Add "VarEqv", "ord:152"
    ole.Add "VarIdiv", "ord:153"
    ole.Add "VarImp", "ord:154"
    ole.Add "VarMod", "ord:155"
    ole.Add "VarMul", "ord:156"
    ole.Add "VarOr", "ord:157"
    ole.Add "VarPow", "ord:158"
    ole.Add "VarSub", "ord:159"
    ole.Add "CreateTypeLib", "ord:160"
    ole.Add "LoadTypeLib", "ord:161"
    ole.Add "LoadRegTypeLib", "ord:162"
    ole.Add "RegisterTypeLib", "ord:163"
    ole.Add "QueryPathOfRegTypeLib", "ord:164"
    ole.Add "LHashValOfNameSys", "ord:165"
    ole.Add "LHashValOfNameSysA", "ord:166"
    ole.Add "VarXor", "ord:167"
    ole.Add "VarAbs", "ord:168"
    ole.Add "VarFix", "ord:169"
    ole.Add "OaBuildVersion", "ord:170"
    ole.Add "ClearCustData", "ord:171"
    ole.Add "VarInt", "ord:172"
    ole.Add "VarNeg", "ord:173"
    ole.Add "VarNot", "ord:174"
    ole.Add "VarRound", "ord:175"
    ole.Add "VarCmp", "ord:176"
    ole.Add "VarDecAdd", "ord:177"
    ole.Add "VarDecDiv", "ord:178"
    ole.Add "VarDecMul", "ord:179"
    ole.Add "CreateTypeLib2", "ord:180"
    ole.Add "VarDecSub", "ord:181"
    ole.Add "VarDecAbs", "ord:182"
    ole.Add "LoadTypeLibEx", "ord:183"
    ole.Add "SystemTimeToVariantTime", "ord:184"
    ole.Add "VariantTimeToSystemTime", "ord:185"
    ole.Add "UnRegisterTypeLib", "ord:186"
    ole.Add "VarDecFix", "ord:187"
    ole.Add "VarDecInt", "ord:188"
    ole.Add "VarDecNeg", "ord:189"
    ole.Add "VarDecFromUI1", "ord:190"
    ole.Add "VarDecFromI2", "ord:191"
    ole.Add "VarDecFromI4", "ord:192"
    ole.Add "VarDecFromR4", "ord:193"
    ole.Add "VarDecFromR8", "ord:194"
    ole.Add "VarDecFromDate", "ord:195"
    ole.Add "VarDecFromCy", "ord:196"
    ole.Add "VarDecFromStr", "ord:197"
    ole.Add "VarDecFromDisp", "ord:198"
    ole.Add "VarDecFromBool", "ord:199"
    ole.Add "GetErrorInfo", "ord:200"
    ole.Add "SetErrorInfo", "ord:201"
    ole.Add "CreateErrorInfo", "ord:202"
    ole.Add "VarDecRound", "ord:203"
    ole.Add "VarDecCmp", "ord:204"
    ole.Add "VarI2FromI1", "ord:205"
    ole.Add "VarI2FromUI2", "ord:206"
    ole.Add "VarI2FromUI4", "ord:207"
    ole.Add "VarI2FromDec", "ord:208"
    ole.Add "VarI4FromI1", "ord:209"
    ole.Add "VarI4FromUI2", "ord:210"
    ole.Add "VarI4FromUI4", "ord:211"
    ole.Add "VarI4FromDec", "ord:212"
    ole.Add "VarR4FromI1", "ord:213"
    ole.Add "VarR4FromUI2", "ord:214"
    ole.Add "VarR4FromUI4", "ord:215"
    ole.Add "VarR4FromDec", "ord:216"
    ole.Add "VarR8FromI1", "ord:217"
    ole.Add "VarR8FromUI2", "ord:218"
    ole.Add "VarR8FromUI4", "ord:219"
    ole.Add "VarR8FromDec", "ord:220"
    ole.Add "VarDateFromI1", "ord:221"
    ole.Add "VarDateFromUI2", "ord:222"
    ole.Add "VarDateFromUI4", "ord:223"
    ole.Add "VarDateFromDec", "ord:224"
    ole.Add "VarCyFromI1", "ord:225"
    ole.Add "VarCyFromUI2", "ord:226"
    ole.Add "VarCyFromUI4", "ord:227"
    ole.Add "VarCyFromDec", "ord:228"
    ole.Add "VarBstrFromI1", "ord:229"
    ole.Add "VarBstrFromUI2", "ord:230"
    ole.Add "VarBstrFromUI4", "ord:231"
    ole.Add "VarBstrFromDec", "ord:232"
    ole.Add "VarBoolFromI1", "ord:233"
    ole.Add "VarBoolFromUI2", "ord:234"
    ole.Add "VarBoolFromUI4", "ord:235"
    ole.Add "VarBoolFromDec", "ord:236"
    ole.Add "VarUI1FromI1", "ord:237"
    ole.Add "VarUI1FromUI2", "ord:238"
    ole.Add "VarUI1FromUI4", "ord:239"
    ole.Add "VarUI1FromDec", "ord:240"
    ole.Add "VarDecFromI1", "ord:241"
    ole.Add "VarDecFromUI2", "ord:242"
    ole.Add "VarDecFromUI4", "ord:243"
    ole.Add "VarI1FromUI1", "ord:244"
    ole.Add "VarI1FromI2", "ord:245"
    ole.Add "VarI1FromI4", "ord:246"
    ole.Add "VarI1FromR4", "ord:247"
    ole.Add "VarI1FromR8", "ord:248"
    ole.Add "VarI1FromDate", "ord:249"
    ole.Add "VarI1FromCy", "ord:250"
    ole.Add "VarI1FromStr", "ord:251"
    ole.Add "VarI1FromDisp", "ord:252"
    ole.Add "VarI1FromBool", "ord:253"
    ole.Add "VarI1FromUI2", "ord:254"
    ole.Add "VarI1FromUI4", "ord:255"
    ole.Add "VarI1FromDec", "ord:256"
    ole.Add "VarUI2FromUI1", "ord:257"
    ole.Add "VarUI2FromI2", "ord:258"
    ole.Add "VarUI2FromI4", "ord:259"
    ole.Add "VarUI2FromR4", "ord:260"
    ole.Add "VarUI2FromR8", "ord:261"
    ole.Add "VarUI2FromDate", "ord:262"
    ole.Add "VarUI2FromCy", "ord:263"
    ole.Add "VarUI2FromStr", "ord:264"
    ole.Add "VarUI2FromDisp", "ord:265"
    ole.Add "VarUI2FromBool", "ord:266"
    ole.Add "VarUI2FromI1", "ord:267"
    ole.Add "VarUI2FromUI4", "ord:268"
    ole.Add "VarUI2FromDec", "ord:269"
    ole.Add "VarUI4FromUI1", "ord:270"
    ole.Add "VarUI4FromI2", "ord:271"
    ole.Add "VarUI4FromI4", "ord:272"
    ole.Add "VarUI4FromR4", "ord:273"
    ole.Add "VarUI4FromR8", "ord:274"
    ole.Add "VarUI4FromDate", "ord:275"
    ole.Add "VarUI4FromCy", "ord:276"
    ole.Add "VarUI4FromStr", "ord:277"
    ole.Add "VarUI4FromDisp", "ord:278"
    ole.Add "VarUI4FromBool", "ord:279"
    ole.Add "VarUI4FromI1", "ord:280"
    ole.Add "VarUI4FromUI2", "ord:281"
    ole.Add "VarUI4FromDec", "ord:282"
    ole.Add "BSTR_UserSize", "ord:283"
    ole.Add "BSTR_UserMarshal", "ord:284"
    ole.Add "BSTR_UserUnmarshal", "ord:285"
    ole.Add "BSTR_UserFree", "ord:286"
    ole.Add "VARIANT_UserSize", "ord:287"
    ole.Add "VARIANT_UserMarshal", "ord:288"
    ole.Add "VARIANT_UserUnmarshal", "ord:289"
    ole.Add "VARIANT_UserFree", "ord:290"
    ole.Add "LPSAFEARRAY_UserSize", "ord:291"
    ole.Add "LPSAFEARRAY_UserMarshal", "ord:292"
    ole.Add "LPSAFEARRAY_UserUnmarshal", "ord:293"
    ole.Add "LPSAFEARRAY_UserFree", "ord:294"
    ole.Add "LPSAFEARRAY_Size", "ord:295"
    ole.Add "LPSAFEARRAY_Marshal", "ord:296"
    ole.Add "LPSAFEARRAY_Unmarshal", "ord:297"
    ole.Add "VarDecCmpR8", "ord:298"
    ole.Add "VarCyAdd", "ord:299"
    ole.Add "DllUnregisterServer", "ord:300"
    ole.Add "OACreateTypeLib2", "ord:301"
    ole.Add "VarCyMul", "ord:303"
    ole.Add "VarCyMulI4", "ord:304"
    ole.Add "VarCySub", "ord:305"
    ole.Add "VarCyAbs", "ord:306"
    ole.Add "VarCyFix", "ord:307"
    ole.Add "VarCyInt", "ord:308"
    ole.Add "VarCyNeg", "ord:309"
    ole.Add "VarCyRound", "ord:310"
    ole.Add "VarCyCmp", "ord:311"
    ole.Add "VarCyCmpR8", "ord:312"
    ole.Add "VarBstrCat", "ord:313"
    ole.Add "VarBstrCmp", "ord:314"
    ole.Add "VarR8Pow", "ord:315"
    ole.Add "VarR4CmpR8", "ord:316"
    ole.Add "VarR8Round", "ord:317"
    ole.Add "VarCat", "ord:318"
    ole.Add "VarDateFromUdateEx", "ord:319"
    ole.Add "GetRecordInfoFromGuids", "ord:322"
    ole.Add "GetRecordInfoFromTypeInfo", "ord:323"
    ole.Add "SetVarConversionLocaleSetting", "ord:325"
    ole.Add "GetVarConversionLocaleSetting", "ord:326"
    ole.Add "SetOaNoCache", "ord:327"
    ole.Add "VarCyMulI8", "ord:329"
    ole.Add "VarDateFromUdate", "ord:330"
    ole.Add "VarUdateFromDate", "ord:331"
    ole.Add "GetAltMonthNames", "ord:332"
    ole.Add "VarI8FromUI1", "ord:333"
    ole.Add "VarI8FromI2", "ord:334"
    ole.Add "VarI8FromR4", "ord:335"
    ole.Add "VarI8FromR8", "ord:336"
    ole.Add "VarI8FromCy", "ord:337"
    ole.Add "VarI8FromDate", "ord:338"
    ole.Add "VarI8FromStr", "ord:339"
    ole.Add "VarI8FromDisp", "ord:340"
    ole.Add "VarI8FromBool", "ord:341"
    ole.Add "VarI8FromI1", "ord:342"
    ole.Add "VarI8FromUI2", "ord:343"
    ole.Add "VarI8FromUI4", "ord:344"
    ole.Add "VarI8FromDec", "ord:345"
    ole.Add "VarI2FromI8", "ord:346"
    ole.Add "VarI2FromUI8", "ord:347"
    ole.Add "VarI4FromI8", "ord:348"
    ole.Add "VarI4FromUI8", "ord:349"
    ole.Add "VarR4FromI8", "ord:360"
    ole.Add "VarR4FromUI8", "ord:361"
    ole.Add "VarR8FromI8", "ord:362"
    ole.Add "VarR8FromUI8", "ord:363"
    ole.Add "VarDateFromI8", "ord:364"
    ole.Add "VarDateFromUI8", "ord:365"
    ole.Add "VarCyFromI8", "ord:366"
    ole.Add "VarCyFromUI8", "ord:367"
    ole.Add "VarBstrFromI8", "ord:368"
    ole.Add "VarBstrFromUI8", "ord:369"
    ole.Add "VarBoolFromI8", "ord:370"
    ole.Add "VarBoolFromUI8", "ord:371"
    ole.Add "VarUI1FromI8", "ord:372"
    ole.Add "VarUI1FromUI8", "ord:373"
    ole.Add "VarDecFromI8", "ord:374"
    ole.Add "VarDecFromUI8", "ord:375"
    ole.Add "VarI1FromI8", "ord:376"
    ole.Add "VarI1FromUI8", "ord:377"
    ole.Add "VarUI2FromI8", "ord:378"
    ole.Add "VarUI2FromUI8", "ord:379"
    ole.Add "OleLoadPictureEx", "ord:401"
    ole.Add "OleLoadPictureFileEx", "ord:402"
    ole.Add "SafeArrayCreateVector", "ord:411"
    ole.Add "SafeArrayCopyData", "ord:412"
    ole.Add "VectorFromBstr", "ord:413"
    ole.Add "BstrFromVector", "ord:414"
    ole.Add "OleIconToCursor", "ord:415"
    ole.Add "OleCreatePropertyFrameIndirect", "ord:416"
    ole.Add "OleCreatePropertyFrame", "ord:417"
    ole.Add "OleLoadPicture", "ord:418"
    ole.Add "OleCreatePictureIndirect", "ord:419"
    ole.Add "OleCreateFontIndirect", "ord:420"
    ole.Add "OleTranslateColor", "ord:421"
    ole.Add "OleLoadPictureFile", "ord:422"
    ole.Add "OleSavePictureFile", "ord:423"
    ole.Add "OleLoadPicturePath", "ord:424"
    ole.Add "VarUI4FromI8", "ord:425"
    ole.Add "VarUI4FromUI8", "ord:426"
    ole.Add "VarI8FromUI8", "ord:427"
    ole.Add "VarUI8FromI8", "ord:428"
    ole.Add "VarUI8FromUI1", "ord:429"
    ole.Add "VarUI8FromI2", "ord:430"
    ole.Add "VarUI8FromR4", "ord:431"
    ole.Add "VarUI8FromR8", "ord:432"
    ole.Add "VarUI8FromCy", "ord:433"
    ole.Add "VarUI8FromDate", "ord:434"
    ole.Add "VarUI8FromStr", "ord:435"
    ole.Add "VarUI8FromDisp", "ord:436"
    ole.Add "VarUI8FromBool", "ord:437"
    ole.Add "VarUI8FromI1", "ord:438"
    ole.Add "VarUI8FromUI2", "ord:439"
    ole.Add "VarUI8FromUI4", "ord:440"
    ole.Add "VarUI8FromDec", "ord:441"
    ole.Add "RegisterTypeLibForUser", "ord:442"
    ole.Add "UnRegisterTypeLibForUser", "ord:443"

End Sub
