VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPEEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com
'
'These pe editor classes are v 0.22
'
'5.4.15 (mods pulled from MAP\ShellExt)
'       fixed bug for section offsets for pe32+
'       now loads .net attributes (version, anycpu, isDotNet)
'
'2.13.18 - everything should be x64 safe now...

Public OptionalHeader As New COptionalHeader
Public Sections As New Collection

'loaded on demand
Private relocs As CRelocations
Private res As CResources
Private cimp As CLoadImports
Private cexp As CLoadExports
Private dbgDir As CDebugDir

'Public CanAddSection As Boolean
Public errMessage As String

Public isLoaded As Boolean
Public MachineType As String
Public is64bit As Boolean
Public is32Bit As Boolean
Public isDotNet As Boolean
Public isDotNetAnyCpu As Boolean
Private m_dotNetVersion As String


Private DOSHEADER As IMAGEDOSHEADER
Private NTHEADER As IMAGE_NT_HEADERS

Public LoadedFile As String

Public Enum eMachineType
    mt_UNKNOWN = &H0
    mt_AM33 = &H1D3
    mt_AMD64 = &H8664
    mt_ARM = &H1C0
    mt_EBC = &HEBC
    mt_I386 = &H14C
    mt_IA64 = &H200
    mt_M32R = &H9041
    mt_MIPS16 = &H266
    mt_MIPSFPU = &H366
    mt_MIPSFPU16 = &H466
    mt_POWERPC = &H1F0
    mt_POWERPCFP = &H1F1
    mt_R4000 = &H166
    mt_SH3 = &H1A2
    mt_SH3DSP = &H1A3
    mt_SH4 = &H1A6
    mt_SH5 = &H1A8
    mt_THUMB = &H1C2
    mt_WCEMIPSV2 = &H169
End Enum

Public OSVersion As String
Public LinkerVersion As String
Public ImageVersion As String
Public SubSysVersion As String

Private hash As CWinHash

Property Get dotNetVersion(Optional full As Boolean = False) As String
    On Error Resume Next
    If full Then
        dotNetVersion = m_dotNetVersion
    Else
        dotNetVersion = Left(m_dotNetVersion, 4)
    End If
End Property

Property Get PEVersionReport() As String
    Dim tmp() As String
    push tmp, "OSVersion: " & OSVersion
    push tmp, "LinkerVersion: " & LinkerVersion
    push tmp, "ImageVersion: " & ImageVersion
    push tmp, "SubSysVersion: " & SubSysVersion
    PEVersionReport = Join(tmp, vbCrLf)
End Property

Private Sub SetVersions64(ih As IMAGE_OPTIONAL_HEADER_64)
    With ih
        LinkerVersion = .MajorLinkerVersion & "." & .MinorLinkerVersion
        OSVersion = .MajorOperatingSystemVersion & "." & .MinorOperatingSystemVersion
        ImageVersion = .MajorImageVersion & "." & .MinorImageVersion
        SubSysVersion = .MajorSubsystemVersion & "." & .MinorSubsystemVersion
    End With
End Sub

'we load it on demand..
Property Get Imports() As CLoadImports
    Dim f As Long
    
    If Not isLoaded Then
        Set Imports = New CLoadImports
        Exit Property
    End If
    
    If Not cimp Is Nothing Then
        Set Imports = cimp
        Exit Property
    End If
    
    Set cimp = New CLoadImports
    
    f = FreeFile
    Open LoadedFile For Binary Access Read As f
    cimp.LoadImports f, OptionalHeader.ddVirtualAddress(Import_Table), Me
    Close f
    
    Set Imports = cimp
    
End Property

'we load it on demand..
Property Get Exports() As CLoadExports
    
    Dim f As Long
    
    If Not isLoaded Then
        Set Exports = New CLoadExports
        Exit Property
    End If
    
    If Not cexp Is Nothing Then
        Set Exports = cexp
        Exit Property
    End If
    
    Set cexp = New CLoadExports
    
    f = FreeFile
    Open LoadedFile For Binary Access Read As f
    cexp.LoadExports f, OptionalHeader.ddVirtualAddress(Export_Table), Me
    Close f
    
    Set Exports = cexp
    
End Property

'we load it on demand..
Property Get Resources() As CResources

    If Not isLoaded Then
        Set Resources = New CResources
        Exit Property
    End If
    
    'If Not res Is Nothing Then
    '    Set Resources = res
    '    Exit Property
    'End If
    
    Set res = New CResources
    
    If isLoaded Then res.loadSelf Me
    Set Resources = res
    
End Property

'we load it on demand..
Property Get Relocations() As CRelocations

    If Not isLoaded Then
        Set Relocations = New CRelocations
        Exit Property
    End If
    
    'If Not relocs Is Nothing Then
    '    Set Relocations = relocs
    '    Exit Property
    'End If
    
    Set relocs = New CRelocations
    
    If isLoaded Then relocs.LoadRelocs OptionalHeader.ddVirtualAddress(Relocation_Table), Me
    Set Relocations = relocs
    
End Property

'we load it on demand..
Property Get DebugDirectory() As CDebugDir

    If Not isLoaded Then
        Set DebugDirectory = New CDebugDir
        Exit Property
    End If
    
    If Not dbgDir Is Nothing Then
        Set DebugDirectory = dbgDir
        Exit Property
    End If
    
    Set dbgDir = New CDebugDir
    
    dbgDir.loadSelf Me
    Set DebugDirectory = dbgDir
    
End Property



Sub ShowOffsetCalculator(Optional modal As Boolean = False)
    frmOffsets.Initilize Me ', modal
End Sub

Property Get CompiledDate() As String

    On Error Resume Next
    Dim base As Date
    Dim compiled As Date
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", NTHEADER.FileHeader.TimeDateStamp, base)
    CompiledDate = "GMT: " & Format(compiled, "ddd mmm d h:nn:ss yyyy")

End Property

Function LoadFile(fPath As String) As Boolean
    On Error GoTo hell
        
        Dim i As Long
        Dim f As Long
        Dim udtSH As IMAGE_SECTION_HEADER
        Dim cli As Long

        m_dotNetVersion = Empty
        errMessage = Empty
        'CanAddSection = False
        isLoaded = False
        MachineType = Empty
        is64bit = False
        is32Bit = False
        isDotNet = False
        isDotNetAnyCpu = False
        OSVersion = Empty
        LinkerVersion = Empty
        ImageVersion = Empty
        SubSysVersion = Empty
    
        Set OptionalHeader = New COptionalHeader
        Set Sections = New Collection
        Set cimp = Nothing
        Set cexp = Nothing
        Set res = Nothing
        Set relocs = Nothing
        Set dbgDir = Nothing
        
        If Not FileExists(fPath) Then Err.Raise 1

        f = FreeFile
        
        Open fPath For Binary Access Read As f
        Get f, , DOSHEADER
        Get f, DOSHEADER.e_lfanew + 1, NTHEADER
        
        If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then Err.Raise 2
        
        MachineType = getMachineType(NTHEADER.FileHeader.Machine) 'sets is64Bit/is32Bit
        OptionalHeader.LoadStruc f, is64bit
                
        LoadSections f, NTHEADER
        
        Close f
        LoadFile = True
        isLoaded = True
        LoadedFile = fPath
        
        cli = OptionalHeader.ddVirtualAddress(eDATA_DIRECTORY.CLI_Header)
        GetDotNetAttributes cli
        
        
Exit Function
hell:
    
    Close f
    
    Select Case Err.Number
        Case 1: errMessage = "File not found: " & fPath
        Case 2: errMessage = "Invalid PE signature: " & NTHEADER.Signature
        Case Else: errMessage = Err.Description
    End Select
    
End Function

Function SaveChanges() As Boolean
   On Error GoTo hell

    errMessage = Empty
    If Len(LoadedFile) = 0 Then Err.Raise 1
    If Not FileExists(LoadedFile) Then Err.Raise 2

    Dim i As Long
    Dim f As Long
    Dim udtSH As IMAGE_SECTION_HEADER
    Dim csh As CSection

    f = FreeFile

    Open LoadedFile For Binary As f
    Put f, , DOSHEADER
    Put f, DOSHEADER.e_lfanew + 1, NTHEADER
    OptionalHeader.WriteStruc f
    
    For i = 1 To NTHEADER.FileHeader.NumberOfSections
        Set csh = Sections.Sections(i)
        udtSH = csh.GetStruc
        Put f, , udtSH
    Next

    SaveChanges = True

Exit Function
hell:

    Close f

    Select Case Err.Number
        Case 1: errMessage = "No file is currently Loaded"
        Case 2: errMessage = "File not found: " & LoadedFile
        Case Else: errMessage = Err.Description
    End Select

End Function



'Sub AddSection(csh As CSection)  'not tested yet
'    Dim cur As Integer'
'
 '   cur = NTHEADER.FileHeader.NumberOfSections
'    cur = cur + 1'
'
'    SectionHeaders.Add csh, "sect:" & cur
'    NTHEADER.FileHeader.NumberOfSections = cur'
'
'End Sub

Private Sub LoadSections(fhand As Long, nth As IMAGE_NT_HEADERS)
    
    Dim i As Long
    Dim udtSH As IMAGE_SECTION_HEADER
    Dim CSect As CSection
    
    Set Sections = New Collection
    
    For i = 1 To nth.FileHeader.NumberOfSections
        Get fhand, , udtSH
        Set CSect = New CSection
        CSect.LoadStruc udtSH
        CSect.index = i
        Sections.add CSect
    Next
    
'    CanAddSection = False
'
'    Dim b() As Byte
'    ReDim b(Len(udtSH))
'    Get fhand, , b() ' buffer where next section would sit
'
'    'if all empty then is room to add new section
'    For i = 0 To UBound(b)
'        If b(i) <> 0 Then Exit Sub
'    Next
'
'    CanAddSection = True
    
End Sub

Function VAtoRVA(ByVal va As ULong64) As ULong64
    Set VAtoRVA = va.subtract(OptionalHeader.ImageBase)
End Function

Function RVAtoVA(ByVal rva As ULong64) As ULong64
    Set RVAtoVA = OptionalHeader.ImageBase.add(rva)
End Function

Function VAtoOffset(ByVal va As ULong64) As Long
    VAtoOffset = RvaToOffset(VAtoRVA(va))
End Function
 
Function RvaToOffset(ByVal rva As Long, Optional ByRef VirtualSectionBase As Long, Optional sectName As String) As Long
        
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim vBase As Long, vSize As Long, vMax As Long
    
    For Each c In Sections
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            rva = rva - vBase
            rva = rva + c.PointerToRawData
            RvaToOffset = rva
            VirtualSectionBase = vBase
            sectName = c.nameSec
            Exit Function
       End If
    Next
       
    
End Function

Function OffsetToVA(ByVal fOffset As Long, Optional sectName As String) As ULong64
    Dim rva As Long
    
    rva = OffsetToRVA(fOffset, sectName)
    
    If rva <> 0 Then
        Set OffsetToVA = OptionalHeader.ImageBase.add(rva)
    Else
        Set OffsetToVA = New ULong64
    End If
    
End Function

Function OffsetToRVA(ByVal fOffset As Long, Optional sectName As String) As Long
    
    If Sections.Count = 0 Then
        errMessage = "No Sections Loaded"
        Exit Function
    End If
    
    Dim c As CSection
    Dim rBase As Long, rSize As Long, rMax As Long
    
    For Each c In Sections
       rBase = c.PointerToRawData
       rSize = c.SizeOfRawData
       rMax = rBase + rSize
       
       If fOffset >= rBase And fOffset < rMax Then 'right section
            OffsetToRVA = c.VirtualAddress + fOffset - rBase
            sectName = c.nameSec
            Exit Function
       End If
    Next

End Function

Function GetSectionFromRVA(rva As Long) As CSection
      
    If Sections.Count = 0 Then
        errMessage = "No Sections Loaded"
        Exit Function
    End If
    
    Dim c As CSection
    Dim vBase As Long, vSize As Long, vMax As Long
    
    For Each c In Sections
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            Set GetSectionFromRVA = c
            Exit Function
       End If
    Next
    
End Function

'this is compatiable with pefile.imphash
'   https://www.fireeye.com/blog/threat-research/2014/01/tracking-malware-import-hashing.html
'   https://github.com/erocarrera/pefile/blob/8d60469de3b70109ac603c68c48fb3e7b84261e8/pefile.py
Function impHash() As String
    
    Dim e As CImport
    Dim dll, y, fx() As String, f, tmp As String
    
    If hash Is Nothing Then Set hash = New CWinHash
    
    For Each e In Me.Imports.Modules
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
            Else
                push fx, dll & "." & LCase(f)
            End If
        Next
        
    Next
    
    tmp = Join(fx, ",")
    impHash = hash.HashString(tmp)
    
End Function

Private Function getMachineType(m As Integer) As String
    
    Select Case m
        Case mt_UNKNOWN: getMachineType = "Unknown"
        Case mt_AM33: getMachineType = "AM33"
        Case mt_AMD64: getMachineType = "64 Bit": is64bit = True
        Case mt_ARM: getMachineType = "ARM"
        Case mt_EBC: getMachineType = "EBC"
        Case mt_I386: getMachineType = "32 Bit": is32Bit = True
        Case mt_IA64: getMachineType = "64 Bit": is64bit = True
        Case mt_M32R: getMachineType = "M32R"
        Case mt_MIPS16: getMachineType = "MIPS16"
        Case mt_MIPSFPU: getMachineType = "MIPSFPU"
        Case mt_MIPSFPU16: getMachineType = "MIPSFPU16"
        Case mt_POWERPC: getMachineType = "POWERPC"
        Case mt_POWERPCFP: getMachineType = "POWERPCFP"
        Case mt_R4000: getMachineType = "R4000"
        Case mt_SH3: getMachineType = "SH3"
        Case mt_SH3DSP: getMachineType = "SH3DSP"
        Case mt_SH4: getMachineType = "SH4"
        Case mt_SH5: getMachineType = "SH5"
        Case mt_THUMB: getMachineType = "THUMB"
        Case mt_WCEMIPSV2: getMachineType = "WCEMIPSV2"
        Case Else: getMachineType = "Unknown"
    End Select
    
    
End Function


Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function GetDotNetAttributes(cli As Long)
    
    'ok we are going to need to load it more fully...
    Dim tmp As String
    Dim foff_cli As Long
    Dim major As Integer '2 bytes
    Dim minor As Integer '2 bytes
    Dim flags As Long    '4 bytes
    Dim fs As Long, f As Long
    Dim i As Long
    
    Dim metaData_size As Long
    Dim metaData_rva As Long
    Dim foff_metaData As Long
    Dim verLen As Long
    Dim b() As Byte
    
    Const COMIMAGE_FLAGS_32BITREQUIRED = 2
    
    On Error GoTo cleanup
    
    If cli = 0 Then Exit Function
      
    foff_cli = RvaToOffset(cli) + 1
    
    If foff_cli > 1 And foff_cli < FileLen(LoadedFile) Then
        
        isDotNet = True
        f = FreeFile
        Open LoadedFile For Binary As f
        
        If is32Bit Then 'is it really 32bit or is it AnyCpu?
            Get f, foff_cli + 16, flags
            If (flags And COMIMAGE_FLAGS_32BITREQUIRED) = COMIMAGE_FLAGS_32BITREQUIRED Then
                'its 32bit only so no its not anycpu..
            Else
                isDotNetAnyCpu = True
            End If
        End If

        'now we try to determine the target framework it was compiled for...
        Get f, foff_cli + 8, metaData_rva
        Get f, foff_cli + 12, metaData_size

        foff_metaData = RvaToOffset(metaData_rva)

        foff_metaData = foff_metaData + 1
        Get f, foff_metaData, fs
        If fs <> &H424A5342 Then GoTo cleanup 'signature check
         
        Get f, foff_metaData + 12, verLen
        
        If verLen > 1 And verLen < 20 Then
            ReDim b(verLen)
            Get f, , b()
            m_dotNetVersion = Replace(StrConv(b, vbUnicode), Chr(0), Empty)
        End If
   
cleanup:
        On Error Resume Next
        Close f
        
    End If
        
    

End Function


'Private Function GetDotNetAttributes(cli As Long)
'
'    'ok we are going to need to load it more fully...
'    Dim tmp As String
'    Dim foff_cli As Long
'    Dim major As Integer '2 bytes
'    Dim minor As Integer '2 bytes
'    Dim flags As Long    '4 bytes
'    Dim fs As Long, f As Long
'    Dim b As Byte
'    Dim i As Long
'
'    Dim metaData_size As Long
'    Dim metaData_rva As Long
'    Dim foff_metaData As Long
'    Dim streamCount As Integer
'    Dim verLen As Long
'
'    Dim streamOffset() As Long
'    Dim streamSize() As Long
'    Dim streamName() As String
'
'    Const COMIMAGE_FLAGS_32BITREQUIRED = 2
'
'    On Error GoTo cleanup
'
'    If cli = 0 Then Exit Function
'
'    foff_cli = RvaToOffset(cli) + 1
'
'    If foff_cli > 1 And foff_cli < FileLen(LoadedFile) Then
'
'        isDotNet = True
'        f = FreeFile
'        Open LoadedFile For Binary As f
'
'        If is32Bit Then 'is it really 32bit or is it AnyCpu?
'            Get f, foff_cli + 16, flags
'            If (flags And COMIMAGE_FLAGS_32BITREQUIRED) = COMIMAGE_FLAGS_32BITREQUIRED Then
'                'its 32bit only so no its not anycpu..
'            Else
'                isDotNetAnyCpu = True
'            End If
'        End If
'
''below is a good start, but not complete...not all tables are always present..fu ms for shitty design
''at least give me an offset table..and type flag per table..why multiplex all into one..needless complexity
''for literally a dozen bytes?
'
''        'now shit gets hairy trying to determine the target
''        'framework it was compiled for...
''        Get f, foff_cli + 8, metaData_rva
''        Get f, foff_cli + 12, metaData_size
''
''        foff_metaData = RvaToOffset(metaData_rva)
''
''        foff_metaData = foff_metaData + 1
''        Get f, foff_metaData, fs 'signature = 424A5342
''        Get f, foff_metaData + 12, verLen
''        Get f, foff_metaData + 14 + verLen + 4, streamCount
''
''        If streamCount > 200 Then Exit Function
''
''        ReDim streamOffset(1 To streamCount)
''        ReDim streamSize(1 To streamCount)
''        ReDim streamName(1 To streamCount)
''        Dim curPointer As Long
''
''        For i = 1 To streamCount
''            Get f, , streamOffset(i)
''            Get f, , streamSize(i)
''            Get f, , b
''            Do
''                If b = 0 Then Exit Do
''                streamName(i) = streamName(i) & Chr(b)
''                Get f, , b
''            Loop
''            curPointer = Seek(f) - 1
''            While curPointer Mod 4 <> 0   'REALLY?! offset alignment rather than
''                Get f, , b                'length prefix or simple null termination
''                curPointer = curPointer + 1
''            Wend
''        Next
''
''        For i = 1 To streamCount
''            If streamName(i) = "#-" Or streamName(i) = "#~" Then Exit For
''        Next
''
''        If i > streamCount Then Exit Function
''
''        'goto start of #~ stream, then advance past tables header
''        Seek f, foff_metaData + streamOffset(i) + 24
''
''        Dim tableCounts(1 To 15) As Long
''        Get f, , tableCounts()
''
''        Dim asmRefOffset As Long
''
''        'we want the assembly ref table..we need to calc its offset
''        'based on number of items in each table..fuck you again..
''        For i = 1 To 13
''            asmRefOffset = asmRefOffset + TableSizeByIndex(i, tableCounts(i))
''        Next
''
''        Seek f, Seek(f) + asmRefOffset 'fs is now start of asmref table..
''
''        Dim aRefObj As asmRefHeader
''        Dim stringsIndex As Long
''
''        For i = 1 To streamCount
''            If streamName(i) = "#Strings" Then
''                stringsIndex = i
''                Exit For
''            End If
''        Next
''
''        Dim asmrefName As String
''
''        For i = 1 To tableCounts(14) 'asmref table count..
''            Get f, , aRefObj
''            asmrefName = GetStringTableItem(f, foff_metaData + streamOffset(stringsIndex), aRefObj.name)
''            If asmrefName = "mscorlib" Then
''                dotNetVersion = aRefObj.major & "." & aRefObj.minor
''                Exit For
''            End If
''        Next
'
'cleanup:
'        On Error Resume Next
'        Close f
'
'    End If
'
'
'
'End Function
'
'Private Function GetStringTableItem(fHand As Long, tableStartOffset As Long, index As Integer) As String
'    Dim startFs As Long
'    Dim tmp As String
'    Dim b As Byte
'
'    startFs = Seek(fHand) 'cache where file point currently is
'
'    Seek fHand, tableStartOffset + index
'    Do
'        Get fHand, , b
'        If b = 0 Then Exit Do
'        tmp = tmp + Chr(b)
'    Loop
'
'    GetStringTableItem = tmp
'
'    Seek fHand, startFs 'restore it..
'
'End Function
'
'Private Function TableSizeByIndex(index As Long, itemCount As Long) As Long
'
'    Dim sizes()
'    'modules = 10 bytes
'    'typeref = 6 bytes
'    'typedef = 14
'    'field=6
'    'method = 14
'    'param = 6
'    'memberref = 6 bytes
'    'customattr = 6 bytes
'    'standalonesig = 2
'    'propertymap = 4
'    'property = 6
'    'methsemantics = 6
'    'assembly = 22
'    'assemblyref = target
'
'    sizes = Array(0, 10, 6, 14, 6, 14, 6, 6, 6, 2, 4, 6, 6, 22)
'
'    If index = 0 Or index > UBound(sizes) Then
'        MsgBox "Error in tablesizebyindex"
'        Exit Function
'    End If
'
'    TableSizeByIndex = sizes(index) * itemCount
'
'
'End Function
'
'
'
