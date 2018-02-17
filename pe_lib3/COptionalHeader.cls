VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COptionalHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

'all of these are x64 compatiable..
Public EntryPoint As New ULong64
Public ImageBase As New ULong64
Public SizeOfStackReserve As New ULong64                          'changed
Public SizeOfStackCommit As New ULong64                            'changed
Public SizeOfHeapReserve As New ULong64                            'changed
Public SizeOfHeapCommit As New ULong64                           'changed

Public Magic As Integer
Public MajorLinkerVersion As Byte
Public MinorLinkerVersion As Byte
Public SizeOfCode As Long
Public SizeOfInitializedData As Long
Public SizeOfUninitializedData As Long
Public BaseOfCode As Long
Public BaseOfData As Long
Public SectionAlignment As Long
Public FileAlignment As Long
Public MajorOperatingSystemVersion As Integer
Public MinorOperatingSystemVersion As Integer
Public MajorImageVersion As Integer
Public MinorImageVersion As Integer
Public MajorSubsystemVersion As Integer
Public MinorSubsystemVersion As Integer
Public Win32VersionValue As Long
Public SizeOfImage As Long
Public SizeOfHeaders As Long
Public CheckSum As Long
Public Subsystem As Integer
Public DllCharacteristics As Integer
Public LoaderFlags As Long
Public NumberOfRvaAndSizes As Long
Private DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
'Private DataDirectory64(0 To 15) As IMAGE_DATA_DIRECTORY_64

Private is64 As Boolean
Private opt64 As IMAGE_OPTIONAL_HEADER_64
Private opt As IMAGE_OPTIONAL_HEADER

Public Enum eDATA_DIRECTORY
    Export_Table = 0
    Import_Table = 1
    Resource_Table = 2
    Exception_Table = 3
    Certificate_Table = 4
    Relocation_Table = 5
    Debug_Data = 6
    Architecture_Data = 7
    Machine_Value = 8        '(MIPS_GP)
    TLS_Table = 9
    Load_Configuration_Table = 10
    Bound_Import_Table = 11
    Import_Address_Table = 12
    Delay_Import_Descriptor = 13
    CLI_Header = 14
    Reserved = 15
End Enum

Public OSVersion As String
Public LinkerVersion As String
Public ImageVersion As String
Public SubSysVersion As String

Function PEVersionReport(Optional compact As Boolean = False) As String
    Dim tmp() As String
    
    If OSVersion = Empty Then Exit Function 'marker that its not set...
    
    If compact Then
        PEVersionReport = "PEVersion:" & "OS:" & OSVersion & " Link:" & LinkerVersion & _
                                         " Img:" & ImageVersion & " SubSys:" & SubSysVersion
    Else
        push tmp, rpad("OSVersion: ", 16) & OSVersion
        push tmp, rpad("LinkerVersion: ", 16) & LinkerVersion
        push tmp, rpad("ImageVersion: ", 16) & ImageVersion
        push tmp, rpad("SubSysVersion: ", 16) & SubSysVersion
        PEVersionReport = Join(tmp, vbCrLf)
    End If
    
End Function

Friend Function LoadStruc(fhandle As Long, is64bit As Boolean)
    Dim i As Integer
    
    is64 = is64bit
    
    'cant use a variant without making the structs public types so just paste..
    
    If is64bit Then
        Get fhandle, , opt64
        
        With opt64
             Magic = .Magic
             MajorLinkerVersion = .MajorLinkerVersion
             MinorLinkerVersion = .MinorLinkerVersion
             SizeOfCode = .SizeOfCode
             SizeOfInitializedData = .SizeOfInitializedData
             SizeOfUninitializedData = .SizeOfUninitializedData
             EntryPoint = .AddressOfEntryPoint
             BaseOfCode = .BaseOfCode
             BaseOfData = 0
             ImageBase = .ImageBase
             SectionAlignment = .SectionAlignment
             FileAlignment = .FileAlignment
             MajorOperatingSystemVersion = .MajorOperatingSystemVersion
             MinorOperatingSystemVersion = .MinorOperatingSystemVersion
             MajorImageVersion = .MajorImageVersion
             MinorImageVersion = .MinorImageVersion
             MajorSubsystemVersion = .MajorSubsystemVersion
             MinorSubsystemVersion = .MinorSubsystemVersion
             Win32VersionValue = .Win32VersionValue
             SizeOfImage = .SizeOfImage
             SizeOfHeaders = .SizeOfHeaders
             CheckSum = .CheckSum
             Subsystem = .Subsystem
             DllCharacteristics = .DllCharacteristics
             SizeOfStackReserve = .SizeOfStackReserve
             SizeOfStackCommit = .SizeOfStackCommit
             SizeOfHeapReserve = .SizeOfHeapReserve
             SizeOfHeapCommit = .SizeOfHeapCommit
             LoaderFlags = .LoaderFlags
             NumberOfRvaAndSizes = .NumberOfRvaAndSizes
             For i = 0 To 15
                 DataDirectory(i) = .DataDirectory(i)
             Next
        
            LinkerVersion = .MajorLinkerVersion & "." & .MinorLinkerVersion
            OSVersion = .MajorOperatingSystemVersion & "." & .MinorOperatingSystemVersion
            ImageVersion = .MajorImageVersion & "." & .MinorImageVersion
            SubSysVersion = .MajorSubsystemVersion & "." & .MinorSubsystemVersion
        End With
        
    Else
    
        Get fhandle, , opt
        
        With opt
             Magic = .Magic
             MajorLinkerVersion = .MajorLinkerVersion
             MinorLinkerVersion = .MinorLinkerVersion
             SizeOfCode = .SizeOfCode
             SizeOfInitializedData = .SizeOfInitializedData
             SizeOfUninitializedData = .SizeOfUninitializedData
             EntryPoint = .AddressOfEntryPoint
             BaseOfCode = .BaseOfCode
             BaseOfData = .BaseOfData
             ImageBase = .ImageBase
             SectionAlignment = .SectionAlignment
             FileAlignment = .FileAlignment
             MajorOperatingSystemVersion = .MajorOperatingSystemVersion
             MinorOperatingSystemVersion = .MinorOperatingSystemVersion
             MajorImageVersion = .MajorImageVersion
             MinorImageVersion = .MinorImageVersion
             MajorSubsystemVersion = .MajorSubsystemVersion
             MinorSubsystemVersion = .MinorSubsystemVersion
             Win32VersionValue = .Win32VersionValue
             SizeOfImage = .SizeOfImage
             SizeOfHeaders = .SizeOfHeaders
             CheckSum = .CheckSum
             Subsystem = .Subsystem
             DllCharacteristics = .DllCharacteristics
             SizeOfStackReserve = .SizeOfStackReserve
             SizeOfStackCommit = .SizeOfStackCommit
             SizeOfHeapReserve = .SizeOfHeapReserve
             SizeOfHeapCommit = .SizeOfHeapCommit
             LoaderFlags = .LoaderFlags
             NumberOfRvaAndSizes = .NumberOfRvaAndSizes
             For i = 0 To 15
                 DataDirectory(i) = .DataDirectory(i)
             Next
        
            LinkerVersion = .MajorLinkerVersion & "." & .MinorLinkerVersion
            OSVersion = .MajorOperatingSystemVersion & "." & .MinorOperatingSystemVersion
            ImageVersion = .MajorImageVersion & "." & .MinorImageVersion
            SubSysVersion = .MajorSubsystemVersion & "." & .MinorSubsystemVersion
        End With
        
    End If
       
End Function

Friend Sub WriteStruc(fhandle As Long)
    Dim i As Integer

    If is64 Then

        With opt64
            .Magic = Magic
            .MajorLinkerVersion = MajorLinkerVersion
            .MinorLinkerVersion = MinorLinkerVersion
            .SizeOfCode = SizeOfCode
            .SizeOfInitializedData = SizeOfInitializedData
            .SizeOfUninitializedData = SizeOfUninitializedData
            .AddressOfEntryPoint = EntryPoint.rawValue
            .BaseOfCode = BaseOfCode
            .ImageBase = ImageBase.rawValue
            .SectionAlignment = SectionAlignment
            .FileAlignment = FileAlignment
            .MajorOperatingSystemVersion = MajorOperatingSystemVersion
            .MinorOperatingSystemVersion = MinorOperatingSystemVersion
            .MajorImageVersion = MajorImageVersion
            .MinorImageVersion = MinorImageVersion
            .MajorSubsystemVersion = MajorSubsystemVersion
            .MinorSubsystemVersion = MinorSubsystemVersion
            .Win32VersionValue = Win32VersionValue
            .SizeOfImage = SizeOfImage
            .SizeOfHeaders = SizeOfHeaders
            .CheckSum = CheckSum
            .Subsystem = Subsystem
            .DllCharacteristics = DllCharacteristics
            .SizeOfStackReserve = SizeOfStackReserve.rawValue
            .SizeOfStackCommit = SizeOfStackCommit.rawValue
            .SizeOfHeapReserve = SizeOfHeapReserve.rawValue
            .SizeOfHeapCommit = SizeOfHeapCommit.rawValue
            .LoaderFlags = LoaderFlags
            .NumberOfRvaAndSizes = NumberOfRvaAndSizes
             For i = 0 To 15
                .DataDirectory(i) = DataDirectory(i)
             Next
        End With
    
        Put fhandle, , opt64
        
    Else
    
        With opt
            .Magic = Magic
            .MajorLinkerVersion = MajorLinkerVersion
            .MinorLinkerVersion = MinorLinkerVersion
            .SizeOfCode = SizeOfCode
            .SizeOfInitializedData = SizeOfInitializedData
            .SizeOfUninitializedData = SizeOfUninitializedData
            .AddressOfEntryPoint = EntryPoint
            .BaseOfCode = BaseOfCode
            .BaseOfData = BaseOfData
            .ImageBase = ImageBase
            .SectionAlignment = SectionAlignment
            .FileAlignment = FileAlignment
            .MajorOperatingSystemVersion = MajorOperatingSystemVersion
            .MinorOperatingSystemVersion = MinorOperatingSystemVersion
            .MajorImageVersion = MajorImageVersion
            .MinorImageVersion = MinorImageVersion
            .MajorSubsystemVersion = MajorSubsystemVersion
            .MinorSubsystemVersion = MinorSubsystemVersion
            .Win32VersionValue = Win32VersionValue
            .SizeOfImage = SizeOfImage
            .SizeOfHeaders = SizeOfHeaders
            .CheckSum = CheckSum
            .Subsystem = Subsystem
            .DllCharacteristics = DllCharacteristics
            .SizeOfStackReserve = SizeOfStackReserve
            .SizeOfStackCommit = SizeOfStackCommit
            .SizeOfHeapReserve = SizeOfHeapReserve
            .SizeOfHeapCommit = SizeOfHeapCommit
            .LoaderFlags = LoaderFlags
            .NumberOfRvaAndSizes = NumberOfRvaAndSizes
             For i = 0 To 15
                .DataDirectory(i) = DataDirectory(i)
             Next
        End With
    
        Put fhandle, , opt
        
    End If

End Sub

Property Get ddVirtualAddress(i As eDATA_DIRECTORY) As Long
    'Set ddVirtualAddress = New ULong64
    'If is64 Then
    '    ddVirtualAddress = DataDirectory64(i).VirtualAddress
    'Else
        ddVirtualAddress = DataDirectory(i).VirtualAddress
    'End If
End Property

Property Get ddSize(i As eDATA_DIRECTORY) As Long
    'Set ddSize = New ULong64
    'If is64 Then
    '    ddSize = DataDirectory64(i).size
    'Else
        ddSize = DataDirectory(i).size
    'End If
End Property

'2.17.18
Property Get DEP() As Boolean
    DEP = ((Me.DllCharacteristics And &H100) = &H100)
End Property

Property Get ASLR() As Boolean
    ASLR = ((Me.DllCharacteristics And &H40) = &H40)
End Property

Property Get FORCE_INTEGRITY() As Boolean
    FORCE_INTEGRITY = ((Me.DllCharacteristics And &H80) = &H80)
End Property

Property Let DEP(v As Boolean)
    If v Then
        DllCharacteristics = DllCharacteristics Or &H100
    Else
        DllCharacteristics = DllCharacteristics Xor &H100
    End If
End Property

Property Let ASLR(v As Boolean)
    If v Then
        DllCharacteristics = DllCharacteristics Or &H40
    Else
        DllCharacteristics = DllCharacteristics Xor &H40
    End If
End Property

Property Let FORCE_INTEGRITY(v As Boolean)
    If v Then
        DllCharacteristics = DllCharacteristics Or &H80
    Else
        DllCharacteristics = DllCharacteristics Xor &H80
    End If
End Property


