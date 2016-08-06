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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

'all of these are x64 compatiable..
Public EntryPoint As Double
Public ImageBase As Double
Public SizeOfStackReserve As Double                            'changed
Public SizeOfStackCommit As Double                            'changed
Public SizeOfHeapReserve As Double                            'changed
Public SizeOfHeapCommit As Double                            'changed

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

Friend Function LoadStruc(fHandle As Long, is64Bit As Boolean)
    Dim i As Integer
    
    is64 = is64Bit
    
    'cant use a variant without making the structs public types so just paste..
    
    If is64Bit Then
        Get fHandle, , opt64
        
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
        End With
        
    Else
    
        Get fHandle, , opt
        
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
        End With
        
    End If
       
End Function

Friend Sub WriteStruc(fHandle As Long)
    Dim i As Integer

    If is64 Then

        With opt64
            .Magic = Magic
            .MajorLinkerVersion = MajorLinkerVersion
            .MinorLinkerVersion = MinorLinkerVersion
            .SizeOfCode = SizeOfCode
            .SizeOfInitializedData = SizeOfInitializedData
            .SizeOfUninitializedData = SizeOfUninitializedData
            .AddressOfEntryPoint = EntryPoint
            .BaseOfCode = BaseOfCode
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
    
        Put fHandle, , opt64
        
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
    
        Put fHandle, , opt
        
    End If

End Sub

Property Get ddVirtualAddress(i As eDATA_DIRECTORY) As Long
    ddVirtualAddress = DataDirectory(i).VirtualAddress
End Property

Property Get ddSize(i As eDATA_DIRECTORY) As Long
    ddSize = DataDirectory(i).Size
End Property


