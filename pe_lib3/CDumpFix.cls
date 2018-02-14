VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDumpFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

'2.13.18 - should be x64 safe now...

Private Type IMAGEDOSHEADER
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

Private Type IMAGE_SECTION_HEADER
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

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
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

Private Type IMAGE_OPTIONAL_HEADER_64
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

Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
    'OptionalHeader As IMAGE_OPTIONAL_HEADER  'can be 32 or 64 bit struct..load on demand after testing..
End Type

Public errMsg As String

Function QuickDumpFix(fPath As String) As Boolean
    
    Dim DOSHEADER As IMAGEDOSHEADER
    Dim NTHEADER As IMAGE_NT_HEADERS
    Dim Sections() As IMAGE_SECTION_HEADER
    Dim udtSH As IMAGE_SECTION_HEADER
    Dim opt As IMAGE_OPTIONAL_HEADER
    Dim opt64 As IMAGE_OPTIONAL_HEADER_64
    
    Dim rvaImportDir As Long
    Dim f As Long
    Dim i As Long
    Dim is64bit As Boolean
    
    Const mt_AMD64 = &H8664
    Const mt_IA64 = &H200
    
    If Not FileExists(fPath) Then
        errMsg = "File not found: " & fPath
        Exit Function
    End If
    
    f = FreeFile
    Open fPath For Binary As f
    Get f, , DOSHEADER
    Get f, DOSHEADER.e_lfanew + 1, NTHEADER
    
    If NTHEADER.FileHeader.Machine = mt_AMD64 Or NTHEADER.FileHeader.Machine = mt_IA64 Then
        is64bit = True
        Get f, , opt64
    Else
        Get f, , opt
    End If
    
    ReDim Sections(NTHEADER.FileHeader.NumberOfSections)
    
    For i = 1 To NTHEADER.FileHeader.NumberOfSections
        Get f, , udtSH
        
        With udtSH
            .PointerToRawData = .VirtualAddress
            .SizeOfRawData = .VirtualSize
        End With
        
        Sections(i) = udtSH
    Next
    
    Put f, DOSHEADER.e_lfanew + 1, NTHEADER
    
    If is64bit Then
        opt64.FileAlignment = &H1000
        Put f, , opt64
    Else
        opt.FileAlignment = &H1000
        Put f, , opt
    End If
    
    For i = 1 To NTHEADER.FileHeader.NumberOfSections
        Put f, , Sections(i)
    Next
    
    Close f
    
    QuickDumpFix = True
    
End Function



Function FileExists(path) As Boolean
 If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

