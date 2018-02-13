VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadImports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com
'
'2.13.18 - should be x64 safe now...

Public Modules As New Collection  'of CImport
Private fhandle As Long 'my FileHandle
Private mParent As CPEEditor
Private idt As IMAGE_IMPORT_DIRECTORY
    
Friend Sub LoadImports(fhand As Long, rvaImportDir As Long, parent As CPEEditor)
    On Error Resume Next
    
    Dim fOffsetIDT As Long
       
    Set mParent = parent
    If rvaImportDir = 0 Then Exit Sub
    
    fOffsetIDT = mParent.RvaToOffset(rvaImportDir)
    
    If fOffsetIDT < 1 Then
        'MsgBox "Error loading imports getting file offset from rva? rva=" & rvaImportDir & " offset=" & fOffsetIDT, vbExclamation
        Exit Sub
    End If
    
    fhandle = fhand
    fOffsetIDT = fOffsetIDT + 1 'cause 0 based file offset
    Get fhandle, fOffsetIDT, idt
    
    
    While idt.pFuncAry > 0
        LoadModuleImports
        fOffsetIDT = fOffsetIDT + Len(idt)
        'Debug.Print "offset=" & Hex(fOffsetIDT)
        Get fhandle, fOffsetIDT, idt
    Wend
    
End Sub

Private Sub LoadModuleImports()
    'On Error GoTo hell
    
    Dim m  As New CImport

    With idt
       'Debug.Print "pFuncAry=" & Hex(.pFuncAry)
       m.pLookupTable = mParent.RvaToOffset(.pFuncAry)
       m.DllName = ScanString(mParent.RvaToOffset(.pDllName))
       m.FirstThunk = .pThunk
    End With

    Dim ImportLookup As Long, ImportLookup64 As Currency, i As Long
    Dim isOrdial As Boolean, ImportFx As String
    Dim lookup As New ULong64
    Dim address_mask As New ULong64
    Dim ordFlag As New ULong64
    
    address_mask.fromString "0x7fffffffffffffff"
    ordFlag.fromString "0x8000000000000000"
    lookup.useTick = False

    i = m.pLookupTable + 1
    
    If mParent.is64bit Then
        Get fhandle, i, ImportLookup64
        lookup = ImportLookup64
        isOrdial = lookup.and_(ordFlag) <> 0
        lookup = lookup.and_(address_mask)
    Else
        Get fhandle, i, ImportLookup
        isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
        ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)
        lookup = ImportLookup
    End If
        
    While lookup.greaterThan(0) Or isOrdial
    
        If isOrdial Then
            ImportFx = "@" & lookup.toString() 'if we switch to decimal change ordLookup for impHash
            'Debug.Print ImportFx
            m.functions.add ImportFx
        Else
            ImportFx = ScanString(mParent.RvaToOffset(lookup.lo), True) 'is this x64 safe?
            'Debug.Print ImportFx
            m.functions.add ImportFx
        End If
        
        If mParent.is64bit Then
            i = i + 8 'next record is 8 bytes away
            Get fhandle, i, ImportLookup64
            lookup = ImportLookup64
            isOrdial = lookup.and_(ordFlag) <> 0
            lookup = lookup.and_(address_mask)
        Else
            i = i + 4 'next record is 4 bytes away
            Get fhandle, i, ImportLookup
            isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
            ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)
            lookup = ImportLookup
        End If

    Wend
    
    Modules.add m, CStr(m.DllName)

Exit Sub
hell: MsgBox Err.Description

End Sub

Private Function ScanString(offset As Long, Optional getHint As Boolean = False) As String
    Dim b As Byte, i As Long, ret As String
    Dim Hint As Integer '2 bytes
     
    i = offset + 1
    
    If getHint Then
        Get fhandle, i, Hint
        i = i + 2
    End If
    
    Get fhandle, i, b
     
    While b <> 0
        ret = ret & Chr(b)
        Get fhandle, , b
    Wend
    
    ScanString = ret
 
End Function


Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
