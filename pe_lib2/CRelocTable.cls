VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRelocTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'2.2.2010 David Zimmer <dzzie@yahoo.com>

Public Entries As New Collection  'of CReloc
Public TableRVA As Long  'need tables colleciton for class of CRelocTable?
Public TableSize As Long
Public RelocBase As Long
Public NumEntries As Long

Private mParent As CPEEditor

Private Type RelocTable
    va As Long
    Size As Long
End Type

Private Type entry
    value As Integer
End Type
    

Friend Sub LoadTable(fhandle As Long, ByVal rvaRelocDirectoryTable As Long, parent As CPEEditor)
    
    Dim relocExportTable As Long
    Dim rt As RelocTable
    Dim i As Long, cur As Long
    Dim offset As Long
    Dim e As entry
    Dim RType As Long, rvalue As Long
    Dim r As CReloc
    
    Set mParent = parent
    
    If rvaRelocDirectoryTable < 1 Then Exit Sub
    
    relocExportTable = mParent.RvaToOffset(rvaRelocDirectoryTable) + 1
    offset = relocExportTable
    
    Get fhandle, relocExportTable, rt
    
    Me.TableRVA = relocExportTable - 1
    Me.TableSize = rt.Size
    Me.RelocBase = rt.va
    Me.NumEntries = ((rt.Size - 8) / 2) - 1 ' - (sizeof(struc_header) / sizeof(entry)) -1 (for 0based)
    If Me.NumEntries < 1 Then
        Me.NumEntries = 0
        Exit Sub
    End If
    
    'Debug.Print Hex(rt.va) & " "; Hex(rt.Size)

    For i = 0 To Me.NumEntries
        Set r = New CReloc
        Get fhandle, , e
        RType = ShrX(CLng(e.value), 12)
        rvalue = e.value And &HFFF
        With r
            .RType = RType
            .rva = rt.va + rvalue
            .CurBase = mParent.OptionalHeader.ImageBase
            .index = i
        End With
        Entries.Add r
        
    Next
        
    For Each r In Entries
        Get fhandle, mParent.RvaToOffset(r.rva) + 1, cur
        r.CurValue = cur
        'Debug.Print r.GetInfo(&H1000)
    Next


End Sub




