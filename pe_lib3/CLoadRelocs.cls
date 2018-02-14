VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRelocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'2.2.2010 David Zimmer <dzzie@yahoo.com>

Public Tables As New Collection  'of CRelocTable

Private mParent As CPEEditor

Private Type RelocTable
    va As Long
    size As Long
End Type

Friend Sub LoadRelocs(ByVal rvaRelocDirectoryTable As Long, parent As CPEEditor)
    
    Dim relocExportTable As Long
    Dim rt As RelocTable
    Dim i As Long, cur As Long
    Dim offset As Long
    Dim RType As Long, rvalue As Long
    Dim t As CRelocTable
    Dim fhandle As Long
    
    Set mParent = parent
    
    If rvaRelocDirectoryTable = 0 Then Exit Sub
    
    relocExportTable = mParent.RvaToOffset(rvaRelocDirectoryTable)
    If relocExportTable = 0 Then Exit Sub
    relocExportTable = relocExportTable + 1 '1 based file offset
    
    fhandle = FreeFile
    Open parent.LoadedFile For Binary Access Read As fhandle
    Get fhandle, relocExportTable, rt
    
    While rt.size > 0
        Set t = New CRelocTable
        t.LoadTable fhandle, rvaRelocDirectoryTable + offset, mParent
        Tables.add t
        offset = offset + rt.size
        Get fhandle, relocExportTable + offset, rt
    Wend
    
    Close fhandle
    

End Sub

