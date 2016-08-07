VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadRelocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'2.2.2010 David Zimmer <dzzie@yahoo.com>

Public Tables As New Collection  'of CRelocTable

Private mParent As CPEEditor
Private Type RelocTable
    va As Long
    Size As Long
End Type

Friend Sub LoadRelocs(fhandle As Long, ByVal rvaRelocDirectoryTable As Long, parent As CPEEditor)
    
    Dim relocExportTable As Long
    Dim rt As RelocTable
    Dim i As Long, cur As Long
    Dim offset As Long
    Dim RType As Long, rvalue As Long
    Dim t As CRelocTable
    
    Set mParent = parent
    
    If rvaRelocDirectoryTable < 1 Then Exit Sub
    
    relocExportTable = mParent.RvaToOffset(rvaRelocDirectoryTable) + 1
    'offset = relocExportTable
    
    Get fhandle, relocExportTable, rt
    While rt.Size > 0
        Set t = New CRelocTable
        t.LoadTable fhandle, rvaRelocDirectoryTable + offset, mParent
        Tables.Add t
        offset = offset + rt.Size
        Get fhandle, relocExportTable + offset, rt
    Wend
    

End Sub

