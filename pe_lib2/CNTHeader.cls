VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNTHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Signature As String
Public Machine As Integer
Public NumberOfSections As Integer
Public TimeDateStamp As Long
Public PointerToSymbolTable As Long
Public NumberOfSymbols As Long
Public SizeOfOptionalHeader As Integer
Public Characteristics As Integer

Friend Sub LoadSelf(h As IMAGE_NT_HEADERS)
    
    With Me
        .Signature = h.Signature
        .Machine = h.FileHeader.Machine
        .NumberOfSections = h.FileHeader.NumberOfSections
        .TimeDateStamp = h.FileHeader.TimeDateStamp
        .PointerToSymbolTable = h.FileHeader.PointerToSymbolTable
        .NumberOfSections = h.FileHeader.NumberOfSections
        .SizeOfOptionalHeader = h.FileHeader.SizeOfOptionalHeader
        .Characteristics = h.FileHeader.Characteristics
    End With
    
End Sub
