VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadExports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'License:   GPL
'Copyright: David Zimmer 2002
'Site:      http://sandsprite.com
'
'Author:    David Zimmer <dzzie@yahoo.com>
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
'
'2.13.18 - looks to be x64 safe...

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    Name As Long
    base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public functions As New Collection  'of CExport
Public errMessage As String
Public CurrentImage As String

Private mParent As CPEEditor
Private ExportDllName As String 'binary compat already set cant expose :-\

'fixed 11.14.2019 (only took 17 yrs! holy crap i am old)
Friend Sub LoadExports(fhandle As Long, ByVal rvaExportDirectoryTable As Long, parent As CPEEditor)
    
    Dim ied As IMAGE_EXPORT_DIRECTORY
    Dim firstOrdial As Long, oset As Long, pointer As Long, i As Long, nameUb As Long, offsetExportTable As Long, ord As Integer
    Dim func() As Long
    Dim names() As String
    Dim nameOrdinals() As Integer '2 bytes each
    Dim n As String, ns As String

    Set mParent = parent
    
    If rvaExportDirectoryTable < 1 Then Exit Sub

    'ntdll sent me for a loop :(
    'Debug.Print "RvaExportDir: " & Hex(offset.RvaExportDirectory)
    
    offsetExportTable = mParent.RvaToOffset(rvaExportDirectoryTable)
    
    'Debug.Print "File Offset ExportDir: " & Hex(offsetExportTable)
    
    If offsetExportTable = 0 Then
        errMessage = "offsetExportTable = 0"
        Exit Sub
    End If
    
    offsetExportTable = offsetExportTable + 1 '0 based file offsets
    Get fhandle, offsetExportTable, ied
    firstOrdial = ied.base
        
    If ied.Name > 0 Then
        ExportDllName = ExtractExportName(fhandle, mParent.RvaToOffset(ied.Name))
    End If
        
    'Debug.Print "AddressOfFunctions foff: " & Hex(mParent.RvaToOffset(ied.AddressOfFunctions) + 1) & " Count: " & ied.NumberOfFunctions
    'Debug.Print "AddressOfNameOrdinals foff: " & Hex(mParent.RvaToOffset(ied.AddressOfNameOrdinals) + 1)
    'Debug.Print "AddressOfNames foff: " & Hex(mParent.RvaToOffset(ied.AddressOfNames) + 1) & " Count: " & ied.NumberOfNames
    
    'extract ordinals for named exports
    ReDim nameOrdinals(ied.NumberOfNames) 'default all 0
    If ied.AddressOfNameOrdinals > 0 Then
        Get fhandle, mParent.RvaToOffset(ied.AddressOfNameOrdinals) + 1, nameOrdinals
    End If
    
    'extract all named exports names
    oset = mParent.RvaToOffset(ied.AddressOfNames) + 1
    For i = 0 To ied.NumberOfNames - 1
        Get fhandle, (oset + (i * 4)), pointer
        ns = ExtractExportName(fhandle, mParent.RvaToOffset(pointer))
        push names(), ns
        nameOrdinals(i) = nameOrdinals(i) + firstOrdial
        'Debug.Print "Name: "; i & " = " & ns & " ord: " & Hex(nameOrdinals(i) + firstOrdial)
    Next
    
    'build a list of all exported function addresses
    ReDim func(ied.NumberOfFunctions)
    Get fhandle, mParent.RvaToOffset(ied.AddressOfFunctions) + 1, func()
    
    For i = 0 To ied.NumberOfFunctions - 1
        If func(i) <> 0 Then 'unused slots left empty
            ord = i + firstOrdial
            n = OrdForName(nameOrdinals, names, ord)
            'Debug.Print "Fx Addr " & Hex(ord) & ":" & Hex(func(i)) & " " & n
            AddExport func(i), n, CLng(ord)
        End If
    Next

    'OrdialSortCollection
 
End Sub

Private Function OrdForName(ords() As Integer, names() As String, ord As Integer) As String
    Dim i As Integer
    For i = 0 To UBound(ords)
        If ord = ords(i) Then
            OrdForName = names(i)
            Exit Function
        End If
    Next
End Function

Function dumpNames() As Collection
    On Error Resume Next
    Dim c As New Collection
    Dim e As CExport
    
    Set dumpNames = c
    'todo: bug not loading funcs by ordinal?
    
    For Each e In functions
        'MsgBox """" & typename(e.FunctionName)  & """"
        'MsgBox TypeName(e.FunctionName) & " " & Len(e.FunctionName) & " """ & e.FunctionName & """"
        If Len(Trim(Replace(e.FunctionName, Chr(0), ""))) > 0 Then
            c.Add e.FunctionName
        Else
            c.Add "@" & e.FunctionOrdial
        End If
    Next
    
End Function

'Private Function OrdialSortCollection()
'
'    Dim x As CExport
'    Dim tmp As New Collection
'    Dim i As Long
'
'    On Error Resume Next
'    For i = 0 To &HFFF
'        If objKeyExistsInCollection(functions, "ordial:" & i) Then
'            Set x = functions("ordial:" & i)
'            If Not x Is Nothing Then
'                tmp.Add x, "ordial:" & i
'                functions.Remove "ordial:" & i
'            End If
'        End If
'    Next
'
'    For Each x In functions
'        If Len(x.FunctionName) > 0 Then tmp.Add x
'    Next
'
'    Set functions = tmp
'
'
'Exit Function
'hell: DoEvents
'
'End Function

Private Function ExtractExportName(f As Long, pointer As Long) As String
     Dim b As Byte
     Dim tmp As String
     
        If pointer <= 0 Or pointer >= LOF(f) Then
            ExtractExportName = "[No Name]"
        Else
            Get f, (pointer + 1), b
            While b <> 0
                tmp = tmp & Chr(b)
                Get f, , b
            Wend
            ExtractExportName = tmp
        End If
        
End Function


Private Sub AddExport(address As Long, Name As String, Optional ordial As Long)
    Dim m  As New CExport
    m.FunctionAddress = address
    m.FunctionName = Name
    m.FunctionOrdial = ordial
    On Error GoTo hell
    If Not objKeyExistsInCollection(functions, "ordial:" & ordial) Then
        functions.Add m, "ordial:" & ordial
    Else
        If ordial = 0 Then functions.Add m
    End If
    
    Exit Sub
hell:     DoEvents
End Sub



Private Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub



Function objKeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    objKeyExistsInCollection = True
 Exit Function
nope: objKeyExistsInCollection = False
End Function
