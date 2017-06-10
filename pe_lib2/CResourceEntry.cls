VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResourceEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public ID As Long            'only used if in IDs table
Public NameOffset As Long    'only used if in names table

Public DataEntryOffset As Long
Public SubDirOffset As Long

Public Data_RVA As Long
Public Size As Long
Public CodePage As Long
Public Reserved As Long

Public path As String
Public Name As String
Public timeStamp

Function Report() As String
    Dim tmp() As String
    
    'If ID = 0 Then
        push tmp, "Named Entry offset: " & Hex(NameOffset) & " " & Name
    'Else
        push tmp, "ID Entry: " & ID
    'End If
    
    If SubDirOffset <> 0 Then
        push tmp, "Directory Offset: " & Hex(SubDirOffset)
    Else
        push tmp, "DataEntry offset: " & Hex(DataEntryOffset)
        push tmp, "Data_RVA: " & Hex(Data_RVA)
        push tmp, "Size: " & Hex(Size)
    End If
        
    push tmp, "Path: " & path
    push tmp, ""
    Report = Join(tmp, vbCrLf)
    
End Function

Friend Sub setSubDirOffset(DataEntry_orSubDir_Offset As Long)
    
    If (DataEntry_orSubDir_Offset And &H80000000) = &H80000000 Then
        SubDirOffset = DataEntry_orSubDir_Offset Xor &H80000000
    Else
        DataEntryOffset = DataEntry_orSubDir_Offset
    End If
        
End Sub

Property Get NameOrID() As String
    If Len(Name) > 0 Then
        NameOrID = Name
    Else
        NameOrID = ID
    End If
End Property

Property Get DirectoryToName() As String
    Dim v As String
    
    If ID = 0 Then Exit Property
    If ID = 1 Then v = "CURSOR"
    If ID = 2 Then v = "BITMAP"
    If ID = 3 Then v = "ICON"
    If ID = 4 Then v = "MENU"
    If ID = 5 Then v = "DIALOG"
    If ID = 6 Then v = "STRING"
    If ID = 7 Then v = "FONTDIR"
    If ID = 8 Then v = "FONT"
    If ID = 9 Then v = "ACCELERATOR"
    If ID = 10 Then v = "RCDATA"
    If ID = 11 Then v = "MESSAGETABLE"
    If ID = 12 Then v = "GROUP_CURSOR"
    If ID = 14 Then v = "GROUP_ICON"
    If ID = 16 Then v = "VERSION"
    If ID = 17 Then v = "DLGINCLUDE"
    If ID = 19 Then v = "PLUGPLAY"
    If ID = 20 Then v = "VXD"
    If ID = 21 Then v = "ANICURSOR"
    If ID = 22 Then v = "ANIICON"
    If ID = 23 Then v = "HTML"
    If Len(v) = 0 Then v = ID
    
    DirectoryToName = v

End Property

Friend Sub LoadSelf(fhandle As Long, base As Long)
    
    On Error Resume Next
    
    If NameOffset <> 0 Then
        Name = "offset: " & Hex(NameOffset) 'todo load name
    End If
    
    If DataEntryOffset = 0 Then Exit Sub
    
    Dim rde As RESOURCE_DATAENTRY
    Get fhandle, base + DataEntryOffset, rde
    
    CodePage = rde.CodePage
    Data_RVA = rde.Data_RVA   'we will let the user load the data on demand to save memory..they only want what they want..
    Size = rde.Size
    Reserved = rde.Reserved
    
    path = path & "\" & NameOrID()
    
'    'our path is actually wrong at this point because only now is the data loaded..so fix it..
'    Dim a As Long
'    a = InStrRev(path, "\")
'    If a > 0 Then
'        path = Mid(path, 1, a)
'        path = path & NameOrType()
'    End If
            
End Sub

