VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'private to dll clients dont use it..

'Public Entries As New Collection 'of CResEntry

'Private Enum ResTypes
'   RT_CURSOR = 1&
'   RT_BITMAP = 2&
'   RT_ICON = 3&
'   RT_MENU = 4&
'   RT_DIALOG = 5&
'   RT_STRING = 6&
'   RT_FONTDIR = 7&
'   RT_FONT = 8&
'   RT_ACCELERATOR = 9&
'   RT_RCDATA = 10&
'   RT_MESSAGETABLE = 11&
'   RT_GROUP_CURSOR = 12&
'   RT_GROUP_ICON = 14&
'   RT_VERSION = 16&
'   RT_DLGINCLUDE = 17&
'   RT_PLUGPLAY = 19&
'   RT_VXD = 20&
'   RT_ANICURSOR = 21&
'   RT_ANIICON = 22&
'   RT_HTML = 23&
'End Enum

Private Function ResIdToName(id) As String
    Dim v As String
    
    If id = 1 Then v = "CURSOR"
    If id = 2 Then v = "BITMAP"
    If id = 3 Then v = "ICON"
    If id = 4 Then v = "MENU"
    If id = 5 Then v = "DIALOG"
    If id = 6 Then v = "STRING"
    If id = 7 Then v = "FONTDIR"
    If id = 8 Then v = "FONT"
    If id = 9 Then v = "ACCELERATOR"
    If id = 10 Then v = "RCDATA"
    If id = 11 Then v = "MESSAGETABLE"
    If id = 12 Then v = "GROUP_CURSOR"
    If id = 14 Then v = "GROUP_ICON"
    If id = 16 Then v = "VERSION"
    If id = 17 Then v = "DLGINCLUDE"
    If id = 19 Then v = "PLUGPLAY"
    If id = 20 Then v = "VXD"
    If id = 21 Then v = "ANICURSOR"
    If id = 22 Then v = "ANIICON"
    If id = 23 Then v = "HTML"
    If Len(v) = 0 Then v = id
    
    ResIdToName = v

End Function


Friend Sub LoadDir(p As CResources, fileOffset As Long, Optional pth As String)
    
    Dim r As CResEntry
    Dim resRootDir As RESDIRECTORY
    Dim e() As RESOURCE_DIRECTORY_ENTRY
    Dim d As RESOURCE_DATAENTRY
    Dim cd As CResData
    
    'Set parent = p
    
    Get p.fhandle, fileOffset, resRootDir
    
    ReDim e(resRootDir.NumberOfNamedEntries + resRootDir.NumberOfIdEntries - 1)
    Get p.fhandle, , e

    For i = 0 To UBound(e)
        Set r = New CResEntry
        
        If i < resRootDir.NumberOfNamedEntries Then
            r.setName p.fhandle, p.fileOffset, e(i).NameOffset_or_ID
            r.path = pth & "\" & r.Name
        Else
            r.id = e(i).NameOffset_or_ID
            If Len(pth) = 0 Then
                r.path = "\" & ResIdToName(r.id)
            Else
                r.path = pth & "\" & r.id
            End If
        End If
        
        r.setData e(i).DataEntry_orSubDir_Offset
        
        If r.isSubDir Then
            r.subDir.LoadDir p, p.fileOffset + r.SubDirOffset, r.path
        Else
            'finally!, note we add the actual data entry to top level parent.Entries collection for simplicity
            'the path property could be used to recreate the tree if a client really wanted..I dont want to deal with that though..
            Get p.fhandle, p.fileOffset + r.DataOffset, d
            Set cd = New CResData
            cd.CodePage = d.CodePage
            cd.OffsetToDataRVA = d.Data_RVA
            cd.Reserved = d.Reserved
            cd.Size = d.Size
            cd.path = r.path
            cd.timeStamp = timeStampToDate(resRootDir.TimeDateStamp)
            p.Entries.Add cd
        End If
            

        'Debug.Print Hex(r.NameId) & " " & Hex(r.OffsetToData)
        'Entries.Add r
    Next
    
End Sub
