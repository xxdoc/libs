VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'do not hang onto a ref of this class and reload the pe editor with something else

'example paths: \ICON\30001\0
'               \VERSION\1\1033

Private parent As CPEEditor

Public Entries As New Collection 'of CResourceEntry

Public Enum ResTypes
   RT_CURSOR = 1&
   RT_BITMAP = 2&
   RT_ICON = 3&
   RT_MENU = 4&
   RT_DIALOG = 5&
   RT_STRING = 6&
   RT_FONTDIR = 7&
   RT_FONT = 8&
   RT_ACCELERATOR = 9&
   RT_RCDATA = 10&
   RT_MESSAGETABLE = 11&
   RT_GROUP_CURSOR = 12&
   RT_GROUP_ICON = 14&
   RT_VERSION = 16&
   RT_DLGINCLUDE = 17&
   RT_PLUGPLAY = 19&
   RT_VXD = 20&
   RT_ANICURSOR = 21&
   RT_ANIICON = 22&
   RT_HTML = 23&
End Enum

Public rva As Long
Public fileOffset As Long
Public size As Long
Public hadRecursive As Boolean

Private EntryOffsets As New Collection 'make sure no recursive loops..

Friend Sub LoadSelf(owner As CPEEditor)
    
    Set parent = owner
    
    Dim hfile As Long, i As Long
    Dim resRootDir As RESDIRECTORY
    Dim cre As CResourceEntry
    
    rva = owner.OptionalHeader.ddVirtualAddress(Resource_Table)
    size = owner.OptionalHeader.ddSize(Resource_Table)
    fileOffset = owner.RvaToOffset(rva)
    
    If rva = 0 Or size = 0 Or fileOffset = 0 Then Exit Sub
    fileOffset = fileOffset + 1
    
    hfile = FreeFile
    Open owner.LoadedFile For Binary Access Read As #hfile
    Get hfile, fileOffset, resRootDir
    LoadSubDir hfile, resRootDir 'this is recursive...
       
    'collapsed down final entries list
    For Each cre In Entries
        cre.LoadSelf hfile, fileOffset
        Debug.Print cre.Report
    Next

    Close hfile


End Sub

Private Sub LoadSubDir(hfile As Long, resDir As RESDIRECTORY, Optional path As String)
    
    Dim cre As CResourceEntry
    Dim subDirs As New Collection
    Dim childDir As RESDIRECTORY
    Dim resDirEntry As RESOURCE_DIRECTORY_ENTRY
    Dim childOffset As Long
    Dim childPath As String
    
    'all of these structures are sequential in the file..
    'so load them one after another, no need to track file pointer..
    
    For i = 1 To resDir.NumberOfNamedEntries
        Get hfile, , resDirEntry
        
        Set cre = New CResourceEntry
        cre.setSubDirOffset resDirEntry.DataEntry_orSubDir_Offset
        cre.NameOffset = resDirEntry.NameOffset_or_ID
        cre.path = path
        
        If cre.SubDirOffset <> 0 Then
            If EntryExists(cre.SubDirOffset) Then
                hadRecursive = True
                Debug.Print "We caught one with a recursive SubDirOffset!"
            Else
                subDirs.Add cre
            End If
        Else
            If EntryExists(cre.DataEntryOffset) Then
                hadRecursive = True
                Debug.Print "We caught one with a recursive DataEntryOffset!"
            Else
                Entries.Add cre
            End If
        End If
        
        'Debug.Print cre.Report
    Next
    
    For i = 1 To resDir.NumberOfIdEntries
        Get hfile, , resDirEntry
        
        Set cre = New CResourceEntry
        cre.setSubDirOffset resDirEntry.DataEntry_orSubDir_Offset
        cre.ID = resDirEntry.NameOffset_or_ID
        cre.path = path
        
        If cre.SubDirOffset <> 0 Then
            If EntryExists(cre.SubDirOffset) Then
                hadRecursive = True
                Debug.Print "We caught one with a recursive SubDirOffset!"
            Else
                subDirs.Add cre
            End If
        Else
            If EntryExists(cre.DataEntryOffset) Then
                hadRecursive = True
                Debug.Print "We caught one with a recursive DataEntryOffset!"
            Else
                Entries.Add cre
            End If
        End If
        
        'Debug.Print cre.Report
    Next
 
    For Each cre In subDirs
        childOffset = fileOffset + cre.SubDirOffset
        'DumpFile hfile, childOffset
        Get hfile, childOffset, childDir
        
        If Len(path) = 0 Then
            childPath = "\" & cre.DirectoryToName()
        Else
            childPath = path & "\" & cre.NameOrID
        End If
        
        LoadSubDir hfile, childDir, childPath
    Next
    
    
End Sub

Private Function EntryExists(offset As Long) As Boolean
    On Error Resume Next
    Dim o
    o = EntryOffsets("o:" & offset)
    If Err.Number = 0 Then
        EntryExists = True
        EntryOffsets.Add "o:" & offset
    End If
End Function

Private Sub DumpFile(hfile As Long, offset As Long, Optional size As Long = 16)
    Dim curPos As Long
    Dim b() As Byte
    ReDim b(1 To size)
    
    curPos = Seek(hfile)
    Get hfile, offset, b()
    Seek hfile, curPos
    
    Debug.Print Right("00000000" & Hex(offset), 8) & HexDump(StrConv(b, vbUnicode))
    
End Sub

Public Function GetResource(path As String, ByRef outRetVal As CResourceEntry) As Boolean
    Set outRetVal = Nothing
    For Each outRetVal In Entries
        If LCase(outRetVal.path) = LCase(path) Then
            GetResource = True
            Exit Function
        End If
    Next
    Set outRetVal = Nothing
End Function

Public Function GetResourceData(path As String, ByRef bArray) As Boolean
    Dim cre As CResourceEntry
    Dim f As Long
    Dim b() As Byte
    Dim fOffset As Long
    
    Erase bArray
    If Not FileExists(parent.LoadedFile) Then Exit Function
    If Not GetResource(path, cre) Then Exit Function
    
    fOffset = parent.RvaToOffset(cre.Data_RVA)
    If cre.size = 0 Or cre.Data_RVA = 0 Or fOffset = 0 Then Exit Function
    
    ReDim b(cre.size - 1)
    f = FreeFile
    Open parent.LoadedFile For Binary As f
    Get f, fOffset + 1, b
    Close f
    
    bArray = b()
    GetResourceData = True
    
End Function

Public Function SaveResource(filePath As String, resPath As String) As Boolean
    Dim b() As Byte, f As Long
    If FileExists(filePath) Then Kill filePath
    If Not GetResourceData(resPath, b) Then Exit Function
    
    f = FreeFile
    Open filePath For Binary As f
    Put f, , b()
    Close f
    
    If FileLen(filePath) = UBound(b) + 1 Then SaveResource = True
End Function
