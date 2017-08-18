VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCmnDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const LANG_US = &H409

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type oColorDlg
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Enum FilterTypes
    textFiles = 0
    htmlFiles = 1
    exeFiles = 2
    zipFiles = 3
    AllFiles = 4
    CustomFilter = 5
End Enum

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As oColorDlg) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hmem As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_LONGNAMES = &H200000
Private o As OPENFILENAME

Private filters(6) As String
Private extensions(6) As String
Private filt As FilterTypes
Private lastDir As String

#If isDZRT Then
    Function FolderDialog2(Optional ByVal initDir = "") As String
        If Len(initDir) <> 0 Then lastDir = initDir
        If FileExists(initDir) Then initDir = GetParentFolder(initDir)
        FolderDialog2 = frmDlg.BrowseForFolder(lastDir)
    End Function
    
    Function ChooseFont(Optional obj As Object) As CFont
        Set ChooseFont = frmSelectFont.SelectFont(Me, obj)
    End Function
#End If

Property Let Filter(x As FilterTypes)
    If x < 0 Or x > 5 Then x = 5
    filt = x
End Property

Property Get Filter() As FilterTypes
    Filter = filt
End Property

Sub SetCustomFilter(displayText As String, Optional wildCardExtMatch = "*.*")
    filters(5) = "____" + Chr$(0) + "___" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    filters(5) = Replace(filters(5), "____", displayText)
    filters(5) = Replace(filters(5), "___", wildCardExtMatch)
    extensions(5) = Replace(wildCardExtMatch, "*", "")
    filt = CustomFilter
End Sub

Private Sub Class_Initialize()
    
    Dim af As String
    Const sf_DESKTOP = &H0
    
    af = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    filters(0) = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + af
    filters(1) = "Html Files (*.htm*)" + Chr$(0) + "*.htm*" + Chr$(0) + af
    filters(2) = "Exe Files (*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + af
    filters(3) = "Zip Files (*.zip)" + Chr$(0) + "*.zip" + Chr$(0) + af
    filters(4) = af

    extensions(0) = "txt"
    extensions(1) = "html"
    extensions(2) = "exe"
    extensions(3) = "zip"
    extensions(4) = "bin"
    
    filt = AllFiles
    lastDir = GetSpecialFolder(sf_DESKTOP)
    
End Sub

Function OpenDialog(Optional initDir = "", Optional title = "Open File") As String
         
    If Len(initDir) <> 0 Then lastDir = initDir
    
    o.lStructSize = Len(o)
    o.hWndOwner = GetForegroundWindow()
    o.hInstance = 0
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = lastDir
    o.lpstrTitle = CStr(title)
    o.flags = 0

    OpenDialog = IIf(GetOpenFileName(o), trim$(o.lpstrFile), "")
    OpenDialog = Replace(OpenDialog, Chr(0), Empty)
    
End Function

Function OpenMulti(Optional initDir = "", Optional title = "Open File") As Collection
         
    Dim c As New Collection
    Set OpenMulti = c
    
    If Len(initDir) <> 0 Then lastDir = initDir
    
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long, i As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long
    
    With tOPENFILENAME
        .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hWndOwner = GetForegroundWindow()
        .nMaxFile = 2048
        .lpstrFilter = filters(filt)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lStructSize = Len(tOPENFILENAME)
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = lastDir
        .hInstance = 0
        .lpstrTitle = CStr(title)
    End With
    
    lResult = GetOpenFileName(tOPENFILENAME)
    
    If lResult > 0 Then
        With tOPENFILENAME
            vFiles = split(Left(.lpstrFile, InStr(.lpstrFile, Chr(0) & Chr(0)) - 1), Chr(0))
            If UBound(vFiles) = 0 Then
                c.add vFiles(0)
            Else
                'otherwise the first element (0) is the folder path, and the rest are file names
                For lIndex = 1 To UBound(vFiles)
                    c.add vFiles(0) & "\" & vFiles(lIndex)
                Next
            End If
        End With
    End If
     
End Function

Function SaveDialog(Optional ByVal defaultFileName, Optional initDir = "", Optional title = "Save File") As String
    
    If Len(initDir) <> 0 Then lastDir = initDir
    
    o.lStructSize = Len(o)
    o.hWndOwner = GetForegroundWindow()
    o.hInstance = 0
    o.lpstrFilter = filters(filt)
    o.lpstrFile = Space$(254)
    o.nMaxFile = 255
    o.lpstrFileTitle = Space$(254)
    o.nMaxFileTitle = 255
    o.lpstrInitialDir = lastDir
    o.lpstrTitle = CStr(title)
    o.lpstrDefExt = extensions(filt)
    o.flags = 0
    
    If Len(defaultFileName) > 0 Then
        o.lpstrFile = CStr(defaultFileName) & Space$(254)
        o.nMaxFile = Len(o.lpstrFile) + 1
    End If

    Dim tmp As String
    tmp = IIf(GetSaveFileName(o), trim$(o.lpstrFile), "")
    If tmp <> "" Then
        If FileExists(tmp) Then
            If MsgBox("File Already Exists" & vbCrLf & vbCrLf & "Are you sure you wish to overwrite existing file?", vbYesNo + vbExclamation, "Confirm Overwrite") = vbYes Then SaveDialog = tmp
        Else
            SaveDialog = tmp
        End If
    Else
       SaveDialog = tmp
    End If
    
    SaveDialog = Replace(SaveDialog, Chr(0), Empty)
       
End Function

Function ColorDialog() As Long
    Dim c As oColorDlg
    Dim cColors() As Byte
    
    c.lStructSize = Len(c)
    c.hWndOwner = GetForegroundWindow()
    c.hInstance = App.hInstance
    c.lpCustColors = StrConv(cColors, vbUnicode, LANG_US)
    c.flags = 0

    If ChooseColor(c) <> 0 Then
        ColorDialog = c.rgbResult
        cColors = StrConv(c.lpCustColors, vbFromUnicode, LANG_US)
    Else
        ColorDialog = -1
    End If
    
End Function

Function FolderDialog() As String
    Dim bInfo As BrowseInfo, ret As String, ptrList As Long, nullChar As Long
    Const BIF_RETURNONLYFSDIRS = 1
    Const BIF_NEWDIALOGSTYLE = &H40

    With bInfo
        .hWndOwner = GetForegroundWindow()
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE
    End With

    ptrList = SHBrowseForFolder(bInfo)
    If ptrList Then
        ret = String$(260, 0)
        SHGetPathFromIDList ptrList, ret 'Get the path from the IDList
        CoTaskMemFree ptrList            'free the block of memory
        nullChar = InStr(ret, vbNullChar)
        If nullChar > 0 Then ret = Left$(ret, nullChar - 1)
    End If

    FolderDialog = Replace(ret, Chr(0), Empty)

End Function

Private Function GetSpecialFolder(sf As SpecialFolders) As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, sf, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      GetSpecialFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
  
End Function

Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x
  
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


