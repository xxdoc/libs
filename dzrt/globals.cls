VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "globals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const LB_GETCOUNT = &H18B
Private Const LB_SETTOPINDEX = &H197
Private startTime As Long

Private Declare Function GetTickCount_ Lib "kernel32" Alias "GetTickCount " () As Long
Private Declare Function LoadLibrary_ Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary_ Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SendMessage_ Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub Sleep_ Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
 
Private Declare Sub vc_srand Lib "utypes.dll" (ByVal v1 As Long)
Private Declare Function vc_rand Lib "utypes.dll" () As Long
 
Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

'all but jet are odbc jet is ole
Public Enum dbServers
    Access
    JetAccess2k
    MsSql2k
    mysql
    DSN
    FileDsn
    dBase
End Enum

Public Enum CompressTypes
    ct_Zlib = 0
    ct_LZNT1 = 1      'so the compress seems to work..but is 1 byte short?
    ct_LZNT1_MAX = 2  'the decompress works with other known samples, but not with one compressed with our compress..havent been able to fix the error yet!
End Enum

Public Enum InputFormats
    ifHex = 1
    ifDecimal = 2
    ifString = 3
    ifHexDump = 4
End Enum

Public Enum hexOutFormats
    hoDump
    hoSpaced
    hoHexOnly
End Enum

Private OnBits(0 To 31) As Long
Private crc_table() As Long

Public base64 As New CBase64
Public hash As New CWinHash
Public fso As New CFileSystem3
Public cmdline As New CCmdLine
Public zlib As New CZlib

Sub srand(seed As Long)
    If ensureUTypes() Then vc_srand seed
End Sub

Function rand() As Long
    If ensureUTypes() Then rand = vc_rand()
End Function


'this function will convert any of the following to a byte array:
'   read a file if path supplied and allowFilePaths = true
'   byte(), integer() or long() arrays
'   all other data types it will attempt to convert them to string, then to byte array
'   if the data type you pass can not be converted with cstr() it will throw an error.
'   no other types make sense to support explicitly
'   this assumes all arrays are 0 based..
Function LoadData(fileStringOrByte, Optional allowFilePaths As Boolean = True) As Byte()
    LoadData = modGeneral.LoadData(fileStringOrByte, allowFilePaths)
End Function

Function cast(varIn, varOut, Optional allowFilePaths As Boolean = False) As Boolean
    On Error Resume Next
    Dim b() As Byte
    Dim i() As Integer
    Dim l() As Long
    Dim c() As Currency
    Dim d() As Double
    
    Dim sz As Long
    Dim bytes As Long
    
    On Error GoTo hell
    
    If TypeName(varIn) = "Object" Or TypeName(varOut) = "Object" Then
        Err.Raise "Invalid use of cast object type not supported"
    End If
    
    If Not allowFilePaths Then
        If TypeName(varIn) = TypeName(varOut) Then
            varOut = varIn
            Exit Function
        End If
    End If
        
    b() = LoadData(varIn, allowFilePaths)
    bytes = UBound(b) + 1
    
    If TypeName(varOut) = "String" Or TypeName(varOut) = "Variant" Then
        varOut = StrConv(b, vbUnicode, LANG_US)
    ElseIf TypeName(varOut) = "Byte()" Then
        varOut = b()
    ElseIf TypeName(varOut) = "Integer()" Then
        If bytes Mod 2 <> 0 Then
            sz = bytes + (bytes Mod 2)
            ReDim Preserve b(sz)
        End If
        ReDim i((bytes / 2) - 1)
        CopyMemory_ ByVal VarPtr(i(0)), ByVal VarPtr(b(0)), bytes
        varOut = i()
    ElseIf TypeName(varOut) = "Long()" Then
        If bytes Mod 4 <> 0 Then
            sz = bytes + (bytes Mod 4)
            ReDim Preserve b(sz)
        End If
        ReDim l((bytes / 4) - 1)
        CopyMemory_ ByVal VarPtr(l(0)), ByVal VarPtr(b(0)), bytes
        varOut = l()
    ElseIf TypeName(varOut) = "Double()" Then
        If bytes Mod 8 <> 0 Then
            sz = bytes + (bytes Mod 8)
            ReDim Preserve b(sz)
        End If
        ReDim d((bytes / 8) - 1)
        CopyMemory_ ByVal VarPtr(d(0)), ByVal VarPtr(b(0)), bytes
        varOut = d()
    ElseIf TypeName(varOut) = "Currency()" Then
        If bytes Mod 8 <> 0 Then
            sz = bytes + (bytes Mod 8)
            ReDim Preserve b(sz)
        End If
        ReDim c((bytes / 8) - 1)
        CopyMemory_ ByVal VarPtr(c(0)), ByVal VarPtr(b(0)), bytes
        varOut = c()
    Else
        varOut = Empty
        cast = False
        Exit Function
    End If
    
    cast = True
    
Exit Function
hell:
    cast = False

End Function

Function GetTickCount() As Long
    GetTickCount = GetTickCount_()
End Function

Function LoadLibrary(ByVal lpLibFileName As String) As Long
    LoadLibrary = LoadLibrary_(lpLibFileName)
End Function

 Function FreeLibrary(ByVal hLibModule As Long) As Long
    FreeLibrary = FreeLibrary_(hLibModule)
 End Function
 
 Sub Sleep(ByVal dwMilliseconds As Long)
    Sleep_ dwMilliseconds
 End Sub
 
 Sub CopyMemory(Destination As Long, Source As Long, ByVal Length As Long)
     CopyMemory_ ByVal Destination, ByVal Source, Length
 End Sub
 
 Function SendMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    SendMessage = SendMessage_(hWnd, msg, wParam, lParam)
 End Function

Property Get MAX_LONG() As Long
    MAX_LONG = 2147483647
End Property

Property Get MIN_LONG() As Long
    MIN_LONG = -2147483648#
End Property

Property Get MAX_INT() As Integer
    MAX_INT = 32767
End Property

Property Get MIN_INT() As Integer
    MIN_INT = -32768
End Property

Property Get LANG_US() As Long
    LANG_US = &H409
End Property

Sub StartBenchMark(Optional ByRef t As Long)
    If IsMissing(t) Then
        startTime = GetTickCount()
    Else
        t = GetTickCount()
    End If
End Sub

Function EndBenchMark(Optional ByRef t As Long) As String
    Dim endTime As Long, loadTime As Long
    endTime = GetTickCount()
    If IsMissing(t) Then
        loadTime = endTime - startTime
    Else
        loadTime = endTime - t
    End If
    EndBenchMark = loadTime / 1000 & " seconds"
End Function

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function


Function TopMost(frm As Object, Optional ontop As Boolean = True)
    On Error Resume Next
    Dim s
    s = IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos frm.hWnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Public Function GetShortName(ByVal sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name !!
    If Not FileExists(sFile) Then
        iCreated = True
        WriteFile sFile, ""
    End If
    
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile
    
    If iCreated Then Kill sFile
    
End Function

Sub FormPos(fform As Object, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Sub SaveMySetting(key, Value)
    SaveSetting App.EXEName, "Settings", key, Value
End Sub

Function GetMySetting(key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", key, defaultval)
End Function

Function FolderExists(path) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Asc(Left(tmp, 1)) = 0 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

'can load a byte array, string, or string array of lines split at either crlf or lf
Function ReadFile(filename, ByRef data) As Boolean 'this one should be binary safe...
  On Error GoTo hell
  Dim f As Long
  Dim b() As Byte
  Dim tmp As String
  Dim t2() As String
  
  f = FreeFile
  Open filename For Binary As #f
  ReDim b(lof(f) - 1)
  Get f, , b()
  Close #f
  
  If TypeName(data) = "Byte()" Then 'Or TypeName(data) = "Variant"
        data = b()
  ElseIf TypeName(data) = "String()" Then
        tmp = CStr(StrConv(b(), vbUnicode, LANG_US))
        If InStr(tmp, vbCrLf) > 0 Then
            t2 = split(tmp, vbCrLf)
        Else
            t2 = split(tmp, vbLf)
        End If
        data = t2
  Else
  
        data = CStr(StrConv(b(), vbUnicode, LANG_US))
  End If
  
  ReadFile = True
  
  Exit Function
hell:   ReadFile = False
End Function

Function WriteFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte, f As Long
    
    If FileExists(path) Then Kill path
    f = FreeFile
    
    If TypeName(it) = "Byte()" Then
        b() = it
    Else
        b() = StrConv(it, vbFromUnicode, LANG_US)
    End If
    
    Open path For Binary As #f
    Put f, , b()
    Close f
    
    WriteFile = True
    Exit Function
    
hell: WriteFile = False
End Function

Function AppendFile(path, it) As Boolean  'not binary safe
    Dim f As Long
    On Error Resume Next
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
    AppendFile = (Err.Number = 0)
End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
       
    x = UBound(ary)
    ReDim Preserve ary(x + 1)
    
    If IsObject(Value) Then
        Set ary(x + 1) = Value
    Else
        ary(x + 1) = Value
    End If
    
    Exit Sub
init:
    ReDim ary(0)
    If IsObject(Value) Then
        Set ary(0) = Value
    Else
        ary(0) = Value
    End If
End Sub

Function pop(ary, Optional count As Long = 1) As Variant
    On Error GoTo isEmpty
    Dim x
    
    x = UBound(ary)
    
    If count > x Then
        Erase ary
        Exit Function
    End If
    
    ReDim Preserve ary(x - count + 1)
        
    If IsObject(ary(x)) Then
        Set pop = ary(x)
    Else
        pop = ary(x)
    End If
    
    ReDim Preserve ary(x - count)
    
    Exit Function
isEmpty: Erase ary
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim x
    If IsObject(UBound(ary)) Then AryIsEmpty = False
    'x = UBound(ary)
  Exit Function
oops: AryIsEmpty = True
End Function

Function keyExists(ByVal t As String, c As Collection) As Boolean
    On Error GoTo hell
    Dim x
    
    If IsObject(c(t)) Then
        Set x = c(t)
    Else
        x = c(t)
    End If
    
    keyExists = True
Exit Function
hell:
End Function

Function anyOfTheseInstr(data, match, Optional compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim tmp() As String
    Dim x
    tmp = split(match, ",")
    For Each x In tmp
        If InStr(1, data, x, compare) > 0 Then
            anyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Public Function isHexNum(v, Optional ByRef dataOut) As Boolean
    On Error Resume Next
    dataOut = CLng("&h" & v)
    If Err.Number = 0 Then isHexNum = True
End Function

Function lbCount(hListBox As Long) As Long
   lbCount = SendMessage(hListBox, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
End Function

Public Sub lbScrollTo(hListBox As Long, Optional ByVal index As Long = -1)
    If index = -1 Then index = lbCount(hListBox)
    SendMessage hListBox, LB_SETTOPINDEX, ByVal index - 1, ByVal 0
End Sub

Function lbCopy(lstBox As Object) As String
    
    Dim i As Long
    Dim tmp() As String
    
    For i = 0 To lstBox.ListCount
        push tmp, lstBox.List(i)
    Next
    
    lbCopy = Join(tmp, vbCrLf)
    
End Function

Sub lbAppend(lstBox As Object, txt, Optional ByVal index = -1)
    On Error GoTo hell
    If index = -1 Then index = lstBox.ListCount - 1
    If index < 0 Then Exit Sub
    lstBox.List(index) = lstBox.List(index) & txt
hell:
End Sub

Sub lbUpdate(lstBox As Object, txt, Optional ByVal index = -1)
    On Error GoTo hell
    If index = -1 Then index = lstBox.ListCount - 1
    If index < 0 Then Exit Sub
    lstBox.List(index) = txt
hell:
End Sub


Public Sub lvColumnSort(ListViewControl As Object, Column As Object)
    On Error Resume Next
    Const lvwAscending As Long = 0
    Const lvwDescending As Long = 1
     
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .sorted = -1
    End With
End Sub

Function lvGetAllElements(lv As Object) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next
    
    push ret, tmp
    push ret, String(50, "-")
        
    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next
    
    lvGetAllElements = Join(ret, vbCrLf)
    
End Function

Function lvGetColumn(lv As Object, colIndex As Long) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li 'As ListItem
    
    On Error Resume Next
    
    If colIndex > lv.ColumnHeaders.count Then Exit Function
    If colIndex < 0 Then Exit Function
    
    'push ret, lv.ColumnHeaders(colIndex + 1).Text
    'push ret, String(50, "-")
        
    For Each li In lv.ListItems
        If colIndex = 0 Then
            push ret, li.Text
        Else
            push ret, li.SubItems(colIndex)
        End If
    Next
    
    lvGetColumn = Join(ret, vbCrLf)
    
End Function

Function Rpad(v, Optional l As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        Rpad = v & String(l - x, char)
    Else
hell:
        Rpad = v
    End If
End Function

Function lpad(v, Optional l As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        lpad = String(l - x, char) & v
    Else
hell:
        lpad = v
    End If
End Function

Function printf(ByVal msg As String, ParamArray vars()) As String
    printf = modGeneral.printf(msg, paramsToArray(vars))
End Function

Function ado_buildInsert(tblName, fields, ParamArray params()) As String

    Dim sSql As String, sTmp As String, i As Integer, x
    Dim values()
    
    values() = params() 'force byval
    sSql = "Insert into " & tblName & " (" & fields & ") VALUES(____)"
    
    Dim tn As String
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "date" Then
            values(i) = "'" & replace(values(i), "'", "''") & "'"
        End If
    Next
    
    sTmp = Join(values, ",")
    ado_buildInsert = replace(sSql, "____", sTmp)
     
    
End Function

Function ado_buildUpdate(tblName, criteria, fields, ParamArray params()) As String

    Dim sSql As String, sTmp As String, i As Integer
    Dim values()
    
    values() = params() 'force byval
    sSql = "Update " & tblName & " SET ____" & " " & criteria
    
    Dim tn As String
    For i = 0 To UBound(values)
        tn = LCase(TypeName(values(i)))
        If tn = "string" Or tn = "textbox" Or tn = "date" Then
            values(i) = "'" & replace(values(i), "'", "''") & "'"
        End If
    Next
    
    Dim tmp() As String
    tmp() = split(fields, ",")
    
    If UBound(tmp()) <> UBound(values()) Then MsgBox "Err in update fields/values"
    
    For i = 0 To UBound(tmp)
        tmp(i) = tmp(i) & "=" & values(i)
    Next
    
    sTmp = Join(tmp, ",")
    ado_buildUpdate = replace(sSql, "____", sTmp)
    
End Function


''endian swap and return spaced out hex string
'Private Function lng2Hex(x As Long) As String
'    Dim b(1 To 4) As Byte
'    CopyMemory b(1), x, 4
'    lng2Hex = Hex(b(1)) & " " & Hex(b(2)) & " " & Hex(b(3)) & " " & Hex(b(4))
'End Function


Function isPrintable(b As Byte) As Boolean
     
    isPrintable = True
    If b = 9 Or b = 10 Or b = 13 Then
        'tab, cr, lf are ok..
    ElseIf b >= &H20 And b <= &H7E Then
        'its a printable character and is ok..
    Else
        isPrintable = False
    End If
 
End Function

Function countOccurances(it, find) As Integer
    Dim tmp() As String
    
    If InStr(1, it, find, vbTextCompare) < 1 Then Exit Function
    tmp = split(it, find, , vbTextCompare)
    countOccurances = UBound(tmp)
    
End Function

Function h2l(ByVal x, ByRef outVal As Long) As Boolean
    On Error Resume Next
    x = trim(replace(x, "0x", Empty, , , vbTextCompare))
    x = trim(replace(x, "&h", Empty, , , vbTextCompare))
    outVal = CLng("&h" & x)
    h2l = (Err.Number = 0)
End Function

Function HexDump(bAryOrStrData, Optional ByVal Length As Long = -1, Optional ByVal startAt As Long = 1, Optional hexFormat As hexOutFormats = hoDump) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, x
    Dim hexOnly As Long
    
    offset = 0
    If hexFormat <> hoDump Then hexOnly = 1
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        ary() = bAryOrStrData
    Else
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    If startAt < 1 Then startAt = 1
    If Length < 1 Then Length = -1
    
    While startAt Mod 16 <> 0
        startAt = startAt - 1
    Wend
    
    startAt = startAt + 1
    
    chars = "   "
    For i = startAt To UBound(ary) + 1
        tt = Hex(ary(i - 1))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i - 1)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
        If Length <> -1 Then
            Length = Length - 1
            If Length = 0 Then Exit For
        End If
    Next
    
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        If hexFormat = hoHexOnly Then HexDump = replace(HexDump, " ", "")
        HexDump = replace(HexDump, vbCrLf, "")
    End If
    
End Function

Public Function toBytes(ByVal hexstr, ByRef outVar, Optional ByVal inputformat As InputFormats = ifHex) As Boolean

'supports:
'11 22 33 44   spaced hex chars
'11223344      run together hex strings
'11,22,33,44   csv hex
'1,2,3,4       csv hex with no lead 0
'121,99,44,255 decimal csv or spaced values
'%xx%yy
'%uxxxx\u7766
'%u6162%63
'isDecimal flag requires csv or spaced values..
'ignores common C source prefixes and characters

    Dim ret As String, x As String, str As String
    Dim r() As Byte, b As Byte, b1 As Byte
    Dim foundDecimal As Boolean, tmp, i, a, a2
    Dim pos As Long, marker As String
    
    On Error GoTo hell
    
    If inputformat = ifHexDump Then
        x = ExtractHexFromDump(hexstr)      'returns just the hex string
        If Not toBytes(x, r) Then GoTo hell 'now we convert it to actual bytes..
        GoTo retNow
    End If
    
    If inputformat = ifString Then
        r() = StrConv(hexstr, vbFromUnicode, LANG_US)
        GoTo retNow
    End If
    
    str = replace(hexstr, vbCr, Empty)
    str = replace(str, vbLf, Empty)
    str = replace(str, vbTab, Empty)
    str = replace(str, Chr(0), Empty)
    str = replace(str, "{", Empty)
    str = replace(str, "}", Empty)
    str = replace(str, ";", Empty)
    str = replace(str, "+", Empty)
    str = replace(str, """""", Empty)
    str = replace(str, "'", Empty)
    hexstr = str
    
    If InStr(hexstr, "\u") > 0 Then hexstr = replace(hexstr, "\u", "%u")
    
    If InStr(hexstr, "%u") > 0 Then
        tmp = split(hexstr, "%u")
        For i = 1 To UBound(tmp)
            a = InStr(tmp(i), "%")
            x = ""
            If a > 1 Then
                x = Mid(tmp(i), a)
                tmp(i) = Mid(tmp(i), 1, a - 1)
            End If
            If Len(tmp(i)) = 3 Then tmp(i) = "0" & tmp(i)
            If Len(tmp(i)) = 4 Then
                a = Mid(tmp(i), 1, 2)
                a2 = Mid(tmp(i), 3, 2)
                tmp(i) = a2 & a
            End If
            If Len(x) > 0 Then tmp(i) = tmp(i) & x
        Next
        hexstr = Join(tmp, "")
    End If
    
    If InStr(hexstr, "%") > 0 Then
        tmp = split(hexstr, "%")
        For i = 1 To UBound(tmp)
            If Len(tmp(i)) < 2 Then
                tmp(i) = 0 & tmp(i)
            End If
        Next
        hexstr = Join(tmp, "")
    End If
    
    If Len(hexstr) > 4 Then
        pos = FirstOccurance(hexstr, " ,", marker)
        If pos > 0 And pos < 5 Then   'make sure all are double digit hex chars...(also account for decimal 1,11,111,
            tmp = split(hexstr, marker)
            
            If inputformat = ifDecimal Then
                For i = 0 To UBound(tmp)
                    tmp(i) = Hex(CLng(tmp(i)))
                Next
            End If
            
            For i = 0 To UBound(tmp)
                If Len(tmp(i)) = 1 Then tmp(i) = "0" & tmp(i)
            Next
            
            hexstr = Join(tmp, "")
        End If
    End If
        
    str = replace(hexstr, " ", Empty)
    str = replace(str, "0x", Empty)
    str = replace(str, ",", Empty)
    
    For i = 1 To Len(str) Step 2
        x = Mid(str, i, 2)
        If Not isHexChar(x, b) Then Exit Function
        bpush r(), b
    Next
    
retNow:
    If TypeName(outVar) = "Byte()" Then
        outVar = r
    Else
        outVar = StrConv(r, vbUnicode, LANG_US)
    End If
    
    toBytes = True
    Exit Function
    
hell:
    toBytes = False
    
End Function

Function FirstOccurance(it, ByVal csvFind As String, ByRef outFoundVal) As Long
    If Len(csvFind) = 0 Then Exit Function
    
    Dim find() As String, x, lowestOffset As Long, lowestIndex As Long, i As Long, a As Long
    
    outFoundVal = Empty
    lowestOffset = MAX_LONG
    find = split(csvFind, ",")
    
    For i = 0 To UBound(find)
        If Len(find(i)) = 0 Then find(i) = ","
        a = InStr(1, it, find(i), vbTextCompare)
        If a > 0 And a < lowestOffset Then
            lowestOffset = a
            lowestIndex = i
        End If
    Next
    
    If lowestOffset = MAX_LONG Then Exit Function
    
    outFoundVal = find(lowestIndex)
    FirstOccurance = lowestOffset
    
End Function

Private Sub bpush(bAry() As Byte, b As Byte) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    
    x = UBound(bAry) '<-throws Error If Not initalized
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    
    Exit Sub

init:
    ReDim bAry(0)
    bAry(0) = b
    
End Sub

Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope 'expecting hex char code like FF or 90
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope 'invalid hex code
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope  'shouldnt happen.. > 255 cant be with len() <=2 ?

    isHexChar = True
    
    Exit Function
nope:
    Err.Clear
    isHexChar = False
End Function

Public Function xorDecode(fileByteOrString As Variant, ByVal password As Variant, Optional strRet As Boolean = True)
    Dim key() As Byte, ByteArray() As Byte, ko As Long, i As Long
    
    Dim plen As Long
    
    ByteArray() = LoadData(fileByteOrString)
    
    If TypeName(password) = "Byte()" Then
        key() = password
    ElseIf TypeName(password) = "Byte" Then
        ReDim key(0)
        key(0) = password
    Else
        If Len(password) = 0 Then Exit Function
         key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
    End If
    
    For i = 0 To UBound(ByteArray)
        ko = i Mod (UBound(key) + 1)
        ByteArray(i) = ByteArray(i) Xor key(ko)
    Next
    
    If strRet Then
        xorDecode = StrConv(ByteArray, vbUnicode, LANG_US)
    Else
        xorDecode = ByteArray
    End If
    
End Function


Public Function rc4(fileByteOrString As Variant, ByVal password As Variant, Optional strRet As Boolean = True)
On Error Resume Next
    Dim RB(0 To 255) As Integer, x As Long, Y As Long, Z As Long, key() As Byte, ByteArray() As Byte, temp As Byte
    
    Dim plen As Long
    
    ByteArray() = LoadData(fileByteOrString)
    
    If TypeName(password) = "Byte()" Then
        key() = password
        If UBound(key) > 255 Then ReDim Preserve key(255)
    Else
        If Len(password) = 0 Then
            Exit Function
        End If

        If Len(password) > 256 Then
            key() = StrConv(Left$(CStr(password), 256), vbFromUnicode, LANG_US)
        Else
            key() = StrConv(CStr(password), vbFromUnicode, LANG_US)
        End If
    End If
    
    plen = UBound(key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For x = 0 To 255
        RB(x) = x
    Next x
    
    x = 0
    Y = 0
    Z = 0
    For x = 0 To 255
        Y = (Y + RB(x) + key(x Mod plen)) Mod 256
        temp = RB(x)
        RB(x) = RB(Y)
        RB(Y) = temp
    Next x
    
    x = 0
    Y = 0
    Z = 0
    For x = 0 To UBound(ByteArray)
        Y = (Y + 1) Mod 256
        Z = (Z + RB(Y)) Mod 256
        temp = RB(Y)
        RB(Y) = RB(Z)
        RB(Z) = temp
        ByteArray(x) = ByteArray(x) Xor (RB((RB(Y) + RB(Z)) Mod 256))
    Next x
    
    If strRet Then
        rc4 = StrConv(ByteArray, vbUnicode, LANG_US)
    Else
        rc4 = ByteArray
    End If
    
End Function

Function b64Encode(bArrayStringOrFile, Optional retAsString As Boolean = True) As Variant
    b64Encode = base64.Encode(bArrayStringOrFile, retAsString)
End Function

Function b64Decode(bArrayStringOrFile, ByRef dataOut As Variant, Optional retAsString As Boolean = True) As Boolean
     b64Decode = base64.Decode(bArrayStringOrFile, dataOut, retAsString)
End Function

Function entropy(fileByteOrString, Optional offset As Long = 0, Optional leng As Long = -1) As Single
    
    If FileExists(fileByteOrString) Then
        entropy = fileEntropy(CStr(fileByteOrString), offset, leng)
        Exit Function
    End If
    
    Dim b() As Byte
    b() = LoadData(fileByteOrString, False)
    entropy = memEntropy(b, offset, leng)
        
End Function

Function crc32(fileByteOrString) As String

    Dim b() As Byte
    Dim v As Long
    
    If Not ensureUTypes() Then Exit Function 'use the C dll version its faster
    
    b() = LoadData(fileByteOrString)
    

    If AryIsEmpty(b) Then Exit Function
    
    v = ut_crc32(b(0), UBound(b) + 1)
    crc32 = Hex(v)
    
'    vb native implementation...
'    Dim c As Long, n As Long, x As Long
'    Dim b() As Byte
'
'    c = -1
'    If AryIsEmpty(crc_table) Then make_crc_table
'
'    If TypeName(bAryOrString) = "Byte()" Then
'        b() = bAryOrString
'    Else
'        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
'    End If
'
'    For n = 0 To UBound(b)
'        c = crc_table((c Xor b(n)) And &HFF) Xor rshift(c, 8)
'    Next
'
'    crc32 = Hex(c Xor &HFFFFFFFF)
    
End Function

'Private Sub make_crc_table()
'    Dim c As Long, n As Long, k As Long
'
'    ReDim crc_table(256)
'
'    For n = 0 To 255
'          c = n
'          For k = 0 To 7
'                If c And 1 Then
'                     c = &HEDB88320 Xor rshift(c)
'                Else
'                    c = rshift(c)
'                End If
'           Next
'          crc_table(n) = c
'    Next
'
'End Sub

'Public Function lshift(ByVal Value As Long, Optional ByVal Shift As Integer = 1) As Long
'    MakeOnBits
'    If (Value And (2 ^ (31 - Shift))) Then 'GoTo OverFlow
'        lshift = ((Value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
'    Else
'        lshift = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))
'    End If
'End Function
'
'Public Function rshift(ByVal Value As Long, Optional ByVal Shift As Integer = 1) As Long
'    Dim hi As Long
'    MakeOnBits
'    If (Value And &H80000000) Then hi = &H40000000
'    rshift = (Value And &H7FFFFFFE) \ (2 ^ Shift)
'    rshift = (rshift Or (hi \ (2 ^ (Shift - 1))))
'End Function
'
'Private Sub MakeOnBits()
'    Dim j As Integer, v As Long
'
'    For j = 0 To 30
'        v = v + (2 ^ j)
'        OnBits(j) = v
'    Next j
'
'    OnBits(j) = v + &H80000000
'
'End Sub

Public Function inc(ByRef x, Optional Increment As Long = 1)
    x = x + Increment
End Function

Public Function dec(ByRef x, Optional Increment As Long = 1)
    x = x - Increment
End Function

Function paramsToArray(ParamArray values()) As Variant()
    paramsToArray = values
End Function

Function max(values() As Variant)
   On Error Resume Next
   Dim Item
   For Each Item In values
      max = IIf(max < Item, Item, max)
   Next
End Function

Function Min(values() As Variant)
   Dim Item
   On Error Resume Next
   Min = &H7FFFFFFF
   For Each Item In values
      Min = IIf(Min > Item, Item, Min)
   Next
End Function

Function md5(fileByteOrString) As String
    Dim b() As Byte
    b() = LoadData(fileByteOrString)
    md5 = hash.HashBytes(b)
End Function

Function ExtractHexFromDump(dump) As String
    
    On Error Resume Next
    Dim Y() As String, tmp() As String, i As Long
    Dim dat As String
    Dim a As Long
    Dim x
    
    dat = trim(dump)
    If InStr(dat, " ") > 0 Then
        tmp = split(dat, vbCrLf)
        For i = 0 To UBound(tmp)
            x = trim(tmp(i))
            If InStr(1, x, "Offset", vbTextCompare) = 1 Then GoTo nextone
            If Len(x) = 0 Then GoTo nextone
            a = InStr(x, " ") '8
            If a > 0 Then 'remove offset
                x = trim(Mid(x, a))
            End If
            
            a = Len(x)
            If a >= 48 Then
                x = trim(Mid(x, 1, 48))
            Else
                'they didnt copy the ascii area for last line? it will be copied anyway...
                'If i = UBound(tmp) Then
            End If
            
            
            push Y, replace(x, " ", Empty)
nextone:
        Next
        dat = Join(Y, "")
    End If
    
    ExtractHexFromDump = dat
     
End Function

Function Compress(fileStringOrByte, ByRef outVal, Optional method As CompressTypes = ct_Zlib) As Boolean
    
    Dim b() As Byte
    Dim bOut() As Byte
    Dim success As Boolean
    Dim f As Long
    
    b() = LoadData(fileStringOrByte)
    
    Select Case method
        Case ct_Zlib:      success = zlib.CompressData(b, bOut)
        Case ct_LZNT1:     success = RTLCompress(b, bOut)
        Case ct_LZNT1_MAX: success = RTLCompress(b, bOut, True)
    End Select
    
    If success Then
        If TypeName(outVal) = "String" Or TypeName(outVal) = "Variant" Then
           outVal = StrConv(bOut, vbUnicode, LANG_US)
        Else
           outVal = bOut
        End If
    End If
         
    Compress = success
    
End Function

'if outval is a string or variant this will return a string. if its a bytearray() it will return a byte array
Function DeCompress(fileStringOrByte, ByRef outVal, Optional method As CompressTypes = ct_Zlib) As Boolean
    
    Dim b() As Byte
    Dim bOut() As Byte
    Dim f As Long
    Dim success As Boolean
    
    On Error Resume Next
    b() = LoadData(fileStringOrByte)
    If Err.Number <> 0 Then Exit Function
    
    Select Case method
        Case ct_Zlib:  success = zlib.UncompressData(b, bOut)
        Case ct_LZNT1:   success = RTLDeCompress(b, bOut)
        Case ct_LZNT1_MAX:   success = RTLDeCompress(b, bOut, True)
    End Select
    
    If success Then
        If TypeName(outVal) = "String" Or TypeName(outVal) = "Variant" Then
           outVal = StrConv(bOut, vbUnicode, LANG_US)
        Else
           outVal = bOut
        End If
    End If
    
    DeCompress = success
         
End Function

Function DownloadFile(url, localPath) As Boolean
    DownloadFile = (URLDownloadToFile(0&, url, localPath, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS)
End Function

Function a2c(ary) As Collection
    Dim a As New Collection, x
    For Each x In ary
        a.add x
    Next
    Set a2c = a
End Function

Function c2a(c As Collection) As String()
    Dim t() As String, f
    For Each f In c
        push t, f
    Next
    c2a = t
End Function

Function c2s(c As Collection, Optional delimiter = vbCrLf) As String
    c2s = Join(c2a(c), delimiter)
End Function

Function cjoin(a As Collection, b As Collection) As Collection
    Dim ret As New Collection, f
    For Each f In a
        ret.add f
    Next
    For Each f In b
        ret.add b
    Next
    Set cjoin = ret
End Function

