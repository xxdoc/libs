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

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_GETCOUNT = &H18B
Private Const LB_SETTOPINDEX = &H197

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

Private OnBits(0 To 31) As Long
Private crc_table() As Long
Public base64 As New CBase64
Public hash As New CWinHash

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
    SetWindowPos frm.hwnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Public Function GetShortName(ByVal sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name !!
    If Not FileExists(sFile) Then
        iCreated = True
        writeFile sFile, ""
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
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Sub SaveMySetting(Key, value)
    SaveSetting App.EXEName, "Settings", Key, value
End Sub

Function GetMySetting(Key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", Key, defaultval)
End Function

Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
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

Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  Dim f As Long
  Dim b() As Byte
  
  f = FreeFile
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte, f As Long
    
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = True
    Exit Function
hell: writeFile = False
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

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
       
    x = UBound(ary)
    ReDim Preserve ary(x + 1)
    
    If IsObject(value) Then
        Set ary(x) = value
    Else
        ary(x) = value
    End If
    
    Exit Sub
init:
    ReDim ary(0)
    If IsObject(value) Then
        Set ary(0) = value
    Else
        ary(0) = value
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
  
    x = UBound(ary)
    AryIsEmpty = False
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

Function AnyofTheseInstr(data, match, Optional compare As VbCompareMethod = vbTextCompare) As Boolean
    Dim tmp() As String
    Dim x
    tmp = Split(match, ",")
    For Each x In tmp
        If InStr(1, data, x, compare) > 0 Then
            AnyofTheseInstr = True
            Exit Function
        End If
    Next
End Function

Public Function isHexNum(v, Optional ByRef dataOut) As Boolean
    On Error Resume Next
    dataOut = CLng("&h" & v)
    If Err.Number = 0 Then isHexNum = True
End Function

Function LB_Count(hListBox As Long) As Long
   LB_Count = SendMessage(hListBox, LB_GETCOUNT, ByVal CLng(0), ByVal CLng(0))
End Function

Public Sub LB_ScrollTo(hListBox As Long, Optional ByVal index As Long = -1)
    If index = -1 Then index = LB_Count(hListBox)
    SendMessage hListBox, LB_SETTOPINDEX, ByVal index - 1, ByVal 0
End Sub


Public Sub LV_ColumnSort(ListViewControl As Object, Column As Object)
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
       .Sorted = -1
    End With
End Sub

Function LV_GetAllElements(lv As Object) As String
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
    
    LV_GetAllElements = Join(ret, vbCrLf)
    
End Function

Function rpad(v, Optional l As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        rpad = v & String(l - x, char)
    Else
hell:
        rpad = v
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

'supports %x, %c, %s, %d, %10d \t \n %%
Function printf(ByVal msg As String, ParamArray vars()) As String

    Dim t
    Dim ret As String
    Dim i As Long, base, marker
    
    msg = Replace(msg, Chr(0), Empty)
    msg = Replace(msg, "\t", vbTab)
    msg = Replace(msg, "\n", vbCrLf) 'simplified
    msg = Replace(msg, "%%", Chr(0))
    
    t = Split(msg, "%")
    If UBound(t) <> UBound(vars) + 1 Then
        MsgBox "Format string mismatch.."
        Exit Function
    End If
    
    ret = t(0)
    For i = 1 To UBound(t)
        base = t(i)
        marker = ExtractSpecifier(base)
        If Len(marker) > 0 Then
            ret = ret & HandleMarker(base, marker, vars(i - 1))
        Else
            ret = ret & base
        End If
    Next
    
    ret = Replace(ret, Chr(0), "%")
    printf = ret
    
End Function

Private Function HandleMarker(base, ByVal marker, var) As String
    Dim newBase As String
    Dim mType As Integer
    Dim nVal As String
    Dim spacer As String
    Dim prefix As String
    Dim count As Long
    
    If Len(base) > Len(marker) Then
        newBase = Mid(base, Len(marker) + 1) 'remove the marker..
    End If
    
    mType = Asc(Mid(marker, Len(marker), 1))  'last character
    
    Select Case mType
        Case Asc("x"): nVal = Hex(var)
        Case Asc("X"): nVal = UCase(Hex(var))
        Case Asc("s"): nVal = var
        Case Asc("S"): nVal = UCase(var)
        Case Asc("c"): nVal = Chr(var)
        Case Asc("d"): nVal = var
        
        Case Else: nVal = var
    End Select
    
    If Len(marker) > 1 Then 'it has some more formatting involved..
        marker = Mid(marker, 1, Len(marker) - 1) 'trim off type
        If Left(marker, 1) = "0" Then
            spacer = "0"
            marker = Mid(marker, 2)
        Else
            spacer = " "
        End If
        count = CLng(marker) - Len(nVal)
        If count > 0 Then prefix = String(count, spacer)
    End If
    
    HandleMarker = prefix & nVal & newBase
            
End Function

Private Function ExtractSpecifier(v)
    
    Dim ret As String
    Dim b() As Byte
    Dim i As Long
    If Len(v) = 0 Then Exit Function
    
    b() = StrConv(v, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        ret = ret & Chr(b(i))
        If b(i) = Asc("x") Then Exit For
        If b(i) = Asc("X") Then Exit For
        If b(i) = Asc("c") Then Exit For
        If b(i) = Asc("s") Then Exit For
        If b(i) = Asc("S") Then Exit For
        If b(i) = Asc("d") Then Exit For
    Next
    
    ExtractSpecifier = ret
    
End Function

Public Function BuildConnectionString(dbServer As dbServers, dbName As String, Optional server As String, Optional Port = 3306, Optional user As String, Optional pass As String) As String
    Dim dbPath As String, baseString As String, blnInlineAuth As Boolean
    
    Select Case dbServer
        Case Access
            baseString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
        Case FileDsn
            baseString = "FILEDSN=____;"
        Case DSN
            baseString = "DSN=____;"
        Case dBase
            baseString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=____;"
        Case mysql
            baseString = "Driver={mySQL};Server=" & server & ";Port=" & Port & ";Stmt=;Option=16834;Database=____;"
        Case MsSql2k
            baseString = "Driver={SQL Server};Server=" & server & ";Database=____;"
        Case JetAccess2k
            baseString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=____;" & _
                         "User Id=" & user & ";" & _
                         "Password=" & pass & ";"
                         blnInlineAuth = True
    End Select
                         
        
    If Not blnInlineAuth Then
        If user <> Empty Then baseString = baseString & "Uid:" & user & ";"
        If pass <> Empty Then baseString = baseString & "Pwd:" & user & ";"
    End If
       
    '%AP% is like enviromental variable for app.path i am lazy :P
    dbPath = Replace(dbName, "%AP%", App.path)
    
    BuildConnectionString = Replace(baseString, "____", dbPath)
    
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

Function CountOccurances(it, find) As Integer
    Dim tmp() As String
    
    If InStr(1, it, find, vbTextCompare) < 1 Then Exit Function
    tmp = Split(it, find, , vbTextCompare)
    CountOccurances = UBound(tmp)
    
End Function

Function h2l(ByVal x, ByRef outVal As Long) As Boolean
    On Error Resume Next
    x = Trim(Replace(x, "0x", Empty, , , vbTextCompare))
    x = Trim(Replace(x, "&h", Empty, , , vbTextCompare))
    outVal = CLng("&h" & x)
    h2l = (Err.Number = 0)
End Function

Function HexDump(bAryOrStrData, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    Const LANG_US = &H409
    Dim i As Long, tt, h, x

    offset = 0
    
    If TypeName(bAryOrStrData) = "Byte()" Then
        ary() = bAryOrStrData
    Else
        ary = StrConv(CStr(bAryOrStrData), vbFromUnicode, LANG_US)
    End If
    
    chars = "   "
    For i = 1 To UBound(ary) + 1
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
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Public Function toBytes(ByVal hexstr, Optional strRet As Boolean = False, Optional isDecimal As Boolean = False)

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
    
    str = Replace(hexstr, vbCr, Empty)
    str = Replace(str, vbLf, Empty)
    str = Replace(str, vbTab, Empty)
    str = Replace(str, Chr(0), Empty)
    str = Replace(str, "{", Empty)
    str = Replace(str, "}", Empty)
    str = Replace(str, ";", Empty)
    str = Replace(str, "+", Empty)
    str = Replace(str, """""", Empty)
    str = Replace(str, "'", Empty)
    hexstr = str
    
    If InStr(hexstr, "\u") > 0 Then hexstr = Replace(hexstr, "\u", "%u")
    
    If InStr(hexstr, "%u") > 0 Then
        tmp = Split(hexstr, "%u")
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
        tmp = Split(hexstr, "%")
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
            tmp = Split(hexstr, marker)
            
            If isDecimal Then
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
        
    str = Replace(hexstr, " ", Empty)
    str = Replace(str, "0x", Empty)
    str = Replace(str, ",", Empty)
    
    For i = 1 To Len(str) Step 2
        x = Mid(str, i, 2)
        If Not isHexChar(x, b) Then Exit Function
        bpush r(), b
    Next
    
    If strRet Then
        toBytes = StrConv(r, vbUnicode, LANG_US)
    Else
        toBytes = r
    End If
    
hell:
End Function

Function FirstOccurance(it, ByVal csvFind As String, ByRef outFoundVal) As Long
    If Len(csvFind) = 0 Then Exit Function
    
    Dim find() As String, x, lowestOffset As Long, lowestIndex As Long, i As Long, a As Long
    
    outFoundVal = Empty
    lowestOffset = MAX_LONG
    find = Split(csvFind, ",")
    
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

Public Function rc4(data As Variant, ByVal Password As Variant) As String
On Error Resume Next
    Dim RB(0 To 255) As Integer, x As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, temp As Byte
    
    Dim plen As Long
    
    If TypeName(data) = "Byte()" Then
        ByteArray() = data
    Else
        If Len(data) = 0 Then Exit Function
        ByteArray() = StrConv(CStr(data), vbFromUnicode, LANG_US)
    End If
    
    If TypeName(Password) = "Byte()" Then
        Key() = Password
        If UBound(Key) > 255 Then ReDim Preserve Key(255)
    Else
        If Len(Password) = 0 Then
            Exit Function
        End If

        If Len(Password) > 256 Then
            Key() = StrConv(Left$(CStr(Password), 256), vbFromUnicode, LANG_US)
        Else
            Key() = StrConv(CStr(Password), vbFromUnicode, LANG_US)
        End If
    End If
    
    plen = UBound(Key) + 1
 
    'Debug.Print "key=" & HexDump(Key)
    'Debug.Print "data=" & HexDump(ByteArray)
    
    For x = 0 To 255
        RB(x) = x
    Next x
    
    x = 0
    Y = 0
    Z = 0
    For x = 0 To 255
        Y = (Y + RB(x) + Key(x Mod plen)) Mod 256
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
    
    rc4 = StrConv(ByteArray, vbUnicode, LANG_US)
    
End Function

Function b64Encode(bArrayStringOrFile, Optional retAsString As Boolean = True) As Variant
    b64Encode = base64.Encode(bArrayStringOrFile, retAsString)
End Function

Function b64Decode(bArrayStringOrFile, ByRef dataOut As Variant, Optional retAsString As Boolean = True) As Boolean
     b64Decode = base64.Decode(bArrayStringOrFile, dataOut, retAsString)
End Function

Function crc32(bAryOrString) As String
    
    Dim c As Long, n As Long, x As Long
    Dim b() As Byte
    
    c = -1
    If AryIsEmpty(crc_table) Then make_crc_table

    If TypeName(bAryOrString) = "Byte()" Then
        b() = bAryOrString
    Else
        b() = StrConv(CStr(bAryOrString), vbFromUnicode, LANG_US)
    End If
    
    For n = 0 To UBound(b)
        c = crc_table((c Xor b(n)) And &HFF) Xor RShift(c, 8)
    Next
    
    crc32 = Hex(c Xor &HFFFFFFFF)
    
End Function

Private Sub make_crc_table()
    Dim c As Long, n As Long, k As Long
    
    ReDim crc_table(256)
    
    For n = 0 To 255
          c = n
          For k = 0 To 7
                If c And 1 Then
                     c = &HEDB88320 Xor RShift(c)
                Else
                    c = RShift(c)
                End If
           Next
          crc_table(n) = c
    Next
     
End Sub

Public Function LShift(ByVal value As Long, Optional ByVal Shift As Integer = 1) As Long
    MakeOnBits
    If (value And (2 ^ (31 - Shift))) Then 'GoTo OverFlow
        LShift = ((value And OnBits(31 - (Shift + 1))) * (2 ^ (Shift))) Or &H80000000
    Else
        LShift = ((value And OnBits(31 - Shift)) * (2 ^ Shift))
    End If
End Function

Public Function RShift(ByVal value As Long, Optional ByVal Shift As Integer = 1) As Long
    Dim hi As Long
    MakeOnBits
    If (value And &H80000000) Then hi = &H40000000
    RShift = (value And &H7FFFFFFE) \ (2 ^ Shift)
    RShift = (RShift Or (hi \ (2 ^ (Shift - 1))))
End Function

Private Sub MakeOnBits()
    Dim j As Integer, v As Long
  
    For j = 0 To 30
        v = v + (2 ^ j)
        OnBits(j) = v
    Next j
  
    OnBits(j) = v + &H80000000

End Sub

Public Function inc(ByRef x, Optional Increment As Long = 1)
    x = x + Increment
End Function

Public Function dec(ByRef x, Optional Increment As Long = 1)
    x = x - Increment
End Function

Function paramsToArray(ParamArray values()) As Variant()
    paramsToArray = values
End Function

Function Max(values() As Variant)
   On Error Resume Next
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(values() As Variant)
   Dim item
   On Error Resume Next
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function md5(fileStringOrByte) As String
    
    Dim b() As Byte
    
    If FileExists(fileStringOrByte) Then
        md5 = hash.HashFile(CStr(fileStringOrByte))
    ElseIf TypeName(fileStringOrByte) = "Byte()" Then
        b() = fileStringOrByte
        md5 = hash.HashBytes(b)
    Else
        md5 = hash.HashString(CStr(fileStringOrByte))
    End If
           
End Function

Function ExtractHexFromDump(dump) As String
    
    On Error Resume Next
    Dim Y() As String, tmp() As String, i As Long
    Dim dat As String
    Dim a As Long
    Dim x
    
    dat = Trim(dump)
    If InStr(dat, " ") > 0 Then
        tmp = Split(dat, vbCrLf)
        For i = 0 To UBound(tmp)
            x = Trim(tmp(i))
            If InStr(1, x, "Offset", vbTextCompare) = 1 Then GoTo nextOne
            If Len(x) = 0 Then GoTo nextOne
            a = InStr(x, " ") '8
            If a > 0 Then 'remove offset
                x = Trim(Mid(x, a))
            End If
            
            a = Len(x)
            If a >= 48 Then
                x = Trim(Mid(x, 1, 48))
            Else
                'they didnt copy the ascii area for last line? it will be copied anyway...
                'If i = UBound(tmp) Then
            End If
            
            
            push Y, Replace(x, " ", Empty)
nextOne:
        Next
        dat = Join(Y, "")
    End If
    
    ExtractHexFromDump = dat
     
End Function


