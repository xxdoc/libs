Attribute VB_Name = "Module1"
Option Explicit
'copyright David Zimmer <dzzie@yahoo.com> 2001

Public Type POINTAPI
        x As Long
        y As Long
End Type

Type LongWords
    LoWord As String
    HiWord As String
End Type

Private Type FLASHWINFO
    cbSize As Long
    hwnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type
 
Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Public Const WM_GETMINMAXINFO = &H24

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) As Long


Public Const SW_SHOWNORMAL = 1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40

Private Const GWL_STYLE = (-16)
Private Const SM_CXHSCROLL = 21
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000

Public Const WM_LBUTTONUP = &H202

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Function DisplayMessage(Msg As String, Optional f As Form)
     If Not f Is Nothing Then
        f.Caption = Msg
     End If
     Call FlashWindow
End Function

Sub FlashWindow(Optional hwnd = 0, Optional xTimes = 2, Optional TimeOut = 200)
    Dim f As FLASHWINFO
    f.cbSize = Len(f)
    f.dwFlags = 5 'caption and timer
    f.dwTimeout = TimeOut
    f.hwnd = IIf(hwnd < 1, GetForegroundWindow(), hwnd)
    f.uCount = xTimes
    FlashWindowEx f
End Sub

Function GetMySetting(key, def)
    GetMySetting = GetSetting(App.Title, "General", key, def)
End Function
   
Sub SaveMySetting(key, Value)
    SaveSetting App.Title, "General", key, Value
End Sub

Sub SaveFormPosition(frm As Form, Optional AndSize As Boolean = True, Optional AndPostion As Boolean = True)
    If frm.Tag = Empty Then MsgBox "You forgot to set my tag, caption=" & frm.Caption: Exit Sub
    If frm.WindowState <> vbMinimized Then
        If AndPostion Then
            SaveSetting App.Title, frm.Tag, "MainLeft", frm.left
            SaveSetting App.Title, frm.Tag, "MainTop", frm.top
        End If
        If Not AndSize Then Exit Sub
        SaveSetting App.Title, frm.Tag, "MainWidth", frm.Width
        SaveSetting App.Title, frm.Tag, "MainHeight", frm.Height
    End If
End Sub

Sub RestoreSavedFormPosition(frm As Form, Optional AndSize As Boolean = True, Optional AndPostion As Boolean = True)
 Dim x As Long
 On Error Resume Next
 If frm.Tag = Empty Then MsgBox "You forgot to set my tag, caption=" & frm.Caption: Exit Sub
 If AndPostion Then
    x = GetSetting(App.Title, frm.Tag, "MainLeft", 0)
    If x <> 0 Then frm.left = x
    x = GetSetting(App.Title, frm.Tag, "MainTop", 0)
    If x <> 0 Then frm.top = x
 End If
 If AndSize Then
    x = GetSetting(App.Title, frm.Tag, "MainWidth", 0)
    If x <> 0 Then frm.Width = x
    x = GetSetting(App.Title, frm.Tag, "MainHeight", 0)
    If x <> 0 Then frm.Height = x
 End If
End Sub

Public Function InIDE() As Boolean
    Debug.Assert Not TestIDE(InIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
    Test = True
End Function

Sub ShowRtClkMenu(f As Form, t As Object, m As Menu)
        LockWindowUpdate t.hwnd
        t.Enabled = False
        DoEvents
        f.PopupMenu m
        t.Enabled = True
        LockWindowUpdate 0&
End Sub

Sub SetWindowTopMost(f As Object)
   SetWindowPos f.hwnd, HWND_TOPMOST, f.left / 15, _
        f.top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub


Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
   Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Sub EnableTimer(t As Timer, Interval As Long)
    t.Interval = Interval
    t.Enabled = True
End Sub

Sub ResetTimer(t As Timer)
    t.Tag = "Resetting"
    t.Enabled = False
    t.Enabled = True
    t.Tag = ""
End Sub


Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Function LongToWords(l As Long) As LongWords
    Dim w(3) As Byte
    CopyMemory w(0), l, 4
    LongToWords.HiWord = Hexit(w(3)) & Hexit(w(2))
    LongToWords.LoWord = Hexit(w(1)) & Hexit(w(0))
End Function

Private Function Hexit(x) As String
    Hexit = Hex(x)
    If Len(Hexit) < 2 Then Hexit = "0" & Hexit
End Function

Function HiWord(l As Long) As Integer
    HiWord = CInt("&h" & LongToWords(l).HiWord)
End Function

Function LoWord(l As Long) As Integer
    LoWord = CInt("&h" & LongToWords(l).LoWord)
End Function

Function MakeLong(ByVal HiWord As Integer, ByVal LoWord As Integer) As Long
      Call CopyMemory(MakeLong, LoWord, 2)
      Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), HiWord, 2)
End Function

Function ListviewHasHScroll(lv As ListView) As Boolean
    Dim s As ScrollBarConstants
    
    s = VisibleScrollBars(lv)
    If s = vbBoth Or s = vbHorizontal Then
        ListviewHasHScroll = True
    End If
    
End Function

Function VisibleScrollBars(ControlName As Control) As ScrollBarConstants
   Dim MyStyle As Long

   MyStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)

   'Use a bitwise comparison
   If (MyStyle And (WS_VSCROLL Or WS_HSCROLL)) = _
      (WS_VSCROLL Or WS_HSCROLL) Then
      VisibleScrollBars = vbBoth
   ElseIf (MyStyle And WS_VSCROLL) = WS_VSCROLL Then
      VisibleScrollBars = vbVertical
   ElseIf (MyStyle And WS_HSCROLL) = WS_HSCROLL Then
      VisibleScrollBars = vbHorizontal
   Else
      VisibleScrollBars = vbSBNone
   End If
   
End Function

Function ScrollBarHeightTwips() As Long
    ScrollBarHeightTwips = GetSystemMetrics(SM_CXHSCROLL) * 15  'to twips
End Function

Sub UnloadIntellisense()
    On Error Resume Next
    Dim f
    For Each f In VB.Forms
        If f Is frmIntellisense Then Unload f
    Next
End Sub

Public Function DownloadFile(url, LocalFilename) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, CStr(url), CStr(LocalFilename), 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Function AryIndexExists(ary, index) As Boolean
    On Error GoTo oops
    Dim i
    i = ary(index) '<-non existant index will throw Error
    AryIndexExists = True
    Exit Function
oops:     AryIndexExists = False
End Function


