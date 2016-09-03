Attribute VB_Name = "mLVDuplicate"
Option Explicit
'*****************************************************************
' Main module to duplicate remote ListView using remote API call.
'
' Written by Arkadiy Olovyannikov (ark@msun.ru)
' Copyright 2005 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'*****************************************************************
'
'Note: I have stripped the original capability to extract and replicate icons since it had
'      problems with DEP and I really only care about extracting text data..
'      I have also combined 3-4 of Ark's original modules in this one for simplicity - dz 9.2.16
'
'      Great code and thanks for releasing open source!!

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Private Type ITEM_TEXT
   pszText As String * 80
End Type

Public Const LVM_FIRST = &H1000
Private Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Private Const LVM_GETITEM = (LVM_FIRST + 5)
Private Const LVM_SETITEM = (LVM_FIRST + 6)
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Private Const LVM_SETVIEW = (LVM_FIRST + 142)
Private Const LVM_GETVIEW = (LVM_FIRST + 143)

Private Const LVIF_TEXT = &H1
Private Const LVIF_IMAGE = &H2
Private Const LVIF_PARAM = &H4
Private Const LVIF_STATE = &H8
Private Const LVIF_INDENT = &H10
Private Const LVIF_ALL = LVIF_TEXT Or LVIF_IMAGE Or LVIF_PARAM Or LVIF_STATE Or LVIF_INDENT

Private Const LVIS_SELECTED = &H2

Private Const LVS_ICON = &H0
Private Const LVS_REPORT = &H1
Private Const LVS_SMALLICON = &H2
Private Const LVS_LIST = &H3
Private Const LVS_TILE = &H4
Private Const LVS_TYPEMASK = &H3
Private Const LVS_SHAREIMAGELISTS = &H40
Private Const LVS_OWNERDRAWFIXED = &H400

Private Const LVS_ALIGNTOP = &H0
Private Const LVS_AUTOARRANGE = &H100
Private Const LVS_ALIGNLEFT = &H800
Private Const LVS_ALIGNMASK = &HC00

Private Const LVSIL_NORMAL = 0
Private Const LVSIL_SMALL = 1
Private Const LVSIL_STATE = 2
Private Const CLR_NONE = -1

Private Type HD_ITEM
    mask As Long
    cxy As Long
    pszText As Long
    hbm As Long
    cchTextMax As Long
    fmt As Long
    lParam As Long
    ' 4.70:
    iImage As Long
    iOrder As Long
End Type

'Private Type ITEM_TEXT
'   pszText As String * 80
'End Type

Private Const HDM_FIRST = &H1200
Private Const HDM_GETITEMCOUNT = HDM_FIRST + 0
Private Const HDM_GETITEMA = HDM_FIRST + 3
Private Const HDM_GETIMAGELIST = (HDM_FIRST + 9)

Private Const HDI_WIDTH = &H1
Private Const HDI_HEIGHT = HDI_WIDTH
Private Const HDI_TEXT = &H2
Private Const HDI_FORMAT = &H4
Private Const HDI_LPARAM = &H8
Private Const HDI_BITMAP = &H10
Private Const HDI_IMAGE = &H20
Private Const HDI_ORDER = &H80
Private Const HDI_ALL = HDI_WIDTH Or HDI_TEXT Or HDI_FORMAT Or HDI_BITMAP Or HDI_IMAGE Or HDI_ORDER

Private Const HDF_LEFT = &H0
Private Const HDF_RIGHT = &H1
Private Const HDF_CENTER = &H2
Private Const HDF_JUSTIFYMASK = &H3

Private Const HDF_IMAGE = &H800
Private Const HDF_BITMAP_ON_RIGHT = &H1000
Private Const HDF_BITMAP = &H2000
Private Const HDF_STRING = &H4000
Private Const HDF_OWNERDRAW = &H8000

Private Const HDS_HIDDEN = &H8
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal nStyle As Long)
Private Const GWL_STYLE = (-16)


Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RELEASE = &H8000

Private Const PAGE_READWRITE = &H4&
Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const ANYSIZE_ARRAY = 1
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8

Private Type LARGE_INTEGER
  LowPart As Long
  HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
  pLuid As LARGE_INTEGER
  Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
  ByVal ProcessHandle As Long, _
  ByVal DesiredAccess As Long, _
  TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
  Alias "LookupPrivilegeValueA" ( _
  ByVal lpSystemName As String, _
  ByVal lpName As String, _
  lpLuid As LARGE_INTEGER) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" ( _
  ByVal TokenHandle As Long, _
  ByVal DisableAllPrivileges As Long, _
  ByRef NewState As TOKEN_PRIVILEGES, _
  ByVal BufferLength As Long, _
  ByRef PreviousState As Any, _
  ByRef ReturnLength As Any) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private bDone As Boolean


Dim bLv As Boolean 'are we replicating the listview to our control or just copying contents..
Dim headerCount As Long
Dim c As Collection
Dim g_colSeperator As String

Public Function LV_Duplicate(ByVal hLV As Long, Optional lv As ListView, Optional colSeperator As String = vbTab) As Collection
   
   Dim tid As Long, pid As Long, hProcess As Long
   Dim nCount As Long, i As Long, j As Long
   Dim liAddr As Long, itAddr As Long, lWritten As Long, align As Long, bkColor As Long
   Dim hIml_small As Long
   Dim hIml_large As Long
   Dim nSmallIcons As Long, nIcons As Long
   Dim itm As ListItem
   Static bIcons As Boolean
   Dim li As LV_ITEM
   Dim it As ITEM_TEXT
   
   Dim tmp() As String
   
   headerCount = 0
   g_colSeperator = colSeperator
   bLv = Not lv Is Nothing
   Set c = New Collection
   Set LV_Duplicate = c
      
   If bLv Then lv.ListItems.Clear
   If GetWindowLong(hLV, GWL_STYLE) And LVS_OWNERDRAWFIXED Then Exit Function
   
   'If Not EnableDebugPrivNT Then Exit Function
   
   tid = GetWindowThreadProcessId(hLV, pid)
   If pid = 0 Then Exit Function
   
   hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
   If hProcess = 0 Then Exit Function
   
   If bLv Then
   
        Select Case GetLVViewStyle(hLV)
           Case LVS_ICON:      lv.View = lvwIcon
           Case LVS_REPORT:    lv.View = lvwReport
           Case LVS_SMALLICON: lv.View = lvwSmallIcon
           Case LVS_LIST:      lv.View = lvwList
           Case Else:          lv.View = lvwIcon
        End Select
   
        align = GetWindowLong(hLV, GWL_STYLE) And LVS_ALIGNMASK
        If align Then
           If align And LVS_ALIGNLEFT Then
              lv.Arrange = lvwAutoLeft
           Else
              lv.Arrange = lvwAutoTop
           End If
        Else
           lv.Arrange = lvwNone
        End If
        
   End If
   
   Call LVHeaders_Duplicate(hLV, lv, hProcess)
   
   nCount = SendMessage(hLV, LVM_GETITEMCOUNT, 0, ByVal 0&)
   If nCount = 0 Then GoTo CleanUp 'empty nothing to do...
   
   liAddr = VirtualAllocEx(hProcess, 0, Len(li), MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   itAddr = VirtualAllocEx(hProcess, 0, LenB(it), MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
   
   For i = 0 To nCount - 1
       ZeroMemory li, Len(li)
       ZeroMemory it, Len(it)
       li.cchTextMax = Len(it)
       li.mask = LVIF_ALL
       li.pszText = itAddr
       li.iItem = i
       WriteProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
       WriteProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
       Call SendMessage(hLV, LVM_GETITEM, i, ByVal liAddr)
       ReadProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
       ReadProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
        
       If bLv Then Set itm = lv.ListItems.Add(, , TrimNull(it.pszText))
       push tmp, TrimNull(it.pszText)
       
       li.mask = LVIF_TEXT
       For j = 1 To headerCount - 1
            li.iSubItem = j
            WriteProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
            WriteProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
            Call SendMessage(hLV, LVM_GETITEM, 0, ByVal liAddr)
            ReadProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
            ReadProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
            If bLv Then itm.SubItems(j) = TrimNull(it.pszText)
            push tmp, TrimNull(it.pszText)
       Next j
       
       c.Add Join(tmp, colSeperator)
       Erase tmp
       
   Next i
   
   VirtualFreeEx hProcess, ByVal liAddr, 0, MEM_RELEASE
   VirtualFreeEx hProcess, ByVal itAddr, 0, MEM_RELEASE
   
CleanUp:
   If hProcess Then CloseHandle hProcess
   
End Function

Private Function GetLVViewStyle(ByVal hLV As Long) As Long
   Dim lStyle As Long
   lStyle = GetWindowLong(hLV, GWL_STYLE) And LVS_TYPEMASK
   If lStyle = 0 Then 'Probably XP?
      lStyle = SendMessage(hLV, LVM_GETVIEW, 0, ByVal 0&)
   End If
   GetLVViewStyle = lStyle
End Function

Private Function GetLVAlign(ByVal hLV As Long) As Long
   GetLVAlign = GetWindowLong(hLV, GWL_STYLE) And LVS_ALIGNMASK
End Function


Private Function IsHeaderVisible(ByVal hHDR As Long) As Boolean
   IsHeaderVisible = Not ((GetWindowLong(hHDR, GWL_STYLE) And HDS_HIDDEN) = HDS_HIDDEN)
End Function

Private Function LVHeaders_Duplicate(ByVal hLV As Long, lv As ListView, Optional ByVal hProcess As Long) As Long

   Dim tid As Long, pid As Long
   Dim hHDR As Long, nCount As Long, i As Long
   Dim hiAddr As Long, itAddr As Long, lWritten As Long, hIml As Long
   Dim tmp() As String
   
   Dim bNeedClose As Boolean
   Dim hi As HD_ITEM
   Dim it As ITEM_TEXT
   
   If bLv Then lv.ColumnHeaders.Clear
   If bLv Then Set lv.ColumnHeaderIcons = Nothing
   
   If hProcess = 0 Then
      bNeedClose = True
      tid = GetWindowThreadProcessId(hLV, pid)
      EnableDebugPrivNT
      hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
      If hProcess = 0 Then Exit Function
   End If
   
   hHDR = SendMessage(hLV, LVM_GETHEADER, 0, ByVal 0&)
   
   If hHDR Then
      
      nCount = SendMessage(hHDR, HDM_GETITEMCOUNT, 0, ByVal 0&)
      headerCount = nCount
      If nCount = 0 Then GoTo CleanUp
      
      hiAddr = VirtualAllocEx(hProcess, 0, Len(hi), MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
      itAddr = VirtualAllocEx(hProcess, 0, LenB(it), MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
      
      For i = 0 To nCount - 1
          ZeroMemory hi, Len(hi)
          ZeroMemory it, Len(it)
          hi.cchTextMax = Len(it)
          hi.mask = HDI_ALL
          hi.pszText = itAddr
          WriteProcessMemory hProcess, ByVal hiAddr, hi, Len(hi), lWritten
          WriteProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
          Call SendMessage(hHDR, HDM_GETITEMA, i, ByVal hiAddr)
          ReadProcessMemory hProcess, ByVal hiAddr, hi, Len(hi), lWritten
          ReadProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
          
          push tmp, TrimNull(it.pszText)
          
          If bLv Then
              With lv.ColumnHeaders.Add(, , TrimNull(it.pszText), hi.cxy * Screen.TwipsPerPixelX, hi.fmt And 3)
                  .Tag = hi.iOrder + 1 'store header item position for reodering columns
              End With
          End If
          
      Next i
      
      If bLv Then
            For i = 1 To nCount
                With lv.ColumnHeaders(i)
                   .position = .Tag 'move headers
                End With
            Next i
      End If
      
      VirtualFreeEx hProcess, ByVal hiAddr, 0, MEM_RELEASE
      VirtualFreeEx hProcess, ByVal itAddr, 0, MEM_RELEASE
      If bLv Then lv.HideColumnHeaders = Not IsHeaderVisible(hHDR)
   End If
   
CleanUp:

   c.Add Join(tmp, g_colSeperator)
   If bNeedClose Then CloseHandle hProcess
End Function

Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

Public Function EnableDebugPrivNT() As Boolean
  If bDone Then
     EnableDebugPrivNT = True
     Exit Function
  End If
 
  Dim hToken As Long
  Dim li As LARGE_INTEGER
  Dim tkp As TOKEN_PRIVILEGES
 
  If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES _
                      Or TOKEN_QUERY, hToken) = 0 Then Exit Function
 
  If LookupPrivilegeValue("", SE_DEBUG_NAME, li) = 0 Then Exit Function
 
  tkp.PrivilegeCount = 1
  tkp.Privileges(0).pLuid = li
  tkp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
 
  bDone = AdjustTokenPrivileges(hToken, False, tkp, 0, ByVal 0&, 0)
  EnableDebugPrivNT = bDone
End Function

