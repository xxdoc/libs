 Attribute VB_Name = "mLVDuplicate"
Option Explicit

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

Public Function LV_Duplicate(ByVal hLV As Long, lv As ListView, _
                             Optional ByVal ImageList_Normal As ImageList, _
                             Optional ByVal ImageList_Small As ImageList, _
                             Optional ByVal ImageList_Header As ImageList) As Long
   
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
   
   
   lv.ListItems.Clear
   Set lv.Icons = Nothing
   Set lv.SmallIcons = Nothing
   
   bkColor = SendMessage(hLV, LVM_FIRST, 0, ByVal 0)
   If bkColor = CLR_NONE Then bkColor = vbWindowBackground
   lv.BackColor = bkColor
   
   If GetWindowLong(hLV, GWL_STYLE) And LVS_OWNERDRAWFIXED Then
      lv.ColumnHeaders.Clear
      lv.View = lvwReport
      lv.HideColumnHeaders = True
      lv.ColumnHeaders.Add , , , 4500
'      lv.ListItems.Add , , "This listview has ownerdraw style"
'      Exit Function
   End If
   
   If Not EnableDebugPrivNT Then Exit Function
   tid = GetWindowThreadProcessId(hLV, pid)
   If pid = 0 Then Exit Function
   hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
   If hProcess = 0 Then Exit Function
   
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
   
   Call LVHeaders_Duplicate(hLV, lv, hProcess, ImageList_Header)
   hIml_large = SendMessage(hLV, LVM_GETIMAGELIST, LVSIL_NORMAL, ByVal 0&)
   hIml_small = SendMessage(hLV, LVM_GETIMAGELIST, LVSIL_SMALL, ByVal 0&)
     
   If hIml_large Then
      If Not ImageList_Normal Is Nothing Then
         nIcons = IL_Duplicate(hProcess, hIml_large, ImageList_Normal)
         ImageList_Normal.BackColor = bkColor
         Set lv.Icons = ImageList_Normal
      End If
   End If
   
   If hIml_small Then
      If Not ImageList_Small Is Nothing Then
         nSmallIcons = IL_Duplicate(hProcess, hIml_small, ImageList_Small)
         ImageList_Small.BackColor = bkColor
         Set lv.SmallIcons = ImageList_Small
      End If
   Else
      If nIcons Then
         nSmallIcons = nIcons
         Set lv.SmallIcons = ImageList_Normal
      End If
   End If
   If Not bIcons Then bIcons = CBool(nIcons) Or CBool(nSmallIcons)
   nCount = SendMessage(hLV, LVM_GETITEMCOUNT, 0, ByVal 0&)
   If nCount = 0 Then
      lv.ListItems.Add , , "ListView is empty"
      GoTo CleanUp
   End If
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
        
       If nIcons Then
          Set itm = lv.ListItems.Add(, , TrimNull(it.pszText), IIf(li.iImage > nIcons, 0, li.iImage + 1), IIf(li.iImage > nSmallIcons, 0, li.iImage + 1))
       Else
          If nSmallIcons Then
             Set itm = lv.ListItems.Add(, , TrimNull(it.pszText), , IIf(li.iImage > nSmallIcons, 0, li.iImage + 1))
          Else
             Set itm = lv.ListItems.Add(, , TrimNull(it.pszText))
          End If
       End If
       With itm
          li.mask = LVIF_TEXT
          For j = 1 To lv.ColumnHeaders.Count - 1
              li.iSubItem = j
              WriteProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
              WriteProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
              Call SendMessage(hLV, LVM_GETITEM, 0, ByVal liAddr)
              ReadProcessMemory hProcess, ByVal liAddr, li, Len(li), lWritten
              ReadProcessMemory hProcess, ByVal itAddr, it, Len(it), lWritten
              .SubItems(j) = TrimNull(it.pszText)
          Next j
       End With
   Next i
   VirtualFreeEx hProcess, ByVal liAddr, 0, MEM_RELEASE
   VirtualFreeEx hProcess, ByVal itAddr, 0, MEM_RELEASE
   If (nIcons = 0) And (nSmallIcons = 0) And bIcons And nCount Then
      MsgBox "Damn! This stupid VB ListView still remember that long time ago it had some icons!"
   End If
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