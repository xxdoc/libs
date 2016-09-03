Attribute VB_Name = "Module2"
'===========mHeaderDuplicate.bas========
'module to duplicate LV column headers
Option Explicit

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

Private Type ITEM_TEXT
   pszText As String * 80
End Type

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

Private Function IsHeaderVisible(ByVal hHDR As Long) As Boolean
   IsHeaderVisible = Not ((GetWindowLong(hHDR, GWL_STYLE) And HDS_HIDDEN) = HDS_HIDDEN)
End Function

Public Function LVHeaders_Duplicate(ByVal hLV As Long, lv As ListView, _
                                    Optional ByVal hProcess As Long, _
                                    Optional ByVal ImageList_Header As ImageList) As Long

   Dim tid As Long, pid As Long
   Dim hHDR As Long, nCount As Long, i As Long
   Dim hiAddr As Long, itAddr As Long, lWritten As Long, hIml As Long
   
   Dim bNeedClose As Boolean
   Dim hi As HD_ITEM
   Dim it As ITEM_TEXT
   
   lv.ColumnHeaders.Clear
   Set lv.ColumnHeaderIcons = Nothing
   If hProcess = 0 Then
      bNeedClose = True
      tid = GetWindowThreadProcessId(hLV, pid)
      EnableDebugPrivNT
      hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
      If hProcess = 0 Then Exit Function
   End If
   hHDR = SendMessage(hLV, LVM_GETHEADER, 0, ByVal 0&)
   If hHDR Then
      hIml = SendMessage(hHDR, HDM_GETIMAGELIST, 0, ByVal 0&)
      If hIml Then
         If IL_Duplicate(hProcess, hIml, ImageList_Header) Then
            Set lv.ColumnHeaderIcons = ImageList_Header
         End If
      End If
      nCount = SendMessage(hHDR, HDM_GETITEMCOUNT, 0, ByVal 0&)
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
          With lv.ColumnHeaders.Add(, , TrimNull(it.pszText), hi.cxy * Screen.TwipsPerPixelX, hi.fmt And 3)
              If Not lv.ColumnHeaderIcons Is Nothing Then
                 .Icon = hi.iImage 'IIf(hi.iImage > lv.ColumnHeaderIcons.ListImages.Count, 0,hi.iImage)
              End If
              .Tag = hi.iOrder + 1 'store header item position for reodering columns
          End With
      Next i
      For i = 1 To nCount
          With lv.ColumnHeaders(i)
             .Position = .Tag 'move headers
          End With
      Next i
      VirtualFreeEx hProcess, ByVal hiAddr, 0, MEM_RELEASE
      VirtualFreeEx hProcess, ByVal itAddr, 0, MEM_RELEASE
      lv.HideColumnHeaders = Not IsHeaderVisible(hHDR)
   End If
CleanUp:
   If bNeedClose Then CloseHandle hProcess
End Function

