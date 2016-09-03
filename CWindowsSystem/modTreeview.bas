Attribute VB_Name = "modTreeview"
Option Explicit
'=======================================================================
'
'  (c) 2002  Jim White, t/as MathImagical Systems
'            Uki, NSW, Australia
'
'  "dmTreeviewFPDE", by Dr Memory
'
'    This module exports functions that provide a Foreign Process
'    Data Extraction service for Treeview controls on both NT-like and
'    Win95/98 platforms.
'
'    As the required FP memory buffers are of constant size, the calling
'    program must first call the dmTreeviewAttach function, which will
'    create the buffers. When the caller has extracted all the data it wants,
'    it should call dmTreeviewRelease to release the buffers.
'
'=======================================================================
'  CAVEAT:
'    Any VB program that dabbles with memory pointers and interaction
'    with FP's (foreign processes) risks crashing the VB IDE, the
'    target application, or even (on non-NT platforms), the whole system.
'
'    For this reason, care must be taken when modifying this code.
'    Also, the calling program should invoke the dmCrashMode function
'    in its Form_Load procedure. This will prevent crashes due to
'    GPF's, and provide a reasonably safe landing (see dmCrashLanding).
'
' !! Nevertheless, for obvious reasons, MathImagics can not be held
' !! responsible for any system damage or collateral damage caused
' !! by the use of this software.
'
'=======================================================================
'  IDE:
'    These functions do NOT use subclassing, so testing in the IDE is
'    generally quite safe.
'
'=======================================================================
'  Export Table:

Public Type dmTreeView     ' information about target ListView
   hWnd        As Long
   Class       As String
   ItemCount   As Long
   Left        As Long
   Top         As Long
   Right       As Long
   Bottom      As Long
   End Type
'
'=======================================================================
'    dmGetTreeViewInfo(Target As dmTreeView) As Boolean
'             => set Target.hWnd before calling
'                returns True if hWIndow is a TreeView
'                fills in Target properties
'=======================================================================
'    dmTreeviewScan(Target As dmTreeView)
'             => scans target TreeView, populating these node property
'                tables which the caller can access =>
'
   Public tvHandle() As Long, tvNext() As Long, tvPrev() As Long
   Public tvParent() As Long, tvChild() As Long
   Public tvText() As String, tvExpanded() As Boolean
   Public tvCount As Long
'
'=======================================================================

'
'  Notes:
'    1. Should work on VB5, VB6, and VC Treeview classes, and on all
'       Win 32-bit platforms (NT4, 2000, 95, 98, etc)
'
'    2. It is not yet known whether these functions will return data
'       from VB6 (MSCOMCTL) Treeviews running in "Virtual Treeview" mode.
'       In virtual mode the control does not store any data, it sends
'       requests back to the window owner. If the class passes data-fetch
'       requests (like TVM_GETITEMTEXT) through the same mechanism, we'll
'       get the data, but if it ignores these requests, then there is no
'       way we can get the data - it exists only within the application.
'
'=======================================================================
   
   Const MAX_TVMSTRING = 255&
   
   Dim tvWindow                As Long     ' foreign process Treeview window handle
   Dim tvProcessId             As Long     '                 Process Id
   Dim myTVitem                As TVITEM   ' TVITEM template
   Dim itemText(MAX_TVMSTRING) As Byte     ' local itemdata buffer
   Dim tvItemPointer           As Long     ' address of TVITEM   in shared mem
   Dim tvDataPointer           As Long     ' address of item data in shared mem
   Dim apiResult               As Long
   Dim zBuffer(MAX_TVMSTRING)  As Byte     ' an empty buffer used to erase shared buffer
'
'======================== Windows API
'
   Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
   Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
   Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
   Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal dwNewLong As Long) As Long
   Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal lIndex As Long, ByVal dwNewLong As Long) As Long
'
'======================== Window messaging
   Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'======================== Treeview UDT's and constants

   Private Type TVITEM
      mask         As Long      '
      hItem        As Long
      state        As Long
      stateMask    As Long
      pszText      As Long
      cchTextMax   As Long
      iImage       As Long
      iSelectImage As Long
      cChildren    As Long
      lParam       As Long
      iIntegral    As Long
      End Type

'======================== Tree view message codes
   Private Const TVM_FIRST = &H1100&
   Private Const TVM_SETBKCOLOR = TVM_FIRST + 29
   Private Const TVM_SETTEXTCOLOR = TVM_FIRST + 30
   Private Const TVM_GETBKCOLOR = TVM_FIRST + 31
   Private Const TVM_GETTEXTCOLOR = TVM_FIRST + 32
   Private Const TVM_GETINDENT = TVM_FIRST + 6
   Private Const TVM_GETITEMA = TVM_FIRST + 12
   Private Const TVM_GETNEXTITEM = TVM_FIRST + 10
   Private Const TVM_GETVISIBLECOUNT = TVM_FIRST + 16
   Private Const TVM_GETCOUNT = TVM_FIRST + 5
   
   Private Const TVM_GETIMAGELIST = TVM_FIRST + 8
   Private Const TVM_SETIMAGELIST = &H1109&
   Private Const TVSIL_NORMAL = 0
   Private Const TVSIL_STATE = 2

   Private Const TVS_HASLINES = 2&
      
   ' item mask flags
   Const TVIF_TEXT = 1&
   Const TVIF_IMAGE = &H2&
   Const TVIF_PARAM = &H4&
   Const TVIF_STATE = &H8&
   Const TVIF_HANDLE = &H10&
   Const TVIF_SELECTEDIMAGE = &H20&
   Const TVIF_CHILDREN = &H40&
   Const TVIF_INTEGRAL = &H80&
   ' GETNEXT options
   Const TVGN_ROOT = &H0&
   Const TVGN_NEXT = &H1&
   Const TVGN_PREVIOUS = &H2&
   Const TVGN_PARENT = &H3&
   Const TVGN_CHILD = &H4&
   Const TVGN_FIRSTVISIBLE = &H5&
   Const TVGN_NEXTVISIBLE = &H6&
   Const TVGN_PREVIOUSVISIBLE = &H7&
   Const TVGN_DROPHILITE = &H8&
   Const TVGN_CARET = &H9&
   Const TVGN_LASTVISIBLE = &HA&
   ' State flags
   Const TVIS_EXPANDED = &H20

Dim tvIndex As Long

  
   Private fpHandle      As Long
 
   Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
   Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
   Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
   Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
   Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)
   Private Declare Function lstrlenA Lib "kernel32" (ByVal lpsz As Long) As Long
   Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
   Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
   Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   
   
   Const PAGE_READWRITE = &H4
   Const MEM_RESERVE = &H2000&
   Const MEM_RELEASE = &H8000&
   Const MEM_COMMIT = &H1000&
   Const PROCESS_VM_OPERATION = &H8
   Const PROCESS_VM_READ = &H10
   Const PROCESS_VM_WRITE = &H20
   Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Const SECTION_QUERY = &H1
   Const SECTION_MAP_WRITE = &H2
   Const SECTION_MAP_READ = &H4
   Const SECTION_MAP_EXECUTE = &H8
   Const SECTION_EXTEND_SIZE = &H10
   Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
   Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Private tvEntries As Collection
Private myTreeView As TreeView
Private tv As Boolean

Public Sub CopyTargetTreeview(TargetTreeview As dmTreeView, myTv As TreeView, entries As Collection)

   Dim item&, nItems&, iText$, iKey$
    
   Set tvEntries = entries
   Set myTreeView = myTv
   
   tv = Not myTv Is Nothing
   
   dmTreeviewScan TargetTreeview    ' take the snapshot of the target tree
   
   '===============================================================================================
   ' 1. the tables tvNext, tvPrev, tvChild, tvParent now correspond to Node.Next, Node.Prev etc
   ' 2. the first root-level node is item 1
   ' 3. the text for each item is in tvText(item)
   '
   ' To build a copy of the tree we traverse the tv-tables just as we would a TreeView.Nodes collection
   '===============================================================================================
   
   nItems = tvCount
   If tv Then myTreeView.Nodes.Clear
   
   item = 1
   While item > 0
      iText = tvText(item)
      If iText = "" Then iText = "<empty>"
      iKey = "N" & item
      If tv Then myTreeView.Nodes.Add , , iKey, iText
      tvEntries.Add iText
      CopySubtree item
      If tv And tvExpanded(item) Then myTreeView.Nodes(iKey).Expanded = True
      item = tvNext(item)
      tvEntries.Add Empty
      Wend
   Exit Sub
   
   Set tvEntries = Nothing
   Set myTreeView = Nothing

End Sub
   
Private Sub CopySubtree(ByVal pItem As Long, Optional depth As Long = 1)
   
   Dim pKey As String, sKey As String   ' parent and sibling keys
   Dim firstchild As Boolean
   Dim item As Long, iKey As String, iText As String
   
   item = tvChild(pItem)
   If item = 0 Then Exit Sub  ' childless
   
   firstchild = True
   pKey = "N" & pItem
      
   While item <> 0
      iText = tvText(item): If iText = "" Then iText = "<empty>"
      iKey = "N" & item
      If firstchild Then
         If tv Then myTreeView.Nodes.Add pKey, tvwChild, iKey, iText
         tvEntries.Add String(depth, vbTab) & iText
      Else
         If tv Then myTreeView.Nodes.Add sKey, tvwNext, iKey, iText
         tvEntries.Add String(depth, vbTab) & iText
         End If
      firstchild = False
      sKey = iKey
      CopySubtree item, (depth + 1)
      If tv And tvExpanded(item) Then myTreeView.Nodes(iKey).Expanded = True
      item = tvNext(item)
   Wend

End Sub





Public Function dmGetTreeviewInfo(Target As dmTreeView) As Boolean
   '
   ' Fills in info about target window (passed via Target.hWnd)
   ' Returns True if hWindow is a handle to a TreeView control
   '
   
   With Target
      tvWindow = .hWnd
      If IsWindow(tvWindow) = 0 Then
         tvWindow = 0
         Exit Function
         End If
      .Class = dmWindowClass(tvWindow)
      
      If InStr(1, .Class, "treeview", vbTextCompare) = 0 Then
         .hWnd = 0
         .ItemCount = 0
         tvWindow = 0
         Exit Function
         End If
      
      .ItemCount = SendMessage(tvWindow, TVM_GETCOUNT, 0&, 0&)
      GetWindowRect tvWindow, VarPtr(.Left)
      dmGetTreeviewInfo = True
      End With
    End Function

Public Sub dmTreeviewScan(Target As dmTreeView)
   '=========================================================================
   ' Unlike the ListView, we can't just iterate, we have to traverse the tree
   ' and build an equivalent. This is quite easy as we have the TVM_GETNEXTITEM
   ' message, which requires no FP buffers, and does the equivalent of the
   ' Node.Root, Node.Child, Node.Parent, and Node.Next functions
   '==========================================================================
   '
   ' perform full scan of the FP tree - NB => target might be inserting/removing items
   '
   '
   ' First,allocate cross-process data pipe
   '    One for the TVITEM structure, and one for the item's data
   '
   GetWindowThreadProcessId tvWindow, tvProcessId
   tvItemPointer = dmMemAllocate(Len(myTVitem), tvProcessId)
   tvDataPointer = dmMemAllocate(MAX_TVMSTRING, tvProcessId)
   
   
   Dim hItem As Long, nIndex As Long
   Target.ItemCount = SendMessage(tvWindow, TVM_GETCOUNT, 0&, 0&)
   tvCount = Target.ItemCount
   tvIndex = 1
   ReDim tvHandle(tvCount), tvParent(tvCount), tvChild(tvCount)
   ReDim tvNext(tvCount), tvPrev(tvCount), tvText(tvCount), tvExpanded(tvCount)
   If tvCount = 0 Then Exit Sub
   '
   ' get base node (1st Root node)
   '
   hItem = SendMessage(tvWindow, TVM_GETNEXTITEM, TVGN_ROOT, 0)
   While hItem <> 0
      nIndex = tvIndex
      tvHandle(nIndex) = hItem
      tvText(nIndex) = TreeviewItem(hItem)
      tvExpanded(nIndex) = TreeviewItemExpanded(hItem)  ' for the demo program
      TreeviewScan nIndex ' scan the subtree
      hItem = SendMessage(tvWindow, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
      If hItem <> 0 Then
         tvIndex = tvIndex + 1
         tvNext(nIndex) = tvIndex
         tvPrev(tvIndex) = nIndex
         End If
      Wend
   ' release pipe buffer   s
   dmMemRelease tvItemPointer
   dmMemRelease tvDataPointer
   tvWindow = 0
   tvProcessId = 0
   End Sub

Private Sub TreeviewScan(ByVal pIndex As Long)
   Dim hItem As Long, hParent As Long, nIndex As Long
   '
   ' this is exactly like a TreeView.Nodes traversal except we use
   ' GETNEXTITEM calls with various options, e.g. TVGN_CHILD, TVGN_NEXT, etc
   ' instead of Node.Child, Node.Next properties
   '
   hParent = tvHandle(pIndex)
   hItem = SendMessage(tvWindow, TVM_GETNEXTITEM, TVGN_CHILD, hParent)
   If hItem = 0 Then Exit Sub  ' childless
   
   tvIndex = tvIndex + 1
   tvChild(pIndex) = tvIndex
   
   While hItem <> 0
      nIndex = tvIndex
      tvHandle(nIndex) = hItem
      tvParent(nIndex) = pIndex
      tvText(nIndex) = TreeviewItem(hItem)
      tvExpanded(nIndex) = TreeviewItemExpanded(hItem)  ' for the demo program
      TreeviewScan nIndex  ' subtree scan
      hItem = SendMessage(tvWindow, TVM_GETNEXTITEM, TVGN_NEXT, hItem)
      If hItem <> 0 Then
         tvIndex = tvIndex + 1
         tvNext(nIndex) = tvIndex
         tvPrev(tvIndex) = nIndex
         End If
      Wend
   End Sub

Private Function TreeviewItem(ByVal hItem As Long) As String
   If tvWindow = 0 Then Exit Function
   
   '
   ' 1. Fill in TVITEM in the normal fashion, with just one difference
   '
   With myTVitem
      .mask = TVIF_TEXT
      .hItem = hItem
      .pszText = tvDataPointer      ' ItemData address is our shared buffer!
      .cchTextMax = MAX_TVMSTRING
      End With
   '
   ' 2. Copy the TVITEM to the shared buffer, and send the GETITEM request
   '
   dmWriteProcessData tvItemPointer, VarPtr(myTVitem), Len(myTVitem)
   apiResult = SendMessage(tvWindow, TVM_GETITEMA, 0, tvItemPointer)
   '
   ' 3. Get the return data from the shared buffer (and convert to VB string)
   '    TVM_GETITEM doesn't return the item's text length
   '
   
   If apiResult <> 0 Then
      Dim myBuffer As Long, zeroes As Long
      Erase itemText
      myBuffer = VarPtr(itemText(0))
      zeroes = VarPtr(zBuffer(0))
      dmReadProcessData tvDataPointer, myBuffer, MAX_TVMSTRING
      TreeviewItem = dmGetStringA(myBuffer)
      End If
   End Function

Private Function TreeviewItemExpanded(ByVal hItem As Long) As Boolean
   '
   ' example of getting an Item's State flags. I could get it along with the text
   ' but this is really only needed for the demo program, so it can display the
   ' tree in exactly the same state as the target, so I do it separately
   '
   If tvWindow = 0 Then Exit Function
   
   '
   ' 1. Fill in TVITEM in the normal fashion, with just one difference
   '
   With myTVitem
      .mask = TVIF_STATE
      .hItem = hItem
      .stateMask = TVIS_EXPANDED
      End With
   '
   ' 2. Copy the TVITEM to the shared buffer, and send the GETITEM request
   '
   dmWriteProcessData tvItemPointer, VarPtr(myTVitem), Len(myTVitem)
   apiResult = SendMessage(tvWindow, TVM_GETITEMA, 0, tvItemPointer)
   dmReadProcessData tvItemPointer, VarPtr(myTVitem), Len(myTVitem)
   '
   ' 3. Check my mask state bit
   '
   TreeviewItemExpanded = ((myTVitem.state And TVIS_EXPANDED) <> 0)
   End Function


Public Sub dmSetTreeViewColor(tv As TreeView, ByVal tColor&)
   Dim lngStyle&, Cwindow&
   Const GWL_STYLE = -16

   Cwindow = tv.hWnd
   Call SendMessage(Cwindow, TVM_SETBKCOLOR, 0&, ByVal tColor)
   ' Now reset the style so that the tree lines appear properly
   lngStyle = GetWindowLong(Cwindow, GWL_STYLE)
   Call SetWindowLong(Cwindow, GWL_STYLE, lngStyle And Not TVS_HASLINES)
   Call SetWindowLong(Cwindow, GWL_STYLE, lngStyle Or TVS_HASLINES)
   Exit Sub
   End Sub



'----------------------------------------------------------------------------
'    doctormemory.bas library functions below...
'----------------------------------------------------------------------------

Private Function dmMemAllocate(ByVal nBytes As Long, ByVal fpID As Long) As Long
    dmMemAllocate = VirtualAllocNT(fpID, nBytes)
End Function

Private Sub dmMemRelease(mPointer As Long)
   VirtualFreeNT mPointer
   mPointer = 0
End Sub
   
Private Sub dmReadProcessData(ByVal pBuffer As Long, ByVal pData As Long, ByVal nBytes As Long)
      ReadProcessMemory fpHandle, pBuffer, pData, nBytes, 0
End Sub

Private Sub dmWriteProcessData(ByVal pBuffer As Long, ByVal pData As Long, ByVal nBytes As Long)
      WriteProcessMemory fpHandle, pBuffer, pData, nBytes, 0
End Sub

Private Function VirtualAllocNT(ByVal fpID As Long, ByVal memSize As Long) As Long
   fpHandle = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, fpID)
   VirtualAllocNT = VirtualAllocEx(fpHandle, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
End Function

Private Sub VirtualFreeNT(ByVal MemAddress As Long)
   Call VirtualFreeEx(fpHandle, ByVal MemAddress, 0&, MEM_RELEASE)
   CloseHandle fpHandle
End Sub

Private Function dmWindowClass(ByVal hWindow As Long) As String
   Dim className As String, cLen As Long
   className = String(64, 0)
   cLen = GetClassName(hWindow, className, 63)
   If cLen > 0 Then className = Left(className, cLen)
   dmWindowClass = className
End Function

Private Function dmGetStringA(ByVal lpszA As Long) As String
   ' if lpszA is a pointer to ANSI null-terminated string
   ' this will fetch it as a VB string (BSTR)
   Dim sBuf As String, sLen As Long
   sLen = lstrlenA(lpszA)        'get length of string (in chars)
   sBuf = String$(sLen + 2, 0)   'make a buffer to copy to
   CopyMemory StrPtr(sBuf), lpszA, sLen
   dmGetStringA = dmTrimSZ(StrConv(sBuf, vbUnicode))
End Function

Private Function dmTrimSZ(sName As String) As String
   ' Keep left portion of string sName up to first 0, useful with Win API
   ' null-terminated strings whose length we might not know
   Dim X As Integer
   X = InStr(sName, Chr$(0))
   If X > 0 Then dmTrimSZ = Left$(sName, X - 1) Else dmTrimSZ = sName
End Function



