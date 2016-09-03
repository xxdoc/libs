Attribute VB_Name = "mEnumWindows"
'============mEnumWindows.bas===========
'Just to enumerate windows containing SysListView32
Option Explicit

Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function EnumChildWindows& Lib "user32" (ByVal hParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Dim lv As ListView
Dim m_ClassName As String
Dim sParent As String

Function EnumWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
  sParent = GetWndText(hWnd)
  Call EnumChildWindows(hWnd, AddressOf EnumChildWinProc, lParam)
  EnumWinProc = 1
End Function

Function EnumChildWinProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
   Dim s As String, s1 As String
   If m_ClassName <> "" Then
      If GetWndClass(hWnd) = m_ClassName Then
         s = "0x" & Right("00000000" & Hex(hWnd), 8)
         If sParent = "No caption" Then
            s1 = GetWndText(GetTopLevelParent(hWnd))
         Else
            s1 = sParent
         End If
         With lv.ListItems.Add(, , s1)
            .SubItems(1) = s
            .ToolTipText = s1
            .Tag = hWnd
         End With
      End If
   Else
      s = "0x" & Right("00000000" & Hex(hWnd), 8)
      s1 = GetWndText(GetTopLevelParent(hWnd))
      With lv.ListItems.Add(, , s1)
         .SubItems(1) = s
         .ToolTipText = s1
         .Tag = hWnd
      End With
   End If
   EnumChildWinProc = 1
End Function

Private Function GetWndClass(hWnd As Long) As String
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetClassName(hWnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No class"
  GetWndClass = sName
End Function

Private Function GetWndText(hWnd As Long) As String
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetWindowText(hWnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No caption"
  GetWndText = sName
End Function

Public Sub GetWindowList(lvw As ListView, Optional ByVal sClassName As String, Optional ByVal hWndAfter As Long)
   Set lv = lvw
   m_ClassName = sClassName
   EnumWindows AddressOf EnumWinProc, 0
End Sub

Private Function GetTopLevelParent(hWnd As Long) As Long
  Dim hwndParent As Long
  Dim hwndTmp As Long
  hwndParent = hWnd
  Do
    hwndTmp = GetParent(hwndParent)
    If hwndTmp Then hwndParent = hwndTmp
  Loop While hwndTmp
  GetTopLevelParent = hwndParent
End Function

