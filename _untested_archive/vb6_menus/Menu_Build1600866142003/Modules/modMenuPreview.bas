Attribute VB_Name = "modMenuPreview"
Option Explicit
'
' Created & released by KSY, 06/14/2003
'
Private m_lpfnMenuBarOldWndProc As Long

Public Sub BeginMenuBarPreviewSubclassing(ByVal hWnd As Long)
   m_lpfnMenuBarOldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndMenuBarPreviewProc)
End Sub

Public Sub EndMenuBarPreviewSubclassing(ByVal hWnd As Long)
   Call SetWindowLong(hWnd, GWL_WNDPROC, m_lpfnMenuBarOldWndProc)
End Sub

Private Function WndMenuBarPreviewProc(ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim nCmdCode As Long, nCmdItemID As Long
   Select Case uMsg
   Case WM_COMMAND
      If lParam = 0 Then 'in case of menu, lParam is always 0.
         'HiWord of wParam is command code source.
         '0 = Menu, CommandButton, 1 = Accelerator, Other = Control
         If HiWord(wParam) = 0 Then 'if menu
            'LoWord of wParam is Menu Item ID
            nCmdItemID = LoWord(wParam) And &HFFFF& 'Fix to unsigned integer
            'Display message.
            MsgBox "The user clicked [" & GetMenuCaption(GetMenu(hWnd), nCmdItemID) & "]." & _
                                                         vbCrLf & "Menu Item ID=" & nCmdItemID
         End If
      End If
   'Case WM_NCPAINT '= &H85&
   '   DrawMenuBar GetMenu(hWnd)
   Case Else
   End Select
   WndMenuBarPreviewProc = CallWindowProc(m_lpfnMenuBarOldWndProc, hWnd, uMsg, wParam, lParam)
   
End Function


