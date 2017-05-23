Attribute VB_Name = "Menu"
'Using the Menu APIs to Grow or Shrink a Menu During Run-time
'(c) Jon Vote, 2003
'
'Idioma Software Inc.
'jon@idioma-software.com
'www.idioma-software.com
'www.skycoder.com

Option Explicit

'Tells SetWindowLong to subclass
Public Const GWL_WNDPROC = (-4)

'Menu item selected
Public Const WM_COMMAND = &H111

'Possible values for wFlags
Public Const MF_BITMAP = &H4&        'Menu item is bitmap. lpNewItem = handle to bitmap.
Public Const MF_CHECKED = &H8&       'Check flag.
Public Const MF_DISABLED = &H2&      'Disable flag.
Public Const MF_ENABLED = &H0&       'Enable flag.
Public Const MF_GRAYED = &H1&        'Greyed flag.
Public Const MF_MENUBARBREAK = &H20& 'Seperator - verticle line if popup.
Public Const MF_MENUBREAK = &H40&    'Seperator - no columns.
Public Const MF_OWNERDRAW = &H100&   'Owner drawn.
Public Const MF_POPUP = &H10&        'Popup menu (Sub-menu).
Public Const MF_SEPARATOR = &H800&   'Seperator - dropdown only.
Public Const MF_STRING = &H0&        'Item is a string.
Public Const MF_UNCHECKED = &H0&     'Un-check flag.
 
'Refer to menu item by position or command (ID).
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

'First run-time menu item ID.
Private Const FIRST_MENU_NUMBER = 100

'Menu Action Enum - possible user responses
Public Enum MenuAction
   ACTION_CONTINUE = 0
   ACTION_INSERT_ITEM_BEFORE = 1
   ACTION_INSERT_ITEM_AFTER = 2
   ACTION_INSERT_SUBMENU_BEFORE = 3
   ACTION_INSERT_SUBMENU_AFTER = 4
   ACTION_DELETE = 5
End Enum

'Previous process handle
Public g_lpPrevWndFunc As Long

'Used to subclass the form.
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'We will us this to pass control to windows after SubClassHandler is done
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
  (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

'GetMenu returns a handle to the menu
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

'Get Submenu handle
Public Declare Function GetSubMenu Lib "user32" _
  (ByVal hMenu As Long, ByVal nPos As Long) As Long

'Refresh menu display
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

'Creates a new popup menu or sub-menu
Public Declare Function CreatePopupMenu Lib "user32" () As Long

'Get menu item caption
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" _
  (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, _
   ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

'Returns ItemID by Position
Public Declare Function GetMenuItemID Lib "user32" _
  (ByVal hMenu As Long, ByVal nPos As Long) As Long

'Returns number of menu items at this level
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

'Append menu item to end of list
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
    (ByVal hMenu As Long, ByVal wFlags As Long, _
     ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

'Insert a menu item at nPosition
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" _
  (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
   ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long

'Remove menu item
Public Declare Function RemoveMenu Lib "user32" _
  (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'SubClass: Causes all messages to frmForm to be
'sent to the routine SubClassHandler.
Public Function SubClass(frmForm As Form) As Long

  'Make sure the form is visble
  frmForm.Show
  'Tell windows to route the form's messages to SubClassHandler
  g_lpPrevWndFunc = SetWindowLong(frmForm.hWnd, GWL_WNDPROC, _
                         AddressOf SubClassHandler)
  
  'Save the pointer to the system handler
  SubClass = g_lpPrevWndFunc
  
End Function

'SubClassHandler: This gets the Windows messages sent to the form,
'looks for a menu-click.
Public Function SubClassHandler(ByVal hWnd _
    As Long, ByVal lngMsg As Long, ByVal wParam _
    As Long, ByVal lParam As Long) As Long
    
  Dim strMenuClickedMessage As String
  
  'Menu click?
  If lngMsg = WM_COMMAND Then
    If lParam = 0 Then
      'Menu click here, pass the form handle and menu ID
      Call ProcessMenu(hWnd, wParam)
    End If
  End If
  
  'Return control to Windows
  SubClassHandler = CallWindowProc(g_lpPrevWndFunc, _
        hWnd, lngMsg, wParam, lParam)

End Function

'UnSubClass: Turn off message trapping.
Public Function UnSubClass(frmForm As Form) As Long

  Dim lngRC As Long
  lngRC = SetWindowLong(frmForm.hWnd, GWL_WNDPROC, g_lpPrevWndFunc)
  
End Function

'ProcessMenu: Called when the user has clicked a menu item.
Private Sub ProcessMenu(hWnd As Long, lngMenuCommand As Long)
  
  Dim lngRC As Long
  Dim strSubMenuCaption As String
  Dim strMenuItemCaption As String
  Dim hMenu As Long
  Dim hSubMenu As Long
  Dim lngMenuID As Long
  Dim hParentMenu As Long
  Dim lngParentPosition As Long
  Dim hFoundMenu As Long
  Dim lngFoundPosition As Long
  Dim maAction As MenuAction
  Dim bRemoveParent As Boolean
  
  'Get the menu caption of the clicked item
  strMenuItemCaption = GetMenuCaptionByCommand(hWnd, lngMenuCommand)
  
  'Display the menu caption just clicked and prompt the user
  maAction = frmMenuItem.ProcessMenuClick(strMenuItemCaption)
  
  'Continue or take some action?
  If maAction <> ACTION_CONTINUE Then
     hMenu = GetMenu(hWnd)
     'FindMenuByID will return TRUE if lngMenuCommand is found,
     'and update hFoundMenu and lngFoundPosition with the
     'sub-menu handle and position.
     If FindMenuByID(hMenu, lngMenuCommand, hFoundMenu, lngFoundPosition) Then
        Select Case maAction
          'Insert menu item before
          Case ACTION_INSERT_ITEM_BEFORE
            'GetMenuItemCaption prompts the user for a
            'menu item caption, updates strMenuCaption
            'and returns TRUE unless user canceled.
            If GetMenuItemCaption(strMenuItemCaption) Then
               lngRC = InsertMenu(hFoundMenu, lngFoundPosition, _
                          MF_STRING + MF_BYPOSITION, _
                                 GetNextMenuNumber(), strMenuItemCaption)
            End If
            
          'Insert menu item after
          Case ACTION_INSERT_ITEM_AFTER
            If GetMenuItemCaption(strMenuItemCaption) Then
               lngRC = InsertMenu(hFoundMenu, lngFoundPosition + 1, _
                                  MF_STRING + MF_BYPOSITION, _
                                     GetNextMenuNumber, strMenuItemCaption)
            End If
            
          'Insert sub-menu before
          Case ACTION_INSERT_SUBMENU_BEFORE
            'GetSubMenuCaptions prompts the user for sub-menu, menu item captions,
            'updates strMenuCaption, strSubMenuCaption
            'and returns TRUE unless user cancel
            If GetSubMenuCaptions(strSubMenuCaption, strMenuItemCaption) Then
               AddPopupMenu hFoundMenu, strSubMenuCaption, strMenuItemCaption, _
                                                             lngFoundPosition
            End If
            
          'Insert sub-menu after
          Case ACTION_INSERT_SUBMENU_AFTER
            If GetSubMenuCaptions(strSubMenuCaption, strMenuItemCaption) Then
              AddPopupMenu hFoundMenu, strSubMenuCaption, strMenuItemCaption, _
                                                            lngFoundPosition + 1
            End If
            
          'Delete menu item
          Case ACTION_DELETE
            'Delete menu will delete this item and recursively
            'delete any orphaned parents.
            DeleteMenuItem hMenu, hFoundMenu, lngFoundPosition
            
        End Select
        
        'Refresh the menu on the form
        lngRC = DrawMenuBar(hWnd)
      End If 'FindMenuByID(...
  End If 'maAction <> ACTION_CONTINUE
  
End Sub

'GetMenuCaptionByCommand: Returns caption for menu item attached to hWnd
'with ID = lngMenuCommand.
Public Function GetMenuCaptionByCommand(ByVal hWnd As Long, _
                                     lngMenuCommand As Long) As String
 
  Dim lngRC As Long
  Dim lngMenuCount As Long
  Dim hMenu As Long
  Dim hSubMenu As Long
  Dim lngItem As Long
  Dim strString As String
  Dim lngMaxCount As Long
  Dim lngFlag As Long
  
  'Get the form's menu bar
  hMenu = GetMenu(hWnd)
  If hMenu <> 0 Then
    'Initialize the buffer
    strString = Space$(256)
    'lngRC gets the number of characters returned...
    lngRC = _
         GetMenuString(hMenu, lngMenuCommand, strString, Len(strString), MF_BYCOMMAND)
    'Return the item caption
    GetMenuCaptionByCommand = Left$(strString, lngRC)
  Else
    'Something went wrong here - nothing found.
    GetMenuCaptionByCommand = ""
  End If
     
End Function

'GetNextMenuNumber: Return the next menu number starting from FIRST_MENU_NUMBER
Public Function GetNextMenuNumber() As Integer
    
  Static intNextNumber As Integer
  
  If intNextNumber = 0 Then
    intNextNumber = FIRST_MENU_NUMBER
  Else
      intNextNumber = intNextNumber + 1
  End If
  GetNextMenuNumber = intNextNumber
  
End Function

'GetSubMenuCaptions: Prompts user for sub-menu and menu item captions.
'Returns FALSE if user cancel, else returns TRUE, updates strSubMenuCaption, strMenuItemCaption
Public Function GetSubMenuCaptions(ByRef strSubMenuCaption As String, _
                                     ByRef strMenuItemCaption As String) As Boolean
  
  strSubMenuCaption = GetDefaultSubMenuCaption()
  strSubMenuCaption = InputBox$("Please enter the new menu caption:", _
                                   "Add menu", strSubMenuCaption)
  
  If strSubMenuCaption <> "" Then
    strMenuItemCaption = GetDefaultMenuItemCaption()
    strMenuItemCaption = InputBox$("Please enter the new sub menu caption:", _
                                     "Add menu", strMenuItemCaption)
  End If
  
  'Return TRUE if both are non-null
  GetSubMenuCaptions = (strSubMenuCaption <> "") And (strMenuItemCaption <> "")
  
End Function

'GetMenuItemCaption: Prompts user for menu item caption.
'Returns FALSE if user cancel, else returns TRUE, updates strMenuItemCaption
Public Function GetMenuItemCaption(ByRef strMenuItemCaption As String)

  strMenuItemCaption = GetDefaultMenuItemCaption()
  strMenuItemCaption = InputBox$("Please enter the new menu caption:", _
                                    "Add menu", strMenuItemCaption)
                                    
  'Return TRUE if not null.
  GetMenuItemCaption = strMenuItemCaption <> ""
  
End Function

'GetDefaultItemCaption: Returns a default menu item caption.
Private Function GetDefaultMenuItemCaption() As String

  Static intCaptionNumber As Integer
  intCaptionNumber = intCaptionNumber + 1
  GetDefaultMenuItemCaption = "Item_" & Format$(intCaptionNumber, "00")
  
End Function

'GetDefaultSubMenuCaption: Returns a default sub menu caption.
Private Function GetDefaultSubMenuCaption() As String

  Static intSubMenuNumber As Integer
  intSubMenuNumber = intSubMenuNumber + 1
  GetDefaultSubMenuCaption = "Submenu_" & Format$(intSubMenuNumber, "00")
  
End Function

'FindMenuByID: Returns TRUE if lngFindMenuID is found.
'Updates hFoundMenu, lngFoundPosition
Public Function FindMenuByID(ByVal hMenu As Long, ByVal lngFindMenuID As Long, _
    ByRef hFoundMenu As Long, ByRef lngFoundPosition As Long) As Boolean
  
  Dim lngPosition As Long
  Dim lngCount As Long
  Dim lngMenuID As Long
  Dim hSubMenu As Long
  
  FindMenuByID = False
  
  'Get the number of items at this level
  lngCount = GetMenuItemCount(hMenu)
  
  'Loop for each item
  For lngPosition = 0 To lngCount - 1
    
    'Get the menu ID for the item at this position
    lngMenuID = GetMenuItemID(hMenu, lngPosition)
    
    'We are done if tnis ID matches lngFindMenuID
    If lngMenuID = lngFindMenuID Then
       hFoundMenu = hMenu
       lngFoundPosition = lngPosition
       FindMenuByID = True
       Exit Function
    'No match here - is this a sub-menu?
    ElseIf lngMenuID = -1 Then
       'We have a sub-menu here get the sub-menu handle.
       hSubMenu = GetSubMenu(hMenu, lngPosition)
       
       'Recurse back with the sub-menu handle.
       'We are done if we got a hit.
       If FindMenuByID(hSubMenu, lngFindMenuID, hFoundMenu, lngFoundPosition) Then
         FindMenuByID = True
         Exit For
       End If
    End If
  Next lngPosition
  
End Function

'AddPopupMenu: Append or Insert a new menu/sub-menu pair to hMenu.
Public Sub AddPopupMenu(ByVal hMenu As Long, strItemCaption As String, _
                        strSubItemCaption As String, Optional varPosition As Variant)
  
  Dim hPopupMenu As Long
  Dim lngRC As Long
  
  'Create a new popup menu handle
  hPopupMenu = CreatePopupMenu()
  
  'Append the new item to the new sub-menu
  lngRC = AppendMenu(hPopupMenu, MF_STRING, GetNextMenuNumber(), _
                     strSubItemCaption)
  
  'Append the new sub-menu if no position passed, else insert at varPosition
  If IsMissing(varPosition) Then
    lngRC = AppendMenu(hMenu, MF_POPUP, hPopupMenu, strItemCaption)
  Else
    lngRC = InsertMenu(hMenu, varPosition, MF_POPUP + MF_BYPOSITION, _
                       hPopupMenu, strItemCaption)
  End If
  
End Sub

'DumpMenu: Accepts a Menu Handle,
'Recursively dumps all items and sub-menu items to debug window.
Public Sub DumpMenu(ByVal hMenu As Long)
  
  Dim lngPosition As Long
  Dim lngCount As Long
  Dim lngMenuID As Long
  Dim hSubMenu As Long
  Dim strMenuCaption As String
  Dim lngStrLen As Long
  
  'Get the number of items at this level
  lngCount = GetMenuItemCount(hMenu)
    
  'Loop for each item
  For lngPosition = 0 To lngCount - 1
    
    'Get the menu caption for this item
    strMenuCaption = Space$(256)
    lngStrLen = GetMenuString(hMenu, lngPosition, strMenuCaption, _
            Len(strMenuCaption), MF_BYPOSITION)
    strMenuCaption = Left$(strMenuCaption, lngStrLen)
    
    'Get the Menu ID for this item
    lngMenuID = GetMenuItemID(hMenu, lngPosition)
            
    'Dump the menu handle, menu ID, position, caption and item count.
    Debug.Print strMenuCaption, hMenu, lngMenuID, lngPosition
    
    'A -1 means this entry is itself another menu
    'If so, we will recursively call this routine,
    'passing the sub-menu handle.
    If lngMenuID = -1 Then
       'We have a sub-menu here,
       'get the sub-menu handle and recurse
       hSubMenu = GetSubMenu(hMenu, lngPosition)
       Call DumpMenu(hSubMenu)
       'Just a menu item here -
    End If
    
  Next lngPosition
  
End Sub

'DeleteMenuItem: Delete's a menu item and recursivly
'deletes any orphaned parents.
Public Sub DeleteMenuItem(ByVal hMenuBar As Long, hDeleteMenu As Long, _
                                                       lngDeletePosition As Long)
  Dim lngItemCount As Long
  Dim hParentMenu As Long
  Dim lngParentPosition As Long
  Dim bDeleteParent As Boolean
  Dim lngRC As Long
   
  'If the item count is 1, this is the last
  'menu item and we want to also delete the parent.
  If GetMenuItemCount(hDeleteMenu) = 1 Then
     'We want to delete the parent here, grab the
     'parent's menu handle and position.
     'We should always get a TRUE here...
     bDeleteParent = GetParentMenu(hMenuBar, hDeleteMenu, hParentMenu, lngParentPosition)
  Else
     bDeleteParent = False
  End If
  
  'Delete this item and recurse to delete the parent if applicable.
  lngRC = RemoveMenu(hDeleteMenu, lngDeletePosition, MF_BYPOSITION)
  If bDeleteParent Then
     DeleteMenuItem hMenuBar, hParentMenu, lngParentPosition
  End If

End Sub

'GetParentMenu: Begins search at hMenuBar.
'Returns TRUE if parent menu found, updates hParentMenu, hParentPosition -
'else returns FALSE
Public Function GetParentMenu(ByVal hMenuBar As Long, ByVal hChildMenu As Long, _
                          ByRef hParentMenu As Long, ByRef hParentPosition As Long) As Long
  
  Dim lngPosition As Long
  Dim lngCount As Long
  Dim lngMenuID As Long
  Dim hSubMenu As Long
  Const NO_PARENT = -1
  
  'Default to no parent
  GetParentMenu = NO_PARENT
  
  'Get the number of items at this level
  lngCount = GetMenuItemCount(hMenuBar)
    
  'Loop for each item
  For lngPosition = 0 To lngCount - 1
    
    'Check each sub-menu looking for hChildMenu
    lngMenuID = GetMenuItemID(hMenuBar, lngPosition)
    If lngMenuID = -1 Then
       'We have a sub-menu here. We are done
       'if the sub-menu handle matches...
       hSubMenu = GetSubMenu(hMenuBar, lngPosition)
       If hSubMenu = hChildMenu Then
          hParentMenu = hMenuBar
          hParentPosition = lngPosition
          GetParentMenu = True
       Else
          'Didn't match here, recurse back to check this sub-menu.
          GetParentMenu = GetParentMenu(hSubMenu, hChildMenu, hParentMenu, hParentPosition)
       End If
    End If
    
  Next lngPosition
  
End Function
