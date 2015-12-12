VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Created & released by KSY, 06/14/2003
'
#Const SUBCLASS_MENUBAR = 1
Private m_colMenus As New Collection

Public Sub Clear()
   Set m_colMenus = New Collection
End Sub

Public Sub Remove(Index)
   m_colMenus.Remove Index
End Sub

Public Sub Swap(Index1 As Long, Index2 As Long)
   'Swap the values of the two items.
   Dim oTemp As VBMenu
   Set oTemp = m_colMenus(Index1)
   SetItem Index1, m_colMenus(Index2)
   SetItem Index2, oTemp
End Sub

Public Sub SetItem(Index As Long, ByVal newMenu As VBMenu)
   'Changes the menu item for the given index.
   m_colMenus.Add newMenu, , , Index
   m_colMenus.Remove Index
End Sub

Public Property Get Items() As Collection
   Set Items = m_colMenus
End Property
Public Property Set Items(ByVal objItems As Collection)
   Set m_colMenus = objItems
End Property

Public Sub ParseMenuText(MenuText As String, _
      Optional MenuName As String = "파일 메뉴", _
      Optional FileName As String, _
      Optional BeginDelimiter As String = ksBegin_VB_Menu, _
      Optional EndDelimiter As String = ksEnd & vbCrLf)
      
   'Parse the given menu text and add VBMenu items to the collection.
   Dim oParser As VBParser
   Set oParser = New VBParser
   With oParser
      Call .ParseMenu(MenuText, BeginDelimiter, EndDelimiter)
      Set m_colMenus = .Matches
   End With
   Set oParser = Nothing
End Sub

Public Sub ParseFile(FileName As String, _
      Optional MenuName As String = "파일 메뉴", _
      Optional BeginDelimiter As String = ksBegin & vbSpace, _
      Optional EndDelimiter As String = ksEnd & vbCrLf)
      
      'Parse the menu text from a given file and add VBMenu items to the collection.
      Call ParseMenuText(VBGetModuleHead(GetFileText(FileName)), _
                              MenuName, FileName, BeginDelimiter, EndDelimiter)
End Sub

Public Property Get Count() As Long
   Count = m_colMenus.Count
End Property
Public Property Get Key(Index As Long) As String
   Key = m_colMenus.Key(Index)
End Property
Public Property Get Item(Index) As VBMenu
Attribute Item.VB_UserMemId = 0
   On Error GoTo Bye
   Set Item = m_colMenus(Index)
Bye:
End Property
Public Property Set Item(Index, ByVal oVBMenu As VBMenu)
   Set m_colMenus(Index) = oVBMenu
End Property

Public Sub SetPopupProperties()
   'Set the IsParent property with level information for all items in collection.
   Dim i As Long, nCount As Long
   nCount = Count
   For i = 1 To nCount - 1
      With Item(i)
         .IsParent = (.Level < Item(i + 1).Level) Or .Level = 1
      End With
   Next i
   With Item(nCount)
      .IsParent = (.Level = 1)
   End With
End Sub

Public Function AddMenu( _
         ByVal Level As Long, Name As String, Optional Caption As String = "-", _
         Optional ByVal ShortcutIndex As eShortCut = SC_000_NONE, _
         Optional ByVal IsParent As Boolean, _
         Optional ByVal Index As Long = -1, _
         Optional ByVal Checked As Boolean, _
         Optional ByVal Visible As Boolean = True, _
         Optional ByVal Enabled As Boolean = True, _
         Optional ByVal HelpContextID As Long, _
         Optional ByVal NegotiatePosition As Long, _
         Optional ByVal WindowList As Boolean) As VBMenu

   Set AddMenu = New VBMenu
   With AddMenu
      .Name = Name
      .Level = Level
      .IsParent = IsParent
      .Index = Index
      .Caption = Caption
      .ShortcutIndex = ShortcutIndex
      .Enabled = Enabled
      .Checked = Checked
      .Visible = Visible
      .HelpContextID = HelpContextID
      .NegotiatePosition = NegotiatePosition
      .WindowList = WindowList
      If Index = -1 Then
         .Key = Name
      Else
         .Key = Name & "(" & Index & ")"
      End If
      m_colMenus.Add AddMenu
   End With
End Function
Private Function CreateDefaultMenu() As String
      
      Set m_colMenus = New Collection
      Call AddMenu(0, "mnuFile", "파일(&F)", , True)
      Call AddMenu(1, "mnuFileOpen", "열기(&O)", SC_CTRL_O)
      Call AddMenu(1, "mnuFileBar1")
      Call AddMenu(1, "mnuFileSave", "저장(&S)", SC_CTRL_S)
      Call AddMenu(1, "mnuFileSaveAs", "다른 이름으로 저장(&A)...", SC_CTRL_F12)
      Call AddMenu(1, "mnuFileBar2")
      Call AddMenu(1, "mnuFileExit", "종료(&X)", SC_CTRL_F4)
      Call AddMenu(0, "mnuEdit", "편집(&E)", , True)
      Call AddMenu(1, "mnuEditUndo", "실행취소(&Z)", SC_CTRL_Z)
      CreateDefaultMenu = GetMenuText(False)
End Function


'========== Creating Menu Text ======================
'Needs Add & Copy for CreateLinearToStructured
'
Public Function GetMenuText(Optional bCreateDefault As Boolean, Optional IsStuctured As Boolean) As String
   'Create the whole menu text to save.
   
   'If this collection is linear, we will create a new copy that is structured.
   'Structured VBMenus collection is useful for creating tree stuctured menu text.
   'Linear VBMenus collection is useful to display on the UI object such as ListBox.
   
   'Internally, we use always linear collection for UI.
   'and, use structured copy to create menu text to save.
   
   'If default menu is requested,
   If bCreateDefault Then
      CreateDefaultMenu
      pCreateMenuText CreateLinearToStructured(Me), GetMenuText
      Exit Function
   End If
   
   'If menu text for this collection is requested.
   If m_colMenus.Count Then
      If IsStuctured Then
         pCreateMenuText Me, GetMenuText
      Else
         pCreateMenuText CreateLinearToStructured(Me), GetMenuText
      End If
   End If
End Function

Public Function Add(Optional ByVal oMenu As VBMenu, Optional Key As String, Optional Before, Optional After) As VBMenu
   'Add a new VBMenu item. If oMenu is Nothing, create a new one before adding.
   If ObjPtr(oMenu) = 0 Then
      Set Add = New VBMenu
   Else
      Set Add = oMenu
   End If
   m_colMenus.Add Add, , Before, After
End Function

Public Function Copy() As VBMenus
   'Copies the whole menu items of this collection.
   'Used to create a structured collection.
   Set Copy = New VBMenus
   Dim i As Long
   For i = 1 To Count
      Copy.Add Item(i).Copy
   Next i
End Function


Private Sub pCreateMenuText(ByVal colStructuredMenus As VBMenus, strResult As String)
   'Create menu text to save. The given collection must be a structured one(i.e.,
   'must have parent-children relationship).
   
   'This precedure is recursive.
   
   Dim oMenu As VBMenu
   Dim i As Long
   For i = 1 To colStructuredMenus.Count
      With colStructuredMenus(i)
         'if parent,
         If .IsParent Then
            'Collect  the properties of the current item.
            strResult = IIf(LenB(strResult) > 0, _
                                 strResult & vbCrLf & .GetMenuText(VBMC_BEGIN), _
                                 .GetMenuText(VBMC_BEGIN))
                                 
            'Collect the menu text for the children with recursive call.
            pCreateMenuText .Children, strResult
            
            'Collect  the end text of the current item.
            strResult = strResult & vbCrLf & .GetMenuText(VBMC_END)
         Else
            'if not parent, just collect the menu text.
           strResult = strResult & vbCrLf & .GetMenuText
         End If
      End With
   Next
End Sub

Public Function CreateLinearToStructured(ByVal colLinearMenus As VBMenus) As VBMenus
   'Returns a new menu collection
   'structured with parent/children relationship.
   
   'Level must be previously set.
   
   Dim i As Long, oParent As VBMenu
   Dim colMenus As VBMenus, oMenu As VBMenu, nCount As Long
   Dim iStartLevel As Long
   
   'Create new collection
   Set CreateLinearToStructured = New VBMenus
   
   'Copy the menus from the given collection.
   'We need to copy it because we don't want to
   'change the linearity of the orignal collection used for UI.
   Set colMenus = colLinearMenus.Copy
      
   With colMenus
      'Set the start level to the firs item's level.
      iStartLevel = .Item(1).Level
      For i = 1 To .Count
         Set oMenu = colMenus(i)
         With oMenu
            If .Level > iStartLevel Then
               'if level is larger than the start level, find its parent.
               Set oParent = FindLastPrevParent(colMenus, i - 1, .Level)
               'Add to the parent's children collection.
               oParent.Children.Add oMenu
            Else
               CreateLinearToStructured.Add oMenu
            End If
         End With
      Next i
   End With
End Function

Private Function FindLastPrevParent(ByVal colLinear As VBMenus, ByVal iStart As Long, ByVal iStartLevel As Long) As VBMenu
   Dim i As Long, oMenu As VBMenu
   With colLinear
      For i = iStart To 1 Step -1
         Set oMenu = .Item(i)
         With oMenu
             If .Level < iStartLevel Then
               'Found. Return the found item.
               Set FindLastPrevParent = oMenu
               Exit Function
            End If
         End With
      Next
      'No found. Return the first item.
      Set FindLastPrevParent = .Item(1)
   End With
End Function

'
'========= Validating Menus Created ===========================
'
Public Function Validate( _
                     ByRef ErrorItem As VBMenu, _
                     ByRef ErrorItemIndex As Long, _
                     ByRef ErrorDescription As String) As eVBMenuValidate
   
   SetPopupProperties
   Validate = ValidateMenu(Me, ErrorItem, ErrorItemIndex, ErrorDescription)
End Function

Private Function ValidateMenu(ByVal colMenus As VBMenus, _
                     ByRef ErrorItem As VBMenu, _
                     ByRef ErrorItemIndex As Long, _
                     ByRef ErrorDescription As String) As eVBMenuValidate

'Public Enum eVBMenuValidate
'   VBM_ERR_NONE = 0
'   VBM_ERR_MENU_MUST_HAVE_NAME
'   VBM_ERR_LEVEL_SKIP
'   VBM_ERR_PARENT_CANNOT_HAVE_SHORTCUT
'   VBM_ERR_PARENT_CANNOT_BE_CHECKED
'   VBM_ERR_PARENT_CANNOT_BE_SEPARATOR
'   VBM_ERR_NAME_CANNOT_BE_DUPLICATED
'   VBM_ERR_INVALID_INDEX
'   VBM_ERR_INDEX_MUST_BE_LARGER_THAN_PREVIOUS
'   VBM_ERR_INDEXED_ITEMS_MUST_BE_SEQUENTIAL
'   VBM_ERR_INDEXED_ITEMS_MUST_HAVE_SAME_LEVEL
'   VBM_ERR_CHILD_MUST_HAVE_DIFFERENT_NAME
'   VBM_ERR_PARENT_MUST_HAVE_DIFFERENT_NAME
'   VBM_ERR_ONLY_ONE_WINDOW_LIST_ALLOWED
'   VBM_ERR_ONLY_PARENT_CAN_HAVE_WINDOW_LIST
'End Enum

   Dim i As Long
   Dim nMenuCount As Long
   Dim ncWindowList As Long
   
   Dim strPrevIndexedItemName  As String
   Dim nPrevIndexedItemSeq As Long
      
   Dim colCheckName As Collection 'Used for name duplication checking
   Set colCheckName = New Collection
   
   Dim oMenu As VBMenu
   nMenuCount = colMenus.Count
   
   For i = 1 To colMenus.Count
      Set oMenu = colMenus(i)
      
      With oMenu
      
         'Check Name
         If LenB(.Name) = 0 Then
            Set ErrorItem = oMenu: ErrorItemIndex = i
            ValidateMenu = VBM_ERR_MENU_MUST_HAVE_NAME
            ErrorDescription = "Menu must have a name."
            Exit Function
         End If
         
         'Check Name Duplication
         On Error Resume Next
         colCheckName.Add 0, .Name & "(" & .Index & ")"
         If Err.Number Then
            Set ErrorItem = oMenu: ErrorItemIndex = i
            ValidateMenu = VBM_ERR_NAME_CANNOT_BE_DUPLICATED
            ErrorDescription = "Menu name cannot be duplicated"
            Exit Function
         End If
         On Error GoTo 0
        
         'Check level skip
         If i < nMenuCount Then
            If (colMenus(i + 1).Level - .Level) > 1 Then
               Set ErrorItem = colMenus(i + 1): ErrorItemIndex = i + 1
               ValidateMenu = VBM_ERR_LEVEL_SKIP
               ErrorDescription = "Menu Item skipped a level."
               Exit Function
            End If
         End If
         
         'Check index
         If .Index < -1 Then
            Set ErrorItem = oMenu: ErrorItemIndex = i
            ValidateMenu = VBM_ERR_INVALID_INDEX
            ErrorDescription = "Invalid index."
            Exit Function
         End If
         
         'check name
         If i > 1 Then
            If .Name = colMenus(i - 1).Name Then
               If .Level <> colMenus(i - 1).Level Then
                  Set ErrorItem = oMenu: ErrorItemIndex = i
                  ValidateMenu = VBM_ERR_INDEXED_ITEMS_MUST_HAVE_SAME_LEVEL
                  ErrorDescription = "Indexed items must be included in the same parent."
                  Exit Function
               End If
            End If
         End If
         
         If .Index > 0 Then
            If strPrevIndexedItemName = .Name Then
               If i - nPrevIndexedItemSeq = 1 Then
                  If nPrevIndexedItemSeq > 0 Then
                     If .IsParent <> colMenus(nPrevIndexedItemSeq).IsParent Then
                        If .IsParent Then
                           'Set ErrorItem = oMenu: ErrorItemIndex = i
                           'ValidateMenu = VBM_ERR_PARENT_MUST_HAVE_DIFFERENT_NAME
                           'ErrorDescription = "Popup item cannot have the same name with the previous non popup item. Indexed name cannot be used for this popup item."
                           'Exit Function
                        Else
                           Set ErrorItem = oMenu: ErrorItemIndex = i
                           ValidateMenu = VBM_ERR_CHILD_MUST_HAVE_DIFFERENT_NAME
                           ErrorDescription = "Child cannot have the same name with the parent. Indexed name cannot be used for this child."
                           Exit Function
                        End If
                     End If
                  End If
               Else
                  Set ErrorItem = oMenu: ErrorItemIndex = i
                  ValidateMenu = VBM_ERR_INDEXED_ITEMS_MUST_BE_SEQUENTIAL
                  Exit Function
               End If
            End If
            'Set the previous index item
            strPrevIndexedItemName = .Name
            nPrevIndexedItemSeq = i
         End If
         
         If i > 1 Then
            If colMenus(i - 1).Name = .Name Then
               If colMenus(i - 1).Index >= .Index Then
                  Set ErrorItem = oMenu: ErrorItemIndex = i
                  ValidateMenu = VBM_ERR_INDEX_MUST_BE_LARGER_THAN_PREVIOUS
                  ErrorDescription = "Index must be larger than the previous item."
                  ''Debug.Print colMenus(i - 1).Key, .Key
                  Exit Function
               End If
            End If
         End If
         'Check Index Sequence
         'Set ErrorItem = oMenu: ErrorItemIndex = i
         'ValidateMenu = VBM_ERR_INVALID_INDEX_SEQUENCE
         'ErrorDescription = "Invalid menu index sequence."
                 
         If .IsParent Then
            If .Caption = "-" Then
               Set ErrorItem = colMenus(i + 1): ErrorItemIndex = i + 1
               ValidateMenu = VBM_ERR_PARENT_CANNOT_BE_SEPARATOR
               ErrorDescription = "Parent Menu cannot be a Separator."
               Exit Function
            End If
            
            'Check whether the popup menu has a short cut.
            If .ShortcutIndex <> SC_000_NONE Then
               Set ErrorItem = oMenu: ErrorItemIndex = i
               ValidateMenu = VBM_ERR_PARENT_CANNOT_HAVE_SHORTCUT
               ErrorDescription = "Parent Menu cannot have a Shortcut."
               Exit Function
            End If
            
            'Check whther the popup menu is checked.
            If .Checked Then
               Set ErrorItem = oMenu: ErrorItemIndex = i
               ValidateMenu = VBM_ERR_PARENT_CANNOT_BE_CHECKED
               ErrorDescription = "Parent Menu cannot be Checked."
               Exit Function
            End If
            
            If .WindowList Then
               ncWindowList = ncWindowList + 1
               If ncWindowList > 1 Then
                  Set ErrorItem = oMenu: ErrorItemIndex = i
                  ValidateMenu = VBM_ERR_ONLY_ONE_WINDOW_LIST_ALLOWED
                  ErrorDescription = "Only one Window List allowed"
                  Exit Function
               End If
            End If
         Else
            'If .WindowList Then
            '   Set ErrorItem = oMenu: ErrorItemIndex = i
            '   ValidateMenu = VBM_ERR_ONLY_PARENT_CAN_HAVE_WINDOW_LIST
            '   ErrorDescription = "Only Parent Menu can be a Window List."
            '   Exit Function
            'End If
         End If
      End With
   Next i

   ValidateMenu = VBM_ERR_NONE
End Function


Public Function SaveMenus(ByVal FileName As String, _
                                 Optional bIsStructured As Boolean, _
                                 Optional bMakeBackup As Boolean, _
                                 Optional bInsertIfMenuNotExists As Boolean = True) As Boolean
   
   'Create the whole menu text  and save to the given file.
   Dim strMenuText As String
   
   If FileExists(FileName) Then
      strMenuText = GetMenuText(False, bIsStructured)
      SaveMenus = VBReplaceMenuText(FileName, GetMenuText(False, bIsStructured), _
                                       bMakeBackup, bInsertIfMenuNotExists)
                                       
   ElseIf LenB(FileName) > 0 Then
      SaveMenus = SaveAsTemplate(FileName, GetFileName(FileName, efpBaseName), _
                                       bIsStructured, bMakeBackup)
   End If
   

End Function

Public Function SaveAsTemplate(ByVal TemplateFileName As String, _
                                 Optional TemplateFileCaption As String = "파일메뉴", _
                                 Optional bIsStructured As Boolean, _
                                 Optional bMakeBackup As Boolean) As Boolean

   Dim strMenuText As String
   
   If bMakeBackup Then
      BackupFile TemplateFileName, False
   End If
   strMenuText = GetMenuText(False, bIsStructured)
   strMenuText = ksVERSION_500 & _
                           vbCrLf & ksBegin_VB_Form & vbSpace & "frmMenuTemplate" & _
                           vbCrLf & vbSpace3 & ksCaption & ksEqual & vbQ & TemplateFileCaption & vbQ & _
                           vbCrLf & strMenuText & _
                           vbCrLf & ksEnd & _
                           vbCrLf & ksOption_Explicit
   SaveAsTemplate = Save(TemplateFileName, strMenuText)

End Function

Public Function Import(ByVal FileName As String, Optional bReplaceMenus As Boolean) As Boolean
   'Collect the menu items form the given file.
   
   If Not FileExists(FileName) Then
      Exit Function
   End If
   Dim colMenus As VBMenus
   Set colMenus = New VBMenus
   colMenus.ParseMenuText GetFileText(FileName)
   
   'If replacing requested,
   If bReplaceMenus Then
      Set m_colMenus = colMenus.Items
   Else 'else, just add the colected items.
      AddMenus colMenus
   End If
   Set colMenus = Nothing
End Function

Public Sub AddMenus(ByVal colMenusFrom As VBMenus, _
                  Optional Before, Optional After, Optional IgnoreError As Boolean)
                  
   'Add the given menu collection to this collection.
   Dim i As Long
   If IgnoreError Then
      On Error Resume Next
   End If
   
   With colMenusFrom
      If IsMissing(Before) And IsMissing(After) Then
         For i = 1 To .Count
            Add .Item(i)
         Next i
      ElseIf IsMissing(After) Then
         For i = .Count To 1 Step -1
            Add .Item(i), , Before
         Next i
      Else
         For i = 1 To .Count
            Add .Item(i), , , After
         Next i
      End If
   End With
   
   If IgnoreError Then
      On Error GoTo 0
   End If
End Sub

Public Sub CopyToClipboard(Index As Long, _
                  Optional bIncludeChildrenIfParent As Boolean = True)
                                    
   'Copy the whole menu text to clipboard with start level info.
   'We use a custom format to identified whether the clipboard item is a menu text.
   Const ksCF_VBMENU As String = "CF_VBMENU"
   Static CF_VBMENU As Long
   Dim strClipData As String
   
   'Register clipboard format.
   If CF_VBMENU = 0 Then
      'If the format is already registerd, RegisterClipboardFormat returns its ID.
      CF_VBMENU = RegisterClipboardFormat(ksCF_VBMENU)
   End If
   
   'If the custom clipboard format is registered.
   If CF_VBMENU Then
      strClipData = GetNodeText(Index, bIncludeChildrenIfParent)
      
      'Copy the created menu text with start level info.
      If LenB(strClipData) Then
         strClipData = "Start Level=" & Item(Index).Level & vbCrLf & vbLine & vbCrLf & _
                                strClipData & vbCrLf
         'Hand off to the cipboard.
         SetClipText strClipData, CF_VBMENU
      End If
   End If
End Sub

Public Function GetNodeText(Index As Long, _
                  Optional bIncludeChildrenIfParent As Boolean = True) As String
  
   Dim i As Long
   Dim colClipMenus As VBMenus
   Dim oMenu As VBMenu
   Dim nCount As Long
  
   'Validate count
   nCount = Count
   If nCount = 0 Or Index < 0 Or Index > nCount Then
      Exit Function
   End If
   
   'Get a temporary reference.
   Set oMenu = Item(Index)

   'Set the popup properties
   SetPopupProperties
      
   'If parent, copy its children
   If bIncludeChildrenIfParent And oMenu.IsParent Then
      Set colClipMenus = New VBMenus
      colClipMenus.Add oMenu.Copy
      For i = Index + 1 To nCount
         If Item(i).Level > oMenu.Level Then
            colClipMenus.Add Item(i).Copy
         Else
            Exit For
         End If
      Next
      GetNodeText = colClipMenus.GetMenuText
   ElseIf oMenu.IsParent Then
      With oMenu
         GetNodeText = .GetMenuText(VBMC_BEGIN) & vbCrLf & .GetMenuText(VBMC_END)
      End With
   Else 'Just copy the data
      GetNodeText = oMenu.GetMenuText
   End If
   
End Function

Public Function GetFromClipboard() As VBMenus
   
   'Read menu text from the clipboard if available.
   'and then parses the text into a new VBMenus collection.
   
   Const ksCF_VBMENU As String = "CF_VBMENU"
   Static CF_VBMENU As Long
   
   Dim strClipData As String
   Dim i As Long, iStartLevel As Long
   
   'Register clipboard format.
   If CF_VBMENU = 0 Then
      'If the format is already registerd, RegisterClipboardFormat returns its ID.
      CF_VBMENU = RegisterClipboardFormat(ksCF_VBMENU)
   End If
   
   Set GetFromClipboard = New VBMenus
   
   If CF_VBMENU Then
      'Get the text from the clipbaord
      strClipData = GetClipText(CF_VBMENU)
      
      If LenB(strClipData) = 0 Then
         Exit Function
      Else
         strClipData = strClipData & vbCrLf
      End If
      
      'Pase the start level.
      iStartLevel = Val(Between(strClipData, "Start Level=", vbCrLf))
      
      'Create collection with the start level info.
      With GetFromClipboard
         'Parse the text.
         .ParseMenuText strClipData
         
         'Adjust the level for all items.
         For i = 1 To .Count
            With .Item(i)
               .Level = .Level + iStartLevel
            End With
         Next i
      End With
   End If
   
End Function

Public Function CanPaste() As Boolean
   'Returns whether a custom menu text is on the clipbaord.
   Const ksCF_VBMENU As String = "CF_VBMENU"
   Static CF_VBMENU As Long
   If CF_VBMENU = 0 Then
      CF_VBMENU = RegisterClipboardFormat(ksCF_VBMENU)
   End If
   If CF_VBMENU Then
      CanPaste = IsClipboardFormatAvailable(CF_VBMENU)
   End If
End Function

Public Function MoveNodes(ByVal FromIdx As Long, ByVal ToIdx As Long, _
                  Optional bIncludeChildrenIfParent As Boolean = True) As String
  
   Dim i As Long
   Dim colClipMenus As New VBMenus
   Dim oMenu As VBMenu
   Dim nCount As Long
  
   'Validate count
   nCount = Count
   If nCount = 0 Or ToIdx < 1 Or ToIdx > nCount _
      Or FromIdx < 1 Or FromIdx > nCount Then
      Exit Function
   End If

   'Set the popup properties
   SetPopupProperties
      
   'Crete temp collection
   Set colClipMenus = New VBMenus
   
   'Get a temporary reference.
   Set oMenu = Item(FromIdx)
   
   'If parent, copy its children
   If bIncludeChildrenIfParent And oMenu.IsParent Then
      colClipMenus.Add oMenu
      For i = FromIdx + 1 To nCount
         If Item(i).Level > oMenu.Level Then
            colClipMenus.Add Item(i)
         Else
            Exit For
         End If
      Next
   Else
      colClipMenus.Add oMenu
   End If
   
   With colClipMenus
      For i = .Count To 1 Step -1
         Items.Remove FromIdx + i - 1
      Next i
      For i = .Count To 1 Step -1
         Items.Add .Item(i), , ToIdx
      Next i
   End With
   
End Function

Public Sub CreateShowAPIPopupPreview()
   Call CreateShowAPIPopupMenu(Me)
End Sub

Private Sub CreateShowAPIPopupMenu( _
               ByVal colLinearMenus As VBMenus, _
               Optional ByVal x As Long, Optional ByVal y As Long)

   Dim hMenu As Long
   Dim hWndOwner As Long
   Dim nItemID As Long
   Dim TotalCount As Long
   Dim colStructured As VBMenus
   
   If colLinearMenus.Count = 0 Then
      Exit Sub
   End If
   
   Set colStructured = colLinearMenus.CreateLinearToStructured(colLinearMenus)
   
   'Create a blank main menu.
   hMenu = CreatePopupMenu()
   Call pCreatePreviewMenu(colStructured, hMenu, TotalCount)
   
   hWndOwner = GetForegroundWindow()
   nItemID = PopupMenuEx(hMenu, hWndOwner, x, y)
   If nItemID <> 0 Then
      MsgBox GetMenuCaption(hMenu, nItemID) & vbCrLf & "ItemID=" & nItemID
   End If
   Call DestroyMenu(hMenu)

End Sub

Public Sub CreateShowAPIMenuBarPreview(Optional ByVal OwnerForm As VB.Form)
   Call CreateShowAPIMenuBarWindow(Me, OwnerForm)
End Sub

Private Function CreateShowAPIMenuBarWindow( _
   ByVal colMenus As VBMenus, _
   Optional ByVal OwnerForm As VB.Form) As Long
   
   Dim hMainMenu As Long
   Dim hWnd As Long
   Dim bInIDE As Boolean
   Dim colStructured As VBMenus
   Dim TotalCount As Long
   Dim hPrevMenu As Long
   
   Dim frm As frmPreview
   Set frm = New frmPreview
   
   hWnd = frm.hWnd
   'NOTE:
   'Creating menubar is supported only when the form has a menu item (invisible item OK).
   'If there is no menu item, the menu bar will be created at first at runtime.
   'But, it is not fixed. In this case, the API drawn menu bar may disappear unexpectedly.
   'And, API drawn menu will not be even created in IDE if the form has no menu item.
   
   'Create menu preview.
   hPrevMenu = GetMenu(hWnd)
   hMainMenu = CreateMenu()
   
   Set colStructured = colMenus.CreateLinearToStructured(colMenus)
   pCreatePreviewMenu colStructured, hMainMenu, TotalCount
   
   'SetItem the menu for the preview window.
   SetMenu hWnd, hMainMenu
   
   'Once changed the menus with SetMenu, we have to draw menubar.
   DrawMenuBar hWnd
   
   If hPrevMenu Then
      DestroyMenu hPrevMenu
   End If
   'Start subclassing to display message when the user click an menu item.
   #If SUBCLASS_MENUBAR Then
      Call BeginMenuBarPreviewSubclassing(hWnd)
   #End If
   
   frm.txtResult.ColourEntireRTB GenerateAPIDrawMenuBarCode(colMenus), True
   
   'Show the preview window.
   If ObjPtr(OwnerForm) Then
      frm.Show vbModal, OwnerForm
   Else
      frm.Show vbModal
   End If
   
   'Unsubclass.
   #If SUBCLASS_MENUBAR Then
      Call EndMenuBarPreviewSubclassing(hWnd)
   #End If
   'Destory the menu created. (actually this may be unnecessary,
   'because the attatched menu is destroyed automtically when the owner
   'window is destroyed).
   DestroyMenu hMainMenu
   Set frm = Nothing
End Function


Private Sub pCreatePreviewMenu( _
   ByVal colStructuredMenus As VBMenus, hMenu As Long, TotalCount As Long)
   
   'Create menu text to save. The given collection must be a structured one(i.e.,
   'must have parent-children relationship).
   
   'This precedure is recursive.
   Dim i As Long
   Dim hSubMenu As Long
   
   With colStructuredMenus
      For i = 1 To .Count
         TotalCount = TotalCount + 1
         With .Item(i)
            'if parent,
            If .IsParent Then
               hSubMenu = CreateSubMenu(hMenu, .Caption, 1000 + TotalCount)
               'Collect the menu text for the children with recursive call.
               Call pCreatePreviewMenu(.Children, hSubMenu, TotalCount)
            Else
               'if not parent, just collect the menu text.
               If .Caption = "-" Then
                  Call AddSeparator(hMenu, 1000 + TotalCount)
               Else
                  Call AddMenuItem(hMenu, .Caption, 1000 + TotalCount, .ShortcutIndex)
               End If
            End If
         End With
      Next
   End With
End Sub

Public Function CreateAPIDrawMenuBarCode() As String
   CreateAPIDrawMenuBarCode = GenerateAPIDrawMenuBarCode(Me)
End Function

Private Function GenerateAPIDrawMenuBarCode(ByVal colLinearMenus As VBMenus) As String

   Dim colText As New Collection
   Dim FuncNameSuffix As String
   
   FuncNameSuffix = Format$(Now(), "YYMMDDHHMMSS")
   With colText
      .Add "'#Const SUBCLASS_MENUBAR_" & FuncNameSuffix & " = 0"
      .Add "Private m_lpfnMenuBarOldWndProc_" & FuncNameSuffix & " As Long"
      .Add "Private m_hWnd_" & FuncNameSuffix & " As Long"
   
      .Add vbNullString
      .Add "Public Sub DrawMenuBar_" & FuncNameSuffix & "(Optional ByVal OwnerForm As VB.Form)"
      .Add vbNullString
      .Add "   Dim frm As frmPreview"
      .Add "   Set frm = New frmPreview"
      .Add "    'Draw menu bar for the given form."
      .Add "    'Creating menubar is supported only when the form has a menu item (invisible item OK)."
      .Add "    'If there is no menu item, the menu bar will be created at first at runtime."
      .Add "    'But, it is not fixed. In this case, the API drawn menu bar may disappear unexpectedly."
      .Add "    'And, API drawn menu will not be even created in IDE if the form has no menu item."
      .Add "   Call ShowAPIDrawMenuBar_" & FuncNameSuffix & "(frm)"
      .Add "    'Start subclassing to process menu messages."
      .Add "   Call BeginMenuBarPreviewSubclassing_" & FuncNameSuffix & "(frm.hWnd)"
      .Add "   frm.Show vbModal, OwnerForm"
      .Add "    'End subclassing before destroying the form completely."
      .Add "   Call EndMenuBarPreviewSubclassing_" & FuncNameSuffix
      .Add "   Set frm = Nothing"
      .Add vbNullString
      .Add "End Sub"
      .Add vbNullString
   
      .Add vbNullString
      .Add "Private Sub ShowAPIDrawMenuBar_" & FuncNameSuffix & "(ByVal frm As VB.Form)"
      .Add vbNullString
      .Add "   Dim hMenu As Long 'Handle to the main popup menu created."
      .Add "   Dim hPrevMenu As Long"
      .Add "   Dim hWnd As Long"
      .Add vbNullString
      .Add "   hWnd = frm.hWnd"
      .Add vbNullString
      .Add "    'Get the previous menu."
      .Add "   hPrevMenu = GetMenu(hWnd)"
      .Add vbNullString
      .Add "   hMenu = CreateMenu() 'Create a blank main menu."
      .Add "    'Create sub menus"
      .Add "   Call pCreateAPIPopupMenus_" & FuncNameSuffix & "(hMenu)"
      .Add "    'SetItem the menu for the form."
      .Add "   SetMenu hWnd, hMenu"
      .Add "    'Once changed the menus with SetMenu, we have to draw menubar."
      .Add "   DrawMenuBar hWnd"
      .Add vbNullString
      .Add "    'Destory the previous menu to free resources."
      .Add "   If hPrevMenu Then"
      .Add "      DestroyMenu hPrevMenu"
      .Add "   End If"
      .Add vbNullString
      .Add "End Sub"
      
      .Add vbNullString
      
      .Add "Private Sub BeginMenuBarPreviewSubclassing_" & FuncNameSuffix & "(ByVal hWnd As Long)"
      .Add "   m_hWnd_" & FuncNameSuffix & " = hWnd"
      .Add "   m_lpfnMenuBarOldWndProc_" & FuncNameSuffix & " = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndMenuBarPreviewProc_" & FuncNameSuffix & ")"
      .Add "End Sub"
      
      .Add vbNullString
      .Add "Private Sub EndMenuBarPreviewSubclassing_" & FuncNameSuffix & "()"
      .Add "   Call SetWindowLong(m_hWnd_" & FuncNameSuffix & ", GWL_WNDPROC, m_lpfnMenuBarOldWndProc_" & FuncNameSuffix & ")"
      .Add "End Sub"
      
      .Add vbNullString
      
      .Add "Private Function WndMenuBarPreviewProc_" & FuncNameSuffix & "(ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long"
      .Add "   Dim nCmdCode As Long, nCmdItemID As Long"
      .Add "   If uMsg = WM_COMMAND Then"
      .Add "      If lParam = 0 Then 'in case of menu, lParam is always 0."
      .Add "          'HiWord of wParam is command code source."
      .Add "          '0 = Menu, CommandButton, 1 = Accelerator, Other = Control"
      .Add "         If HiWord(wParam) = 0 Then 'if menu"
      .Add "             'LoWord of wParam is Menu Item ID"
      .Add "            nCmdItemID = LoWord(wParam) And &HFFFF& 'Fix to unsigned integer"
      .Add "             'Display message."
      .Add "            MsgBox ""The user clicked ["" & GetMenuCaption(GetMenu(hWnd), nCmdItemID) & ""]."" & _"
      .Add "                                                         vbCrLf & ""Menu Item ID="" & nCmdItemID"
      .Add "         End If"
      .Add "      End If"
      .Add "   End If"
      .Add "   WndMenuBarPreviewProc_" & FuncNameSuffix & " = CallWindowProc(m_lpfnMenuBarOldWndProc_" & FuncNameSuffix & ", hWnd, uMsg, wParam, lParam)"
      .Add "End Function"
      .Add vbNullString
      
      .Add vbCrLf
      .Add pCreatePreviewCode(colLinearMenus, FuncNameSuffix)
   End With
   
   GenerateAPIDrawMenuBarCode = JoinCollection(colText, vbCrLf, vbNullString, vbNullString)
End Function

Public Function CreateAPIPopupMenuCode() As String
   CreateAPIPopupMenuCode = GenerateAPIPopupMenuCode(Me)
End Function

Private Function GenerateAPIPopupMenuCode(ByVal colLinearMenus As VBMenus) As String

   Dim colText As New Collection
   Dim FuncNameSuffix As String
   
   FuncNameSuffix = Format$(Now(), "YYMMDDHHMMSS")
   With colText
      .Add "Public Sub ShowAPIPopupMenu_" & FuncNameSuffix & "(Optional ByVal hWndOwner As Long)"
      .Add vbNullString
      .Add "   Dim hMenu As Long 'Handle to the main popup menu created."
      .Add "   Dim nItemID As Long 'ID of the item that the user clicked."
      .Add "   "
      .Add "   hMenu = CreatePopupMenu() 'Create a blank main popup menu."
      .Add vbNullString
      .Add "    'Create sub menus"
      .Add "   Call pCreateAPIPopupMenus_" & FuncNameSuffix & "(hMenu)"
      .Add vbNullString
      .Add "    'hWndOwner is required for TrackPopupMenu (Used in PopupMenuEx)"
      .Add "    'If it is NULL (0), set it to the foregroundwindow."
      .Add "    'This is harmless, beacuse PopupMenuEx does not send notificaiton message"
      .Add "    'to the owner window and returns immediately with the clicked item's ID."
      .Add "   If hWndOwner = 0 Then"
      .Add "      hWndOwner = GetForegroundWindow()"
      .Add "   End If"
      .Add vbNullString
      .Add "   nItemID = PopupMenuEx(hMenu, hWndOwner)"
      .Add "    'If the user clicked a menu item"
      .Add "   If nItemID <> 0 Then"
      .Add "      'Process the message."
      .Add "      'You need to generate code to process the user's response."
      .Add "      MsgBox GetMenuCaption(hMenu, nItemID) & vbCrLf & ""ItemID="" & nItemID"
      .Add "   End If"
      .Add vbNullString
      .Add "    'Now finally destroy the menu created to free system resources."
      .Add "   Call DestroyMenu(hMenu)"
      .Add vbNullString
      .Add "End Sub"
      
      .Add vbCrLf
      .Add pCreatePreviewCode(colLinearMenus, FuncNameSuffix)
   End With
   
   GenerateAPIPopupMenuCode = JoinCollection(colText, vbCrLf, vbNullString, vbNullString)
End Function


Private Function pCreatePreviewCode( _
   ByVal colLinearMenus As VBMenus, _
   Optional ByVal FuncNameSuffix As String) As String
   
   'Create menu text to save. The given collection must be a structured one(i.e.,
   'must have parent-children relationship).
   
   'This precedure is recursive.
   Dim i As Long
   Dim hSubMenu As Long
   Dim colText As New Collection
   Dim nSubMaxLevel As Long
   
   With colLinearMenus
      If .Count = 0 Then
         Exit Function
      End If
      
      For i = 1 To .Count
         With .Item(i)
            'if parent,
            If .IsParent Then
               If nSubMaxLevel < .Level Then
                  nSubMaxLevel = .Level
               End If
               If .Level <= 1 Then
                  Call colText.Add(vbNullString)
               End If
               Call colText.Add("hSubMenu" & .Level & " = CreateSubMenu(" & _
                                 IIf(.Level <= 1, "hMenu", "hSubMenu" & .Level - 1) & _
                                    ", " & Quoto(.Caption) & ", " & 1000 + i & ")")
            Else
               If .Caption = "-" Then
                  Call colText.Add("Call AddSeparator(" & IIf(.Level <= 1, "hMenu", "hSubMenu" & .Level - 1) & ", " & 1000 + i & ")")
               Else
                  Call colText.Add("Call AddMenuItem(" & IIf(.Level <= 1, "hMenu", "hSubMenu" & .Level - 1) & ", " _
                                         & Quoto(.Caption) & ", " & 1000 + i & ", " & eShortCutDesc(.ShortcutIndex) & ")")
               End If
            End If
         End With
      Next
   End With
   
   For i = nSubMaxLevel To 1 Step -1
      Call colText.Add("Dim hSubMenu" & i & " As Long", , 1)
   Next i
   
   pCreatePreviewCode = "Private Sub pCreateAPIPopupMenus" & _
                                             IIf(LenB(FuncNameSuffix) > 0, "_" & FuncNameSuffix, vbNullString) & "(ByVal hMenu As Long)" & _
                                          vbCrLf & vbCrLf & _
                                          JoinCollection(colText, vbCrLf, vbSpace3) & _
                                          vbCrLf & vbCrLf & _
                                          "End Sub"
End Function


