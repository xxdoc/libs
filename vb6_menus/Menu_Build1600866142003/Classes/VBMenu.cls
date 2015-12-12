VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Created & released by KSY, 06/14/2003
'
'Children collection. Used to create a structured menu (that is, has parent-children relationship)
'Internally, used to create menu text.
Private m_colChildren As New VBMenus

'Properties for housekeeping
Private m_bIsParent As Boolean
Private m_strText As String
Private m_lLevel As Long
Private m_lPosition As Long
Private m_lLength As Long
Private m_strKey As String

'Properties for creating menu text.
Private m_strName As String
Private m_lIndex As Long
Private m_lHelpContextID As Long
Private m_enNegotiatePosition As eNegotiatePosition
Private m_bWindowList As Boolean
Private m_strShortCut As String
Private m_strCaption As String
Private m_bVisible As Boolean
Private m_bEnabled As Boolean
Private m_bChecked As Boolean
Private m_strTag As String
Public Property Get Tag() As String
   Tag = m_strTag
End Property
Public Property Let Tag(ByVal strTag As String)
   m_strTag = strTag
End Property



Private Sub Class_Initialize()
   'Set default properties.
   m_lIndex = -1
   m_enNegotiatePosition = vbNegoPosNone
   m_bVisible = True
   m_bEnabled = True
End Sub

Public Property Get Checked() As Boolean
   Checked = m_bChecked
End Property
Public Property Let Checked(ByVal bChecked As Boolean)
   m_bChecked = bChecked
End Property
Public Property Get Enabled() As Boolean
   Enabled = m_bEnabled
End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
   m_bEnabled = bEnabled
End Property
Public Property Get Visible() As Boolean
   Visible = m_bVisible
End Property
Public Property Let Visible(ByVal bVisible As Boolean)
   m_bVisible = bVisible
End Property
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
   Caption = m_strCaption
End Property
Public Property Let Caption(ByVal strCaption As String)
   m_strCaption = strCaption
End Property
Public Property Get ShortCut() As String
   ShortCut = m_strShortCut
End Property
Public Property Let ShortCut(ByVal strShortCut As String)
   m_strShortCut = strShortCut
End Property
Public Property Get ShortCutDesc() As String
   ShortCutDesc = GetShortCutDesc(ShortcutIndex)
End Property
Public Property Get ShortcutIndex() As eShortCut
   ShortcutIndex = GetShortCutIndex(ShortCut)
End Property
Public Property Let ShortcutIndex(ByVal lShortcutIndex As eShortCut)
   ShortCut = GetShortCut(lShortcutIndex)
End Property
Public Property Get WindowList() As Boolean
   WindowList = m_bWindowList
End Property
Public Property Let WindowList(ByVal bWindowList As Boolean)
   m_bWindowList = bWindowList
End Property
Public Property Get NegotiatePosition() As eNegotiatePosition
   NegotiatePosition = m_enNegotiatePosition
End Property
Public Property Let NegotiatePosition(ByVal enNegotiatePosition As eNegotiatePosition)
   m_enNegotiatePosition = enNegotiatePosition
End Property
Public Property Get HelpContextID() As Long
   HelpContextID = m_lHelpContextID
End Property
Public Property Let HelpContextID(ByVal lHelpContextID As Long)
   m_lHelpContextID = lHelpContextID
End Property
Public Property Get Index() As Long
   Index = m_lIndex
End Property
Public Property Let Index(ByVal lIndex As Long)
   m_lIndex = lIndex
End Property
Public Property Get Length() As Long
   Length = m_lLength
End Property
Public Property Let Length(ByVal lLength As Long)
   m_lLength = lLength
End Property
Public Property Get Position() As Long
   Position = m_lPosition
End Property
Public Property Let Position(ByVal lPosition As Long)
   m_lPosition = lPosition
End Property
Public Property Get Key() As String
   Key = m_strKey
End Property
Public Property Let Key(ByVal strKey As String)
   m_strKey = strKey
End Property
Public Property Get Level() As Long
   Level = m_lLevel
End Property
Public Property Let Level(ByVal lLevel As Long)
   m_lLevel = lLevel
End Property
Public Property Get Children() As VBMenus
   Set Children = m_colChildren
End Property
Public Property Set Children(ByVal objChildren As VBMenus)
   Set m_colChildren = objChildren
End Property
Public Property Get Name() As String
   Name = m_strName
End Property
Public Property Let Name(ByVal strName As String)
   m_strName = strName
End Property
Public Property Get IsParent() As Boolean
   IsParent = m_bIsParent
End Property
Public Property Let IsParent(ByVal bIsParent As Boolean)
   m_bIsParent = bIsParent
End Property
Public Property Get Text() As String
   Text = m_strText
End Property
Public Property Let Text(ByVal strText As String)
   m_strText = strText
End Property
Public Function PropertyName(ByVal Index As eVBMenuProperties) As String
   Select Case Index
   Case vbmp_Name:      PropertyName = ksName
   Case vbmp_IsParent:        PropertyName = ksIsParent
   Case vbmp_Key:       PropertyName = ksKey
   Case vbmp_Level:        PropertyName = ksLevel
   Case vbmp_ShortcutIndex:      PropertyName = ksShortcutIndex
   Case vbmp_ShortcutDesc:       PropertyName = ksShortcutDesc
   Case vbmp_Text:      PropertyName = ksText
   Case vbmp_Position:        PropertyName = ksPosition
   Case vbmp_Length:       PropertyName = ksLength

   Case vbmp_Caption:      PropertyName = ksCaption
   Case vbmp_Index:        PropertyName = ksIndex
   Case vbmp_Checked:      PropertyName = ksChecked
   Case vbmp_Enabled:      PropertyName = ksEnabled
   Case vbmp_Visible:      PropertyName = ksVisible
   Case vbmp_Shortcut:        PropertyName = ksShortcut
   Case vbmp_NegotiatePosition:        PropertyName = ksNegotiatePosition
   Case vbmp_HelpContextID:      PropertyName = ksHelpContextID
   Case vbmp_Tag:      PropertyName = ksTag
   Case vbmp_WindowList:      PropertyName = ksWindowList
   Case Else
   End Select
End Function
Public Property Get Properties(PropertyName As String) As Variant
   Select Case PropertyName
   Case ksCaption:      Properties = m_strCaption
   Case ksIndex:        Properties = m_lIndex
   Case ksChecked:      Properties = m_bChecked
   Case ksEnabled:      Properties = m_bEnabled
   Case ksVisible:      Properties = m_bVisible
   Case ksShortcut:        Properties = m_strShortCut
   Case ksNegotiatePosition:        Properties = m_enNegotiatePosition
   Case ksHelpContextID:      Properties = m_lHelpContextID
   Case ksTag:      Properties = m_strTag
   Case ksWindowList:        Properties = m_bWindowList
   
   Case ksName:      Properties = m_strName
   Case ksIsParent:        Properties = m_bIsParent
   Case ksKey:       Properties = m_strKey
   Case ksLevel:        Properties = m_lLevel
   Case ksShortcutIndex:        Properties = ShortcutIndex
   Case ksText:      Properties = m_strText
   Case ksPosition:       Properties = m_lPosition
   Case ksLength:       Properties = m_lLength
   Case ksShortcutDesc:        Properties = ShortCutDesc
   Case Else
   End Select
End Property
Public Property Let Properties(PropertyName As String, ByVal vntProperty As Variant)
   Select Case PropertyName
   Case ksCaption:      m_strCaption = vntProperty
   Case ksVisible:      m_bVisible = vntProperty
   Case ksEnabled:      m_bEnabled = vntProperty
   Case ksChecked:      m_bChecked = vntProperty
   Case ksText:      m_strText = vntProperty
   Case ksName:      m_strName = vntProperty
   Case ksIsParent:        m_bIsParent = vntProperty
   Case ksLevel:        m_lLevel = vntProperty
   Case ksKey:       m_strKey = vntProperty
   Case ksPosition:       m_lPosition = vntProperty
   Case ksLength:       m_lLength = vntProperty
   Case ksIndex:        m_lIndex = vntProperty
   Case ksHelpContextID:      m_lHelpContextID = vntProperty
   Case ksTag:      m_strTag = vntProperty
   Case ksNegotiatePosition:        m_enNegotiatePosition = vntProperty
   Case ksWindowList:        m_bWindowList = vntProperty
   Case ksShortcut:        m_strShortCut = vntProperty
   Case ksShortcutIndex:        ShortcutIndex = vntProperty
   Case Else
   End Select
End Property

'Copies to a new VBMenu. Needed to create a whole menu text with VBMenus collection.
Public Function Copy() As VBMenu
   Set Copy = New VBMenu
   With Copy
      .Caption = m_strCaption
      .Visible = m_bVisible
      .Enabled = m_bEnabled
      .Checked = m_bChecked
      .Length = m_lLength
      .Position = m_lPosition
      .Key = m_strKey
      .Level = m_lLevel
      .Name = m_strName
      .IsParent = m_bIsParent
      .Text = m_strText
      .Tag = m_strTag
      .ShortCut = m_strShortCut
      .WindowList = m_bWindowList
      .NegotiatePosition = m_enNegotiatePosition
      .HelpContextID = m_lHelpContextID
      .Index = m_lIndex
   End With
End Function

Public Function GetMenuText(Optional ByVal Kind As eVBMenuCreateKind = VBMC_DEFAULT) As String
   'Create menu text for this item.
   
   '      Begin VB.Menu mnuFileNew
   '         Caption         =   "New Menu"
   '         Checked         =   -1  'True
   '         Enabled         =   0   'False
   '         HelpContextID   =   1234
   '         Index           =   0
   '         Visible         =   0   'False
   '         WindowList      =   -1  'True
   '         NegotiatePosition=   1  '¿ÞÂÊ
   '      End
   
   Dim kstrSpace As String, kstrSpaceTop As String
   Dim strTemp As String
   
   kstrSpaceTop = Repeat(m_lLevel, vbSpace3)
   kstrSpace = kstrSpaceTop & kstrSpaceTop
   
   'if the end text is requested.
   If Kind = VBMC_END Then
      GetMenuText = kstrSpaceTop & "End"
      Exit Function
   End If
   
   'Heading
   GetMenuText = kstrSpaceTop & ksBegin_VB_Menu & vbSpace & m_strName
   
   'Properites.
   If m_lLevel < 1 Then
      GetMenuText = GetMenuText & vbCrLf & JoinProperties(vbSpace3)
   Else
      GetMenuText = GetMenuText & vbCrLf & JoinProperties(kstrSpace)
   End If
   
   'If not parent item, also returns the End text.
   If Not m_bIsParent Then
      GetMenuText = GetMenuText & (vbCrLf & kstrSpaceTop & "End")
   End If

   m_strText = GetMenuText
End Function

Private Function JoinProperties(PrefixBeforeName As String) As String
   
   'Collect the properties text to create menu text.
   
   Dim i As eVBMenuProperties
   Dim strItem As String, asItem() As String
   Dim iCount As Long
   
   For i = vbmp_PropDescMin To vbmp_PropDescMax
      strItem = GetPropertyText(i)
      If LenB(strItem) Then
         ReDim Preserve asItem(iCount)
         asItem(iCount) = PrefixBeforeName & strItem
         iCount = iCount + 1
      End If
   Next
   If iCount Then
      JoinProperties = Join(asItem, vbCrLf)
   End If
   
End Function

Private Function GetPropertyText(ByVal Index As eVBMenuProperties) As String
   'Returns property text to be saved as menu text.
   
   Const ksPropertyInfix As String = vbSpace9 & ksEqual & vbSpace3
   
   Select Case Index
   Case vbmp_Caption:
      GetPropertyText = ksCaption & ksPropertyInfix & vbQ & m_strCaption & vbQ
   Case vbmp_Index:
      If m_lIndex >= 0 Then
         GetPropertyText = ksIndex & ksPropertyInfix & m_lIndex
      End If
   Case vbmp_Checked:
      If m_bChecked Then
         GetPropertyText = ksChecked & ksPropertyInfix & CInt(m_bChecked) & vbSpace3 & vbSQ & ksTrue
      End If
   Case vbmp_Enabled:
      If Not m_bEnabled Then
         GetPropertyText = ksEnabled & ksPropertyInfix & CInt(m_bEnabled) & vbSpace3 & vbSQ & ksFalse
      End If
   Case vbmp_Visible
      If Not m_bVisible Then
         GetPropertyText = ksVisible & ksPropertyInfix & CInt(m_bVisible) & vbSpace3 & vbSQ & ksFalse
      End If
   Case vbmp_Shortcut:
      If Len(m_strShortCut) Then
         GetPropertyText = ksShortcut & ksPropertyInfix & m_strShortCut
      End If
   Case vbmp_NegotiatePosition:
      If m_enNegotiatePosition <> vbNegoPosNone Then
         GetPropertyText = ksNegotiatePosition & ksPropertyInfix & _
                     m_enNegotiatePosition & vbSpace3 & vbSQ & eNegotiatePositionDesc(m_enNegotiatePosition)
      End If
   Case vbmp_HelpContextID:
      If m_lHelpContextID > 0 Then
         GetPropertyText = ksHelpContextID & ksPropertyInfix & m_lHelpContextID
      End If
   Case vbmp_WindowList:
      If m_bWindowList Then
         GetPropertyText = ksWindowList & ksPropertyInfix & CInt(m_bWindowList) & vbSpace3 & vbSQ & ksTrue
      End If
   Case Else
   End Select
End Function

Private Function eNegotiatePositionDesc(ByVal Index As eNegotiatePosition) As String
   Select Case Index
   Case vbNegoPosNone:        eNegotiatePositionDesc = ksNone
   Case vbNegoPosLeft:        eNegotiatePositionDesc = ksLeft
   Case vbNegoPosCenter:      eNegotiatePositionDesc = ksCenter
   Case vbNegoPosRight:       eNegotiatePositionDesc = ksRight
   Case Else
   End Select
End Function


