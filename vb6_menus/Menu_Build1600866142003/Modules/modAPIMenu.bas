Attribute VB_Name = "modAPIMenu"
Option Explicit

'
' Created & released by KSY, 06/14/2003
'
#Const USE_TYPELIB_API13 = 0
#Const USE_TYPELIB_API2 = 0
#Const USE_APIVB2_MODULE = 0

#Const USE_ENGLISH = 1
'============Private Variables ============================
Private m_lpfnOldWndProc As Long 'Subclassed old window proc
Private m_lDragListBoxMessage As Long
Private m_objListBox As ListBox

'============ Public String Constants =======================
#If USE_ENGLISH = 0 Then
Public Const FF_ALL_FILE As String = "모든 파일(*.*)" & vbNullChar & "*.*" & vbNullChar
Public Const FF_VBF_FILE As String = "Visual Basic 폼/컨트롤 (*.frm;*.ctl;*.pag)" & vbNullChar & "*.frm;*.ctl;*.pag" & vbNullChar
Public Const FF_VBM_FILE As String = "Visual Basic 메뉴 템플레이트 (*.vbm)" & vbNullChar & "*.vbm" & vbNullChar
Public Const DLGTITLE_SAVEAS_FORM As String = "폼으로 저장하기"
Public Const DLGTITLE_SAVEAS_TEMPLATE As String = "메뉴 템플레이트로 저장하기"
Public Const DIALOG_FORM As String = "폼 열기"
Public Const DIALOG_IMPORTFORM As String = "폼에서 메뉴 가져오기"
Public Const DIALOG_TEMPLATE As String = "메뉴 템플레이트 열기"
Public Const DIALOG_IMPORTTEMPLATE As String = "메뉴 템플레이트 가져오기"
Public Const ksLeft As String = "왼쪽"
Public Const ksCenter As String = "가운데"
Public Const ksRight As String = "오른쪽"
Public Const ksNone As String = "없음"
#Else
Public Const FF_ALL_FILE As String = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
Public Const FF_VBF_FILE As String = "Visual Basic Form/Control (*.frm;*.ctl;*.pag)" & vbNullChar & "*.frm;*.ctl;*.pag" & vbNullChar
Public Const FF_VBM_FILE As String = "Visual Basic Menu Template (*.vbm)" & vbNullChar & "*.vbm" & vbNullChar
Public Const DLGTITLE_SAVEAS_FORM As String = "Save As Form"
Public Const DLGTITLE_SAVEAS_TEMPLATE As String = "Save As Menu Template"
Public Const DIALOG_FORM As String = "Open Menu from Form"
Public Const DIALOG_IMPORTFORM As String = "Import Menu from Form"
Public Const DIALOG_TEMPLATE As String = "Open Menu Template"
Public Const DIALOG_IMPORTTEMPLATE As String = "Import Menu Template"
Public Const ksLeft As String = "Left"
Public Const ksCenter As String = "Center"
Public Const ksRight As String = "Right"
Public Const ksNone As String = "None"
#End If

Public Const ksShortcutDesc As String = "ShortcutDesc"
Public Const ksAttribute_VB As String = "Attribute VB_"

Public Enum eVBMenuProperties
   vbmp_Name = 1
   vbmp_IsParent
   vbmp_Key
   vbmp_Level
   vbmp_ShortcutIndex
   vbmp_ShortcutDesc
   vbmp_Text
   vbmp_Position
   vbmp_Length

   vbmp_Caption
   vbmp_Shortcut
   vbmp_Checked
   vbmp_Enabled
   vbmp_HelpContextID
   vbmp_Index
   vbmp_Visible
   vbmp_WindowList
   vbmp_Tag
   vbmp_NegotiatePosition

   vbmp_PropDescMin = vbmp_Caption
   vbmp_PropDescMax = vbmp_NegotiatePosition
End Enum

Public Type ppOpenfilename
   nStructSize       As Long
   hWndOwner         As Long
   hInstance         As Long
   sFilter           As String
   sCustomFilter     As String
   nMaxCustFilter    As Long
   nFilterIndex      As Long
   sFile             As String
   nMaxFile          As Long
   sFileTitle        As String
   nMaxTitle         As Long
   sInitialDir       As String
   sDialogTitle      As String
   flags             As Long
   nFileOffset       As Integer
   nFileExtension    As Integer
   sDefFileExt       As String
   nCustData         As Long
   fnHook            As Long
   sTemplateName     As String
End Type

'======== User Defined Types =================='
'Used to save the menus to the existing form in correct position.
'
#If USE_TYPELIB_API13 = 0 Then

Public Type POSINFO
   Start As Long
   End As Long
   Length As Long
End Type
#End If

Public Type MODULEHEADPOSINFO 'Positions of module elements
   HeadingPos As POSINFO 'Heading such as VERSION 5.00
   ObjectsPos As POSINFO 'Object = {....}
   ModulePropPos As POSINFO 'Form properties in the case of form
   ModulePropEndPos As POSINFO 'End of form properties
   ModulePropAllPos As POSINFO 'from Begin VB.Form ..... End
   ControlsPos As POSINFO 'Controls section
   MenusPos As POSINFO 'Menus section
   AttributesPos As POSINFO 'Module Attributes section
   HeadPos As POSINFO 'Module head (total of above items) postion in the entire module
   'DeclarePos As POSINFO 'Maybe used later
   'BodyPos As POSINFO 'Maybe used later
   'Length As Long 'Maybe used later
End Type

Public Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (lpOpenfilename As ppOpenfilename) As Long


'========== If API TypeLib v1.3 is not used ===========================
#If USE_TYPELIB_API13 = 0 Then
'
Public Const ksAttribute_VB_Name As String = "Attribute VB_Name"
Public Const ksBegin As String = "Begin"
Public Const ksBegin_VB_Form As String = "Begin VB.Form"
Public Const ksBegin_VB_Menu As String = "Begin VB.Menu"
Public Const ksCaption As String = "Caption"
Public Const ksChecked As String = "Checked"
Public Const ksEnabled As String = "Enabled"
Public Const ksEnd As String = "End"
Public Const ksEqual  As String = "="
Public Const ksFalse As String = "False"
Public Const ksHelpContextID As String = "HelpContextID"
Public Const ksIndex As String = "Index"
Public Const ksIsParent As String = "IsParent"
Public Const ksKey As String = "Key"
Public Const ksLevel As String = "Level"
Public Const ksName As String = "Name"
Public Const ksNegotiatePosition As String = "NegotiatePosition"
Public Const ksObject As String = "Object"
Public Const ksOption_Explicit As String = "Option Explicit"
Public Const ksShortcut As String = "Shortcut"
Public Const ksShortcutIndex As String = "ShortcutIndex"
Public Const ksText As String = "Text"
Public Const ksTrue As String = "True"
Public Const ksVERSION_500 As String = "VERSION 5.00"
Public Const ksVisible As String = "Visible"
Public Const ksWindowList As String = "WindowList"
Public Const ksVBMenu As String = "VB.Menu"
Public Const ksPosition As String = "Position"
Public Const ksLength As String = "Length"
Public Const ksTag As String = "Tag"

Public Const vbQ As String = """"
Public Const vbDQ As String = vbQ & vbQ
Public Const vbLine As String = "-------------------------------------------------------------------------------"
Public Const vbSQ As String = "'"
Public Const vbSpace As String = " "
Public Const vbSpace2 As String = vbSpace & vbSpace
Public Const vbSpace3 As String = vbSpace2 & vbSpace
Public Const vbSpace6 As String = vbSpace3 & vbSpace3
Public Const vbSpace9 As String = vbSpace6 & vbSpace3

'======== API Constant Enumerations =================='
Public Enum eFileAttributes
   FILE_ATTRIBUTE_DIRECTORY = &H10
End Enum

Public Enum eFileHandle
   INVALID_HANDLE_VALUE = -1&
End Enum

Public Enum ehKeyNames
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_PERFORMANCE_DATA = &H80000004
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_DYN_DATA = &H80000006
End Enum

Public Enum eRegOption
   REG_OPTION_NON_VOLATILE = 0&  ' Key is preserved when system is rebooted
   REG_OPTION_VOLATILE = 1&  ' Key is not preserved when system is rebooted
End Enum

Public Enum eSecurityMask
   KEY_QUERY_VALUE = &H1&
   KEY_SET_VALUE = &H2&
   KEY_CREATE_SUB_KEY = &H4&
   KEY_ENUMERATE_SUB_KEYS = &H8&
   KEY_NOTIFY = &H10&
   KEY_CREATE_LINK = &H20&
   KEY_READ = &H20019  '((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
   KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
   KEY_EXECUTE = &H20019  '(KEY_READ)
   KEY_ALL_ACCESS = &HF003F  '((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

Public Enum eRegValueType
   REG_NONE = 0 'No defined value type.
   REG_SZ = &H1 'Unicode nul terminated String
   REG_EXPAND_SZ = &H2 'Unicode nul terminated String
   REG_BINARY = &H3 'Binary data in any form.
   REG_DWORD = &H4 '32-bit number
   REG_LINK = 6 'Symbolic Link (unicode)
   REG_MULTI_SZ = &H7 'Multiple Unicode strings
   REG_RESOURCE_LIST = 8 ' Resource list In the resource map
End Enum

Public Enum eGlobalAllocFlag
   GMEM_FIXED = &H0
   GMEM_MOVEABLE = &H2
   GMEM_ZEROINIT = &H40
   GPTR = &H40 '(GMEM_FIXED Or GMEM_ZEROINIT)
   GHND = &H42 '(GMEM_MOVEABLE Or GMEM_ZEROINIT)
End Enum

Public Enum ePenStyle
   PS_SOLID = 0
   PS_DASH = 1
   PS_DASHDOT = 3
   PS_DASHDOTDOT = 4
   PS_DOT = 2
   PS_NULL = 5
End Enum

Public Enum eClipboardFormat
   CF_BITMAP = 2
   CF_DIB = 8
   CF_HDROP = 15
   CF_TEXT = 1
   CF_UNICODETEXT = 13
End Enum
'Private Const WM_USER = &H400
Public Enum eDragListBoxNotification
   DL_BEGINDRAG = 1157  '(WM_USER + 133)
   DL_DRAGGING = 1158 '(WM_USER + 134)
   DL_DROPPED = 1159 '(WM_USER + 135)
   DL_CANCELDRAG = 1160 '(WM_USER + 136)
End Enum

Public Enum eDragListBoxDraggingReturn
   DL_CURSORSET = 0
   DL_STOPCURSOR = 1
   DL_COPYCURSOR = 2
   DL_MOVECURSOR = 3
End Enum

Public Enum eListMsg
   LB_GETITEMHEIGHT = &H1A1
   LB_GETITEMRECT = &H198
   LB_SETITEMHEIGHT = &H1A0
End Enum

Public Enum eSystemColor
   COLOR_BACKGROUND = 1         'Desktop.
   COLOR_WINDOW = 5             'CIconMenuWin background.
   COLOR_HIGHLIGHT = 13         'Item(s) selected in a control.
   COLOR_3DHIGHLIGHT = 20 'COLOR_BTNHIGHLIGHT
End Enum

Public Enum eVirtualKey
   VK_SHIFT = &H10& '= 16
   VK_CONTROL = &H11& '= 17
   VK_ALT = &H12& '= 18
End Enum 'eVirtualKey

Public Enum eMsg
   WM_DROPFILES = &H233&
   WM_COMMAND = &H111&
   WM_NCPAINT = &H85&
End Enum

Public Enum eBOOL
   bFALSE = 0&
   bTRUE = 1&
End Enum

Public Enum ERRSUCCESS
   ERR_FAILED = 0
End Enum

Public Enum eSubMenu
   IDM_FILE = 1000
   IDM_EDIT = 2000
   IDM_VIEW = 3000
   IDM_PROJECT = 4000
   IDM_FORMAT = 5000
   IDM_DEBUG = 6000
   IDM_QUERY = 7000
   IDM_RUN = 8000
   IDM_TOOL = 9000
   IDM_OPTION = 10000
   IDM_WINDOW = 11000
   IDM_HELP = 12000
   IDM_APP = 13000
End Enum

Public Enum eWindowLongIndex
   GWL_EXSTYLE = -20 'Sets a new extended window style. For more information, see CreateWindowEx.
   GWL_STYLE = -16 'Sets a new window style.
   GWL_WNDPROC = -4 'Sets a new address for the window procedure.
   GWL_HINSTANCE = -6 'Sets a new application instance handle.
   GWL_ID = -12 'Sets a new identifier of the window.
   GWL_USERDATA = -21 'Sets the user data associated with the window. This data is intended for use by the application that created the window.
   'Its value is initially zero.
   'The following values are also available when the hWnd parameter identifies a dialog box.
   DWL_DLGPROC = 4 'Sets the new pointer to the dialog box procedure.
   DWL_MSGRESULT = 0 'Sets the return value of a message processed in the dialog box procedure.
   DWL_USER = 8 'Sets new extra information that is private to the application, such as handles or pointers.
End Enum


Public Enum eMenuItemInfoMask
    MIIM_BITMAP = &H80&
    MIIM_CHECKMARKS = &H8&
    MIIM_DATA = &H20&
    MIIM_FTYPE = &H100&
    MIIM_ID = &H2&
    MIIM_STATE = &H1&
    MIIM_STRING = &H40&
    MIIM_SUBMENU = &H4&
    MIIM_TYPE = &H10&
End Enum

Public Enum eMenuItemInfoType
    MFT_BITMAP = &H4&
    MFT_MENUBARBREAK = &H20&
    MFT_MENUBREAK = &H40&
    MFT_OWNERDRAW = &H100&
    MFT_RADIOCHECK = &H200&
    MFT_RIGHTJUSTIFY = &H4000&
    MFT_RIGHTORDER = &H2000&
    MFT_SEPARATOR = &H800&
    MFT_STRING = &H0&
End Enum

Public Enum eMenuItemInfoState
    MFS_CHECKED = &H8&
    MFS_DEFAULT = &H1000&
    MFS_DISABLED = &H3&
    MFS_ENABLED = &H0&
    MFS_GRAYED = &H3&
    MFS_HILITE = &H80&
    MFS_UNCHECKED = &H0&
    MFS_UNHILITE = &H0&
End Enum

Public Enum eMenuItemInfoBmpItem
    HBMMENU_CALLBACK = &HFFFFFFFF
    HBMMENU_MBAR_CLOSE = &H5&
    HBMMENU_MBAR_CLOSE_D = &H6&
    HBMMENU_MBAR_MINIMIZE = &H3&
    HBMMENU_MBAR_MINIMIZE_D = &H7&
    HBMMENU_MBAR_RESTORE = &H2&
    HBMMENU_POPUP_CLOSE = &H8&
    HBMMENU_POPUP_MAXIMIZE = &HA&
    HBMMENU_POPUP_MINIMIZE = &HB&
    HBMMENU_POPUP_RESTORE = &H9&
    HBMMENU_SYSTEM = &H1&
End Enum

Public Enum eMenuPositionFlags
    MF_BY_COMMAND = &H0&
    MF_BY_POSITION = &H400&
End Enum

Public Enum eTrackPopupMenuFlags
    TPM_CENTERALIGN = &H4&
    TPM_LEFTALIGN = &H0&
    TPM_RIGHTALIGN = &H8&
    TPM_BOTTOMALIGN = &H20&
    TPM_TOPALIGN = &H0&
    TPM_VCENTERALIGN = &H10&
    TPM_NONOTIFY = &H80&
    TPM_RETURNCMD = &H100&
    TPM_LEFTBUTTON = &H0&
    TPM_RIGHTBUTTON = &H2&
    TPM_HORNEGANIMATION = &H800&
    TPM_HORPOSANIMATION = &H400&
    TPM_NOANIMATION = &H4000&
    TPM_VERNEGANIMATION = &H2000&
    TPM_VERPOSANIMATION = &H1000&
    TPM_HORIZONTAL = &H0&
    TPM_VERTICAL = &H40&
End Enum
'======== Custom Enumerations =====================
Public Enum eVBMenuValidate
   VBM_ERR_NONE = 0
   VBM_ERR_MENU_MUST_HAVE_NAME
   VBM_ERR_LEVEL_SKIP
   VBM_ERR_PARENT_CANNOT_HAVE_SHORTCUT
   VBM_ERR_PARENT_CANNOT_BE_CHECKED
   VBM_ERR_PARENT_CANNOT_BE_SEPARATOR
   VBM_ERR_NAME_CANNOT_BE_DUPLICATED
   VBM_ERR_INVALID_INDEX
   VBM_ERR_INDEX_MUST_BE_LARGER_THAN_PREVIOUS
   VBM_ERR_INDEXED_ITEMS_MUST_BE_SEQUENTIAL
   VBM_ERR_INDEXED_ITEMS_MUST_HAVE_SAME_LEVEL
   VBM_ERR_CHILD_MUST_HAVE_DIFFERENT_NAME
   VBM_ERR_PARENT_MUST_HAVE_DIFFERENT_NAME
   VBM_ERR_ONLY_ONE_WINDOW_LIST_ALLOWED
   VBM_ERR_ONLY_PARENT_CAN_HAVE_WINDOW_LIST
End Enum

Public Enum eNegotiatePosition
   vbNegoPosNone = 0
   vbNegoPosLeft = 1
   vbNegoPosCenter = 2
   vbNegoPosRight = 3
End Enum

Public Enum eVBMenuCreateKind
   VBMC_DEFAULT = 0
   VBMC_BEGIN = 0
   VBMC_END = 1
End Enum

Public Enum eShortCut
   SC_000_NONE = 0
   SC_CTRL_A = 1  'Shortcut = ^A
   SC_CTRL_B = 2  'Shortcut = ^B
   SC_CTRL_C = 3  'Shortcut = ^C
   SC_CTRL_D = 4  'Shortcut = ^D
   SC_CTRL_E = 5  'Shortcut = ^E
   SC_CTRL_F = 6  'Shortcut = ^F
   SC_CTRL_G = 7  'Shortcut = ^G
   SC_CTRL_H = 8  'Shortcut = ^H
   SC_CTRL_I = 9  'Shortcut = ^I
   SC_CTRL_J = 10  'Shortcut = ^J
   SC_CTRL_K = 11  'Shortcut = ^K
   SC_CTRL_L = 12  'Shortcut = ^L
   SC_CTRL_M = 13  'Shortcut = ^M
   SC_CTRL_N = 14  'Shortcut = ^N
   SC_CTRL_O = 15  'Shortcut = ^O
   SC_CTRL_P = 16  'Shortcut = ^P
   SC_CTRL_Q = 17  'Shortcut = ^Q
   SC_CTRL_R = 18  'Shortcut = ^R
   SC_CTRL_S = 19  'Shortcut = ^S
   SC_CTRL_T = 20  'Shortcut = ^T
   SC_CTRL_U = 21  'Shortcut = ^U
   SC_CTRL_V = 22  'Shortcut = ^V
   SC_CTRL_W = 23  'Shortcut = ^W
   SC_CTRL_X = 24  'Shortcut = ^X
   SC_CTRL_Y = 25  'Shortcut = ^Y
   SC_CTRL_Z = 26  'Shortcut = ^Z
   SC_F1 = 27  'Shortcut = {F1}
   SC_F2 = 28  'Shortcut = {F2}
   SC_F3 = 29  'Shortcut = {F3}
   SC_F4 = 30  'Shortcut = {F4}
   SC_F5 = 31  'Shortcut = {F5}
   SC_F6 = 32  'Shortcut = {F6}
   SC_F7 = 33  'Shortcut = {F7}
   SC_F8 = 34  'Shortcut = {F8}
   SC_F9 = 35  'Shortcut = {F9}
   SC_F10 = 36  'Shortcut = {F10}
   SC_F11 = 37  'Shortcut = {F11}
   SC_F12 = 38  'Shortcut = {F12}
   SC_CTRL_F1 = 39  'Shortcut = ^{F1}
   SC_CTRL_F2 = 40  'Shortcut = ^{F2}
   SC_CTRL_F3 = 41  'Shortcut = ^{F3}
   SC_CTRL_F4 = 42  'Shortcut = ^{F4}
   SC_CTRL_F5 = 43  'Shortcut = ^{F5}
   SC_CTRL_F6 = 44  'Shortcut = ^{F6}
   SC_CTRL_F7 = 45  'Shortcut = ^{F7}
   SC_CTRL_F8 = 46  'Shortcut = ^{F8}
   SC_CTRL_F9 = 47  'Shortcut = ^{F9}
   SC_CTRL_F10 = 48  'Shortcut = ^{F10}
   SC_CTRL_F11 = 49  'Shortcut = ^{F11}
   SC_CTRL_F12 = 50  'Shortcut = ^{F12}
   SC_SHIFT_F1 = 51  'Shortcut = +{F1}
   SC_SHIFT_F2 = 52  'Shortcut = +{F2}
   SC_SHIFT_F3 = 53  'Shortcut = +{F3}
   SC_SHIFT_F4 = 54  'Shortcut = +{F4}
   SC_SHIFT_F5 = 55  'Shortcut = +{F5}
   SC_SHIFT_F6 = 56  'Shortcut = +{F6}
   SC_SHIFT_F7 = 57  'Shortcut = +{F7}
   SC_SHIFT_F8 = 58  'Shortcut = +{F8}
   SC_SHIFT_F9 = 59  'Shortcut = +{F9}
   SC_SHIFT_F10 = 60  'Shortcut = +{F10}
   SC_SHIFT_F11 = 61  'Shortcut = +{F11}
   SC_SHIFT_F12 = 62  'Shortcut = +{F12}
   SC_CTRL_SHIFT_F1 = 63  'Shortcut = +^{F1}
   SC_CTRL_SHIFT_F2 = 64  'Shortcut = +^{F2}
   SC_CTRL_SHIFT_F3 = 65  'Shortcut = +^{F3}
   SC_CTRL_SHIFT_F4 = 66  'Shortcut = +^{F4}
   SC_CTRL_SHIFT_F5 = 67  'Shortcut = +^{F5}
   SC_CTRL_SHIFT_F6 = 68  'Shortcut = +^{F6}
   SC_CTRL_SHIFT_F7 = 69  'Shortcut = +^{F7}
   SC_CTRL_SHIFT_F8 = 70  'Shortcut = +^{F8}
   SC_CTRL_SHIFT_F9 = 71  'Shortcut = +^{F9}
   SC_CTRL_SHIFT_F10 = 72  'Shortcut = +^{F10}
   SC_CTRL_SHIFT_F11 = 73  'Shortcut = +^{F11}
   SC_CTRL_SHIFT_F12 = 74  'Shortcut = +^{F12}
   SC_CTRL_INSERT = 75  'Shortcut = ^{INSERT}
   SC_SHIFT_INSERT = 76  'Shortcut = +{INSERT}
   SC_DEL = 77  'Shortcut = {DEL}
   SC_SHIFT_DEL = 78  'Shortcut = +{DEL}
   SC_ALT_BKSP = 79  'Shortcut = %{BKSP}
End Enum

Public Enum eFileNameParts
   efpBaseName = 0
   efpExtension
   efpPath
   efpPathUnqualified
   efpName
   efpPathPlusBaseName
   efpPathBaseName
   efpDrive
   efpDriveQualified
   efpConvToLocalName
   efpConvToShortName
   efpConvToLongName
End Enum

Public Enum eColor
   clrBlack = 0 '&H0, Black, RGB(0, 0, 0), #000000, clrBlack = 0x0
   clrBlue = 16711680 '&HFF0000, Blue, RGB(0, 0, 255), #0000FF, clrBlue = 0xFF0000
   clrBrown = 13209&
   clrBrown1 = 2763429 '&H2A2AA5, RGB(165, 42, 42), #A52A2A, Unnamed color = 0x2A2AA5
   clrCrimson = 3937500 '&H3C14DC, RGB(220, 20, 60), #DC143C, Unnamed color = 0x3C14DC
   clrCyan = &HFFFF00
   clrGold = 52479
   clrGold1 = 55295 '&HD7FF, RGB(255, 215, 0), #FFD700, Unnamed color = 0xD7FF
   clrGreen = 32768 '&H8000, Green, RGB(0, 128, 0), #008000, clrGreen = 0x8000
   clrIndigo = 10040115
   clrIndigo1 = 8519755 '&H82004B, RGB(75, 0, 130), #4B0082, Unnamed color = 0x82004B
   clrIvory = 15794175 '&HF0FFFF, RGB(255, 255, 240), #FFFFF0, Unnamed color = 0xF0FFFF
   clrLime = 52377
   clrLime1 = 65280 '&HFF00, Bright Green, RGB(0, 255, 0), #00FF00, clrBrightGreen = 0xFF00
   clrMagneta = &HFF00FF
   clrOrange = 26367&
   clrOrange1 = 42495 '&HA5FF, RGB(255, 165, 0), #FFA500, Unnamed color = 0xA5FF
   clrPurple = 8388736 '&H800080, Violet, RGB(128, 0, 128), #800080, clrViolet = 0x800080
   clrRed = 255 '&HFF, Red, RGB(255, 0, 0), #FF0000, clrRed = 0xFF
   clrRose = 13408767
   clrSalmon = 7504122 '&H7280FA, RGB(250, 128, 114), #FA8072, Unnamed color = 0x7280FA
   clrSilver = 12632256 '&HC0C0C0, Gray 25%, RGB(192, 192, 192), #C0C0C0, clrGray25 = 0xC0C0C0
   clrViolet = 8388736
   clrViolet1 = 15631086 '&HEE82EE, RGB(238, 130, 238), #EE82EE, Unnamed color = 0xEE82EE
   clrWhite = 16777215 '&HFFFFFF, White, RGB(255, 255, 255), #FFFFFF, clrWhite = 0xFFFFFF
   clrYellow = 65535 '&HFFFF, Yellow, RGB(255, 255, 0), #FFFF00, clrYellow = 0xFFFF
End Enum

'======== API Types =================='

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As eFileAttributes
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName(259) As Byte
   cAlternate(13) As Byte
End Type


Public Type MENUITEMINFO
    cbSize As Long
    fMask As eMenuItemInfoMask
    fType As eMenuItemInfoType
    fState As eMenuItemInfoState
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeDataPtr As Long
    cch As Long
    hBitmap As eMenuItemInfoBmpItem
End Type

'=========== API Declares ================================
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As eRegValueType, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKeyName As ehKeyNames, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As eRegOption, ByVal samDesired As eSecurityMask, ByVal lpSecurityAttributes As Long, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As eBOOL)
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function GetFileNameFromBrowse Lib "shell32" Alias "#63" (ByVal hWndOwner As Long, ByVal lpstrFile As String, ByVal nMaxFile As Long, ByVal lpstrInitialDir As String, ByVal lpstrDefExt As String, ByVal lpstrFilter As String, ByVal lpstrDialogTitle As String) As eBOOL
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As eClipboardFormat) As eBOOL
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As eFileHandle
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal nBufferLength As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As eBOOL) As eBOOL
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As eGlobalAllocFlag, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As eClipboardFormat, ByVal hMem As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As eClipboardFormat) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal StringPtr As Long) As Long
Public Declare Sub PathUnquoteSpacesA Lib "shlwapi.dll" (ByVal lpsz As String)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As eWindowLongIndex, ByVal dwNewLong As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As ePenStyle, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As eSystemColor) As eColor
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As eVirtualKey) As Integer
Public Declare Function OleTranslateColor Lib "oleaut32" (ByVal clr As Long, ByVal hpal As Long, ByRef pcolorref As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As eBOOL, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal uIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal uFlag As eMenuPositionFlags) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal uFlags As eTrackPopupMenuFlags, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long

'========= End of Declares ===============================
#End If 'USE_TYPELIB_API13 = 0

'========== If API TypeLib v2 is not used ===========================
#If USE_TYPELIB_API2 = 0 Then
Public Const DRAGLISTMSGSTRING = "commctrl_DragListMsg"
Public Declare Function MakeDragList Lib "comctl32" (ByVal hWndListBox As Long) As Boolean
Public Declare Sub DragListDrawInsertIcon Lib "comctl32" Alias "DrawInsert" (ByVal hwndParent As Long, ByVal hWndListBox As Long, ByVal nItem As Long)
Public Declare Function LBItemFromPt Lib "comctl32" (ByVal hWndListBox As Long, ByVal x As Long, ByVal y As Long, ByVal bAutoScroll As eBOOL) As Long
Public Type DRAGLISTINFO
   uNotification As eDragListBoxNotification
   hWnd As Long
   ptCursor As POINTAPI
End Type
#End If 'USE_TYPELIB_API2 = 0 -- If API TypeLib v2 is not used ================


'=========== Code Begin ================================
Private Sub A000_CodeBegin()
   
End Sub

Public Sub BeginSubclassing()
   'frmMenuEditor Subclassing

   With frmMenuEditor
      'Set reference for later use.
      Set m_objListBox = .lstMenu

      'Start file drag & drop
      DragAcceptFiles .hWnd, bTRUE

      'Grow listbox item height for good visual effect during dragging
      ListGrowItemHeight m_objListBox, 1
      'Get the notification message for listbox dragging
      m_lDragListBoxMessage = RegisterWindowMessage(DRAGLISTMSGSTRING)
      'Make the listbox a drag listbox
      Call MakeDragList(m_objListBox.hWnd)

      'Start subclassing to process messages for file drag & drop, and listbox item draggging.
      m_lpfnOldWndProc = SetWindowLong(.hWnd, GWL_WNDPROC, AddressOf WndProc)
   End With 'FRMMENUEDITOR

End Sub

Private Sub SetMouseIcon(ByVal Index As Long)

   Static m_oIconCopy As StdPicture
   Static m_oIconMove As StdPicture

   With m_objListBox
      If ObjPtr(m_oIconCopy) = 0 Then
         Set m_oIconCopy = .DragIcon
      End If
      If ObjPtr(m_oIconMove) = 0 Then
         Set m_oIconMove = .MouseIcon
      End If

      If Index Then
         Set .MouseIcon = m_oIconCopy
      Else 'INDEX = FALSE
         Set .MouseIcon = m_oIconMove
      End If
      .MousePointer = vbCustom
   End With 'M_OBJLISTBOX

End Sub

Public Sub EndSubclassing()

   With frmMenuEditor
      'Stop file drag & drop
      DragAcceptFiles .hWnd, bFALSE
      'Stop subclassing
      Call SetWindowLong(.hWnd, GWL_WNDPROC, m_lpfnOldWndProc)
      m_lDragListBoxMessage = 0
      m_lpfnOldWndProc = 0
   End With 'FRMMENUEDITOR

End Sub

'Window Proc for file drag & drop & listbox item dragging.
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long

   Select Case uMsg
   Case WM_DROPFILES 'File drag & drop
      Call DropFilesOnDropFiles(wParam)
   Case m_lDragListBoxMessage 'Drag listbox message
      WndProc = MsgDragListProc(hWnd, wParam, lParam)
   Case Else
      WndProc = CallWindowProc(m_lpfnOldWndProc, hWnd, uMsg, wParam, lParam)
   End Select

End Function

Private Function MsgDragListProc(hWnd As Long, wParam As Long, lParam As Long) As Long

'Processes draglist box messages

   Static nIdxDragStartItem As Long
   Static nIdxPrevDragging As Long

   Dim lpDragListInfo As DRAGLISTINFO
   Dim nIdxDrop As Long 'Item index at which the user dropped the dragging item.
   Dim hDC As Long 'Device context for drawing indicating line.
   Dim nIdxCursor As Long 'Item under the cursor during dragging.
   Dim rcCursor As RECT, rcPrev As RECT 'Current and previous items' rect

   'Copy draglist info structure from the pointer (lParam)
   CopyMemory lpDragListInfo, ByVal lParam, Len(lpDragListInfo)
   Select Case lpDragListInfo.uNotification
   Case DL_BEGINDRAG 'The drag operation starts. Return False to cancel
      'Get the selected item
      With lpDragListInfo
         nIdxDragStartItem = LBItemFromPt(.hWnd, .ptCursor.x, .ptCursor.y, False)
      End With 'LPDRAGLISTINFO
      'Continue with the drag
      MsgDragListProc = 1

   Case DL_CANCELDRAG  'The drag was canceled
      With m_objListBox
         .Refresh
         .MousePointer = vbDefault
         .Parent.Cls
      End With 'M_OBJLISTBOX
      'Stop the drag
      MsgDragListProc = 0

   Case DL_DRAGGING  'The item is being dragged
      'Draw the insert icon
      With lpDragListInfo

         'Get the index of the item under the cursor.
         nIdxCursor = LBItemFromPt(.hWnd, .ptCursor.x, .ptCursor.y, True)
         'RaiseEvent Dragging(nIdxDragStartItem, nIdxCursor)

         'Get the rect of the previous and current items.
         'These rects are used to draw indicating line.
         Call SendMessage(.hWnd, LB_GETITEMRECT, nIdxPrevDragging, rcPrev)
         Call SendMessage(.hWnd, LB_GETITEMRECT, nIdxCursor, rcCursor)

         'Get the device context of the list box.
         hDC = GetDC(.hWnd)
         'Erase the indicating line for the previous item.
         'Top should be subtracted by one for good visual.
         'For this purpose, we have enlarged itemheight of the listbox by one,
         'when the app starts.
         With rcPrev
            Call DrawLineEx(hDC, .Left, .Top - 1, .Right, .Top - 1, PS_SOLID, 2, TranslateColor(m_objListBox.BackColor))
         End With 'RCPREV
         'Draw the indicating line for the current item.
         With rcCursor
            Call DrawLineEx(hDC, .Left, .Top - 1, .Right, .Top - 1, PS_SOLID, 2, clrGold)
         End With 'RCCURSOR
         'Release the device context
         ReleaseDC .hWnd, hDC
         'Save the current index.
         nIdxPrevDragging = nIdxCursor
         'Draw insert icon also for more visual effect. (Optional)
         'If m_bDrawInsertIcon Then
         Call DragListDrawInsertIcon(hWnd, .hWnd, nIdxCursor)
         'End If
      End With 'LPDRAGLISTINFO
      'Return one of:
      'DL_STOPCURSOR: Changes the cursor to stop
      'DL_COPYCURSOR: Changes the cursor to copy
      'DL_MOVECURSOR: Changes the cursor to move
      SetMouseIcon GetAsyncKeyState(VK_CONTROL)
      'm_objListBox.MousePointer = vbCustom
      MsgDragListProc = eDragListBoxDraggingReturn.DL_CURSORSET

   Case DL_DROPPED
      'Dim Cancel As Boolean
      With lpDragListInfo
         nIdxDrop = LBItemFromPt(.hWnd, .ptCursor.x, .ptCursor.y, True)
         If nIdxDrop <> nIdxDragStartItem Then
            'RaiseEvent DragFinish(nIdxDragStartItem, nIdxDrop, Cancel)
            'If Not Cancel Then
            frmMenuEditor.MoveNodes nIdxDragStartItem, nIdxDrop
            'frmMenuEditor.ListMoveTo nIdxDragStartItem, nIdxDrop
            'ListMoveTo m_objListBox, nIdxDrop
            'End If
         End If
      End With 'LPDRAGLISTINFO
      With m_objListBox
         'Erase the drag indicating line.
         .Refresh
         .MousePointer = vbDefault
         'Erase the drag indicating curosor.
         .Parent.Cls
      End With 'M_OBJLISTBOX
      MsgDragListProc = 0
   End Select

End Function

Public Sub DropFilesOnDropFiles(HDROP As Long)

   Dim sFilename As String
   Dim lFileCount As Long
   Dim asFiles() As String
   Dim i As Long

   'will return number of files dropped on the form
   sFilename = VBA.Space$(256)
   lFileCount = DragQueryFile(HDROP, -1, sFilename, Len(sFilename))

   If lFileCount > 0 Then
      ReDim asFiles(lFileCount - 1)
      For i = 0 To lFileCount - 1
         'sets filename to name of (i+1) th file
         'DragQueryFile wParam, i, filename,127
         DragQueryFile HDROP, i, sFilename, Len(sFilename)
         asFiles(i) = TrimNull(sFilename)
      Next ':(? 'Chr$(160)!!Repeat For-Variable: I

      frmMenuEditor.OpenFile OFT_DRAGDROP, asFiles(0)
      'm_IApp.OnDropFiles hWnd, lFileCount, asFiles
   End If

   DragFinish HDROP

End Sub

Public Function SelectFile( _
                           Optional ByVal hWndOwner As Long, _
                           Optional ByVal DialogTitle As String = "Select a file to open.", _
                           Optional sFilters As String, _
                           Optional ByVal InitalDir As String) As String
Attribute SelectFile.VB_Description = "Opens file open dialog. Filters must be delimitered with vbNullChar."

'Opens file open dialog. Filters must be delimitered with vbNullChar.

   SelectFile = String$(260, 0)
   If LenB(InitalDir) = 0 Then
      InitalDir = App.Path
   End If
   If GetFileNameFromBrowse(hWndOwner, SelectFile, 260, InitalDir, vbNullString, _
      sFilters, DialogTitle) Then
      SelectFile = TrimNull(SelectFile)
   Else 'NOT GETFILENAMEFROMBROWSE(HWNDOWNER,...
      SelectFile = vbNullString
   End If

End Function

'Show the file save dialog
Public Function SelectSaveFile( _
                               Optional ByVal hWndOwner As Long, _
                               Optional ByVal DialogTitle As String = "Select a file to save.", _
                               Optional sFilters As String, _
                               Optional ByVal InitalDir As String) As String
Attribute SelectSaveFile.VB_Description = "Opens file save dialog. Filters and names must be delimitered with vbNullChar."

'Opens file save dialog. Filters and names must be delimitered with vbNullChar.
   Dim OFN As ppOpenfilename

   If LenB(InitalDir) = 0 Then
      InitalDir = App.Path
   End If
   'populate the structure
   With OFN
      .nStructSize = Len(OFN)
      .hWndOwner = hWndOwner
      .sFilter = sFilters
      .nFilterIndex = 0
      .sFile = Space$(1024) & vbNullChar & vbNullChar
      .nMaxFile = Len(.sFile)
      .sDefFileExt = vbNullChar & vbNullChar
      .sFileTitle = vbNullChar & vbNullChar
      .nMaxTitle = Len(OFN.sFileTitle)
      .sInitialDir = InitalDir & vbNullChar
      .sDialogTitle = DialogTitle
      .flags = &H280006 'OFN_EXPLORER Or OFN_LONGNAMES Or
      'OFN_OVERWRITEPROMPT
      'Or OFN_HIDEREADONLY
      'call the API
      If GetSaveFileName(OFN) Then
         'Return the selected file.
         SelectSaveFile = TrimNull(.sFile)
      End If
   End With 'OFN

End Function

Public Sub ListSetTabStop(ByVal hWndList As Long, ParamArray Tabs() As Variant)
Attribute ListSetTabStop.VB_Description = "Sets up a listbox with TAB delimited columns."

'Sets up a listbox with TAB delimited columns.
   Const LB_SETTABSTOPS = &H192&
   Dim alTabs() As Long, i As Long

   ReDim alTabs(UBound(Tabs))
   For i = 0 To UBound(Tabs)
      alTabs(i) = Val(Tabs(i))
   Next i
   Call SendMessage(hWndList, LB_SETTABSTOPS, ByVal 0&, ByVal 0&)
   'Set the tabs.
   Call SendMessage(hWndList, LB_SETTABSTOPS, UBound(alTabs) + 1, alTabs(0))

End Sub

Public Function AssociateFileType( _
                                  ByVal AppPath As String, _
                                  ByVal AppEXEName As String, _
                                  ByVal AppFileTypeDesc As String, _
                                  ByVal Extension As String, _
                                  Optional useNotepadToEdit As Boolean = False) As Boolean
Attribute AssociateFileType.VB_Description = "Associates an extention with the specified application."

'Associates an extention with the specified application.

   Dim lRetVal As Long 'result
   Dim hKey As Long 'handle of open key
   Dim TheAppPath As String

   On Error GoTo e_Trap
   Dim sSetting As String

   If Mid$(AppPath, Len(AppPath) - 1, 1) = "\" Then
      TheAppPath = AppPath & AppEXEName & ".exe"
   Else 'NOT MID(APP.PATH,...'NOT MID$(APPPATH,...
      TheAppPath = AppPath & "\" & AppEXEName & ".exe"
   End If
   lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, AppFileTypeDesc, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
   sSetting = AppFileTypeDesc & Chr$(0)
   lRetVal = RegSetValueExString(hKey, "", 0&, REG_SZ, sSetting, Len(sSetting))

   RegCloseKey (hKey)
   lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, "." & LCase$(Extension), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
   sSetting = AppFileTypeDesc & Chr$(0)
   lRetVal = RegSetValueExString(hKey, "", 0&, REG_SZ, sSetting, Len(sSetting))

   RegCloseKey (hKey)
   lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, AppFileTypeDesc & "\shell\open\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
   sSetting = TheAppPath & " %1" & Chr$(0)
   lRetVal = RegSetValueExString(hKey, "", 0&, REG_SZ, sSetting, Len(sSetting))

   Call RegCloseKey(hKey)

   If useNotepadToEdit Then
      lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, AppFileTypeDesc & "\shell\edit\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
      sSetting = "notepad.exe %1" & Chr$(0)
      lRetVal = RegSetValueExString(hKey, "", 0&, REG_SZ, sSetting, Len(sSetting))
      Call RegCloseKey(hKey)
   End If
   lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, AppFileTypeDesc & "\DefaultIcon", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
   sSetting = TheAppPath & Chr$(0)
   lRetVal = RegSetValueExString(hKey, "", 0&, REG_SZ, sSetting, Len(sSetting))
   Call RegCloseKey(hKey)
   AssociateFileType = True

Exit Function

e_Trap:
   AssociateFileType = False

Exit Function

End Function

Public Function LoWord(ByVal dw As Long) As Integer
Attribute LoWord.VB_Description = "Extracts the LOWORD from a long."

'Extracts the LOWORD from a long.

   If dw And &H8000& Then
      LoWord = dw Or &HFFFF0000
   Else 'NOT DW...
      LoWord = dw And &HFFFF&
   End If

End Function

Public Function HiWord(ByVal dw As Long) As Integer
Attribute HiWord.VB_Description = "Extracts the HIWORD from a long."

'Extracts the HIWORD from a long.

   HiWord = (dw And &HFFFF0000) \ &H10000

End Function

Public Function TrimNull(strIn As String) As String
Attribute TrimNull.VB_Description = "Truncates the input string at first null. If no nulls, perform ordinary Trim."

   Dim nul As Long

   'Truncates the input string at first null. If no nulls, perform ordinary Trim.
   nul = VBA.Instr(strIn, vbNullChar)
   Select Case nul
   Case Is > 1
      TrimNull = VBA.Left$(strIn, nul - 1)
   Case 1
      TrimNull = ""
   Case 0
      TrimNull = VBA.Trim$(strIn)
   End Select

End Function

Public Function TrimCrLfTab(ByVal Text As String) As String
Attribute TrimCrLfTab.VB_Description = "Removes newline, tab, and nullchar from the input strijg, and then trims it."

'Removes newline, tab, and nullchar from the input strijg, and then trims it.

   TrimCrLfTab = VBA.Replace(Text, vbCr, vbSpace)
   TrimCrLfTab = VBA.Replace(TrimCrLfTab, vbLf, vbSpace)
   TrimCrLfTab = VBA.Replace(TrimCrLfTab, vbTab, vbSpace)
   TrimCrLfTab = VBA.Replace(TrimCrLfTab, vbNullChar, vbSpace)
   TrimCrLfTab = Trim$(TrimCrLfTab)

End Function

Public Function GetLongFileName(ByVal ShortFullName As String) As String
Attribute GetLongFileName.VB_Description = "Returns long file name from the given short file name. The file must exists."

'Returns long file name from the given short file name. The file must exists.

   GetLongFileName = ProperCaseDirectory(GetFileName(ShortFullName, efpPathUnqualified)) _
                     & "\" & VBA.Dir$(ShortFullName)

End Function

Public Function ProperCaseDirectory(ByVal DirPathIn As String) As String
Attribute ProperCaseDirectory.VB_Description = "Returns the proper cased directory name for an ucased or lcased directory name."

'Returns the proper cased directory name for an ucased or lcased directory name.

   Dim hSearch As Long
   Dim wfd As WIN32_FIND_DATA
   Dim PathOut As String
   Dim i As Long

   'Trim trailing backslash, unless root dir.
   If VBA.Right$(DirPathIn, 1) = "\" Then
      If VBA.Right$(DirPathIn, 2) <> ":\" Then
         DirPathIn = VBA.Left$(DirPathIn, Len(DirPathIn) - 1)
      Else 'NOT VBA.Right$(PATHIN,...'NOT VBA.RIGHT$(DIRPATHIN,...
         ProperCaseDirectory = VBA.UCase$(DirPathIn)
         Exit Function '>---> Bottom
      End If
   End If

   'Check for UNC share and return just that,
   'if that's all that's VBA.Left$ of DirPathIn.
   If VBA.Instr(DirPathIn, "\\") = 1 Then
      i = VBA.Instr(3, DirPathIn, "\")
      If i > 0 Then
         If VBA.Instr(i + 1, DirPathIn, "\") = 0 Then
            ProperCaseDirectory = DirPathIn
            Exit Function '>---> Bottom
         End If
      End If
   End If

   'Insure that path portion of string uses the
   'same case as the real pathname.
   If VBA.Instr(DirPathIn, "\") Then
      For i = Len(DirPathIn) To 1 Step -1
         If VBA.Mid$(DirPathIn, i, 1) = "\" Then
            'Found end of previous directory.
            'Recurse back up into path.
            PathOut = ProperCaseDirectory(VBA.Left$(DirPathIn, i - 1)) & "\"

            'Use FFF to proper-case current directory.
            hSearch = FindFirstFile(DirPathIn, wfd)
            If hSearch <> INVALID_HANDLE_VALUE Then
               Call FindClose(hSearch)
               If wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                  '-- Declare 버전: cFileName As String * 260
                  'ProperCaseDirectory = PathOut & TrimNull(wfd.cFileName)
                  '-- Type Library 버전: unsigned char cFileName[260];
                  ProperCaseDirectory = PathOut & TrimNull(StrConv(wfd.cFileName, vbUnicode))
               End If
            End If

            'Bail out of loop.
            Exit For '>---> Next
         End If
      Next i
   Else 'NOT VBA.Instr(PATHIN,...'NOT VBA.INSTR(DIRPATHIN,...
      'Just a drive letter and colon,
      'upper-case and return.
      ProperCaseDirectory = VBA.UCase$(DirPathIn)
   End If

End Function

Public Function GetFileName(ByVal FullName As String, ByVal ePortions As eFileNameParts) As String
Attribute GetFileName.VB_Description = "Returns the specifeid file name part(drive, base name, extension, path, etc.) from the given filename."

'Returns the specifeid file name part(drive, base name, extension, path, etc.) from the given filename.

'FullName = "D:\AdvVB\project\vbAdvanced_comp.vbp"
'efpBaseName = vbAdvanced_comp
'efpExtension = vbp
'efpPath=D:\AdvVB\project\
'efpPathUnqualified=D:\AdvVB\project
'efpName = vbAdvanced_comp.vbp
'efpPathPlusBaseName=D:\AdvVB\project\vbAdvanced_comp
'efpPathBaseName = project
'efpDrive = D
'efpDriveQualified = D:

   Dim lFirstPeriod As Long, lFirstBackSlash As Long
   Dim strQualifiedPath As String, strBaseName As String, sExt As String
   Dim sRet As String, sDrive As String

   Select Case ePortions
   Case efpDriveQualified, efpDrive
      lFirstPeriod = VBA.InStrRev(FullName, ":")
      If lFirstPeriod Then
         If ePortions = efpDriveQualified Then
            sDrive = VBA.Left$(FullName, lFirstPeriod)
         Else 'NOT EPORTIONS...
            If lFirstPeriod > 1 Then
               sDrive = VBA.Left$(FullName, lFirstPeriod - 1)
            End If
         End If
         lFirstBackSlash = VBA.InStrRev(sDrive, "/")
         If lFirstBackSlash = 0 Then
            lFirstBackSlash = VBA.InStrRev(sDrive, "\")
         End If
         If lFirstBackSlash Then
            sDrive = VBA.Mid$(sDrive, lFirstBackSlash + 1)
         End If
         GetFileName = sDrive
      End If
      Exit Function '>---> Bottom
   Case efpConvToLocalName
      lFirstPeriod = VBA.InStrRev(FullName, ":")
      If lFirstPeriod Then
         sDrive = VBA.Left$(FullName, lFirstPeriod)
         lFirstBackSlash = VBA.InStrRev(sDrive, "/")
         If lFirstBackSlash = 0 Then
            lFirstBackSlash = VBA.InStrRev(sDrive, "\")
         End If
         If lFirstBackSlash Then
            sDrive = VBA.Mid$(sDrive, lFirstBackSlash + 1)
         End If
         FullName = sDrive & VBA.Mid$(FullName, lFirstPeriod + 1)
      End If
      GetFileName = VBA.Replace(FullName, "/", "\")
      lFirstPeriod = VBA.Instr(1, GetFileName, "\\")
      If lFirstPeriod Then
         GetFileName = VBA.Mid$(GetFileName, lFirstPeriod + 2)
      End If
      Exit Function '>---> Bottom

   Case efpConvToShortName
      sRet = VBA.Space$(1024)
      GetShortPathName FullName, sRet, Len(sRet)
      GetFileName = TrimNull(sRet)
      Exit Function '>---> Bottom

   Case efpConvToLongName
      GetFileName = ProperCaseDirectory(GetFileName(FullName, efpPath)) _
                    & "\" & VBA.Dir$(FullName)
      Exit Function '>---> Bottom
   End Select

   '**** File name parts

   lFirstPeriod = VBA.InStrRev(FullName, ".")
   lFirstBackSlash = VBA.InStrRev(FullName, "\")
   If lFirstBackSlash = 0 Then
      lFirstBackSlash = VBA.InStrRev(FullName, "/")
   End If

   If lFirstBackSlash > 0 Then
      strQualifiedPath = VBA.Left$(FullName, lFirstBackSlash)
   End If

   If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
      sExt = VBA.Mid$(FullName, lFirstPeriod + 1)
      strBaseName = VBA.Mid$(FullName, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
   Else 'NOT LFIRSTPERIOD...
      strBaseName = VBA.Mid$(FullName, lFirstBackSlash + 1)
   End If

   Select Case ePortions
   Case efpBaseName
      GetFileName = strBaseName
   Case efpExtension
      GetFileName = sExt
   Case efpPath
      GetFileName = strQualifiedPath
   Case efpPathUnqualified
      If Len(strQualifiedPath) Then
         GetFileName = VBA.Left$(strQualifiedPath, Len(strQualifiedPath) - 1)
      End If
   Case efpName
      If Len(sExt) Then
         GetFileName = strBaseName & "." & sExt
      Else 'LEN(SEXT) = FALSE
         GetFileName = strBaseName
      End If
   Case efpPathPlusBaseName
      'If Len(sExt) Then
      GetFileName = strQualifiedPath & strBaseName
      'Else
      '   GetFileName = strQualifiedPath & strBaseName
      'End If
   Case efpPathBaseName
      If Len(strQualifiedPath) Then
         FullName = VBA.Left$(strQualifiedPath, Len(strQualifiedPath) - 1)
         GetFileName = GetFileName(FullName, efpBaseName)
      End If
   End Select

End Function

Public Function FileExists(FullName As String) As Boolean
Attribute FileExists.VB_Description = "Checks whether the given file exists."

'Checks whether the given file exists.
   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long

   If LenB(FullName) Then
      hFile = FindFirstFile(FullName, wfd)
      FileExists = hFile <> INVALID_HANDLE_VALUE And _
                   ((wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
      Call FindClose(hFile)
   End If

End Function

Public Function Repeat(Count As Long, Optional Text As String = vbSpace) As String
Attribute Repeat.VB_Description = "Repeats the given string as many as the count. if the count is less then 1, returns vbNullstring."

'Repeats the given string as many as the count. if the count is less then 1, returns vbNullstring.
   Dim i As Long

   For i = 1 To Count
      Repeat = Repeat & Text
   Next i

End Function

#If USE_APIVB2_MODULE = 0 Then
Public Sub ListMoveUp(ByVal ListBox As ListBox)
Attribute ListMoveUp.VB_Description = "Moves the selected item up by one on the list box, and sets list index."

'Moves the selected item up by one on the list box, and sets list index.

   With ListBox
      If .ListIndex > 0 Then
         Call ListSwap(ListBox, .ListIndex - 1, .ListIndex)
         .ListIndex = .ListIndex - 1
      End If
   End With 'LISTBOX

End Sub

Public Sub ListMoveDown(ByVal ListBox As ListBox)
Attribute ListMoveDown.VB_Description = "Moves the selected item down by one on the list box, and sets list index."

'Moves the selected item down by one on the list box, and sets list index.

   With ListBox
      If .ListIndex < .ListCount - 1 Then
         Call ListSwap(ListBox, .ListIndex + 1, .ListIndex)
         .ListIndex = .ListIndex + 1
      End If
   End With 'LISTBOX

End Sub
#End If

Public Sub ListSwap(ByVal ListBox As ListBox, ByVal Index1 As Long, ByVal Index2 As Long)
Attribute ListSwap.VB_Description = "Swaps the two list items."

'Swaps the two list items.
   Dim strTemp As String

   With ListBox
      strTemp = .List(Index1)
      .List(Index1) = .List(Index2)
      .List(Index2) = strTemp
   End With 'LISTBOX

End Sub

Public Sub ListGrowItemHeight(ByVal ListBox As ListBox, Optional dY As Long = 1)

'Grow item height bye one
   Dim nItemHeigth As Long

   With ListBox
      nItemHeigth = SendMessage(.hWnd, LB_GETITEMHEIGHT, 0&, ByVal 0&)
      If nItemHeigth <> -1 Then
         Call SendMessage(.hWnd, LB_SETITEMHEIGHT, 0&, ByVal nItemHeigth + dY)
      End If
   End With 'LISTBOX

End Sub

Public Sub ListMoveTo(ByVal ListBox As ListBox, ByVal FromIdx As Long, ByVal ToIdx As Long)

'Moves the selected item of the given ListBox to the given index.
'Called from the subclassing procedure when an item on the ListBox is dropped.
   Dim ItemData As Long, ListText As String, oMenu As VBMenu
   Dim bCtrlKeyPressed As Boolean, bShiftKeyPressed As Boolean

   With ListBox
      If FromIdx = ToIdx Or FromIdx < 0 And ToIdx < 0 _
         Or ToIdx > .ListCount - 1 And FromIdx > .ListCount - 1 Then
         Exit Sub '>---> Bottom
      End If

      bCtrlKeyPressed = GetAsyncKeyState(VK_CONTROL)
      bShiftKeyPressed = GetAsyncKeyState(VK_SHIFT)

      ListText = .List(FromIdx)
      ItemData = .ItemData(FromIdx)
      'Set oMenu = m_VBMenus(FromIdx + 1)
      .AddItem ListText, ToIdx
      .ItemData(.NewIndex) = ItemData
      'm_VBMenus.Add oMenu, , ToIdx + 1

      If ToIdx < FromIdx Then
         If Not bCtrlKeyPressed Then
            .RemoveItem FromIdx + 1
            'm_VBMenus.Remove FromIdx + 2
         End If
         .ListIndex = ToIdx
      Else 'NOT TOIDX...
         If Not bCtrlKeyPressed Then
            .RemoveItem FromIdx
            'm_VBMenus.Remove FromIdx + 1
            .ListIndex = IIf(ToIdx = 0, 0, ToIdx - 1)
         Else 'NOT NOT...
            .ListIndex = IIf(ToIdx = 0, 0, ToIdx)
         End If
      End If
      'm_VBMenus.SetPopupProperties
   End With 'LISTBOX

End Sub

Public Function GetFileText(ByVal strFilePathName As String) As String
Attribute GetFileText.VB_Description = "Returns the content of text file if exists."

'Returns the content of text file if it exists.

   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   Dim bFileExists As Boolean
   
   On Error GoTo Bye
   'Check whether the file exists.
   hFile = FindFirstFile(strFilePathName, wfd)
   bFileExists = hFile <> INVALID_HANDLE_VALUE And _
                 ((wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
   Call FindClose(hFile)

   'If exists, read it.
   If bFileExists Then
      Dim Buffer() As Byte
      Dim lFileLen As Long
      lFileLen = VBA.FileLen(strFilePathName)
      If lFileLen = 0 Then
         Exit Function
      End If
      ReDim Buffer(lFileLen - 1)
      Open strFilePathName For Binary As #1  'Source
      Get #1, , Buffer
      Close
      GetFileText = VBA.StrConv(Buffer, vbUnicode)
   End If
Bye:
End Function

Public Function VBReplaceMenuText(ByVal FileName As String, _
                                  NewMenuText As String, _
                                  Optional bMakeBackup As Boolean, _
                                  Optional bInsertIfMenuNotExists As Boolean = True) As Boolean
Attribute VBReplaceMenuText.VB_Description = "Replaces the existing menu section with a new one."

'Replaces the existing menu section with a new one.

   If Not FileExists(FileName) Then
      Exit Function '>---> Bottom
   End If

   Dim strText As String
   Dim lStartPos As Long, lEndPos As Long, bDoSave As Boolean

   strText = GetFileText(FileName)
   With VBGetModuleHeadPositions(strText)
      If .MenusPos.Start > 1 Then
         With .MenusPos
            lStartPos = .Start
            lEndPos = .End
            ''Debug.Assert .Length > 0
            bDoSave = True
         End With '.MENUSPOS
      ElseIf .ModulePropEndPos.Start > 1 Then 'NOT .MENUSPOS.START...
         With .ModulePropEndPos
            lStartPos = .Start
            lEndPos = .Start - 1
            ''Debug.Assert .Length > 0
            bDoSave = True
         End With '.MODULEPROPENDPOS
      End If
   End With 'VBGETMODULEHEADPOSITIONS(STRTEXT)

   If bDoSave Then
      If bMakeBackup Then
         BackupFile FileName
      End If
      VBReplaceMenuText = Save(FileName, _
                          Mid$(strText, 1, lStartPos - 1) & NewMenuText & vbCrLf & _
                          Mid$(strText, lEndPos + 1))
   End If

End Function

Public Function BackupFile(FileName As String, Optional bFailExists As Boolean) As Boolean

'Backup a file.

   BackupFile = CopyFile(FileName, _
                GetFileName(FileName, efpPathPlusBaseName) & ".bak." & _
                GetFileName(FileName, efpExtension), Abs(bFailExists))

End Function

Public Function Save( _
                     ByVal FilePathName As String, _
                     Text As String, _
                     Optional Append As Boolean = False) As Boolean

'Save a text file.

   On Error GoTo Error_Handle

   Dim lngFileNum As Long
   lngFileNum = VBA.FreeFile
   If Append Then
      Open FilePathName For Append As #lngFileNum
   Else 'APPEND = FALSE
      Open FilePathName For Output As #lngFileNum
   End If
   Print #lngFileNum, Text
   Close #lngFileNum
   Save = True

Exit Function

Error_Handle:

End Function

Public Function SetClipText(ByVal NewText As String, ByVal uFormat As Long) As Boolean

   Dim hGlobal As Long
   Dim pGlobal As Long
   Dim Buffer() As Byte
   Dim lRet As Long

   'Try to set text onto clipboard.
   lRet = OpenClipboard(0&)

   If lRet Then
      'Convert data to ANSI byte array.
      Buffer = StrConv(NewText & vbNullChar, vbFromUnicode)
      'Allocate enough memory for buffer.
      hGlobal = GlobalAlloc(GHND, UBound(Buffer) + 1)
      If hGlobal Then
         'Copy data to alloc'd memory.
         pGlobal = GlobalLock(hGlobal)
         Call CopyMemory(ByVal pGlobal, Buffer(0), UBound(Buffer) + 1)
         Call GlobalUnlock(hGlobal)
         'Hand data off to clipboard
         Call EmptyClipboard
         SetClipText = CBool(SetClipboardData(uFormat, hGlobal))
      End If
   End If

   Call CloseClipboard

End Function

Public Function GetClipText(ByVal uFormat As eClipboardFormat) As String

'Get text from the clipboard.
'VB clipboard object doesn't support custom format. (<-- Is correct?)
'Anyway, we will this procedure for custom format.

   Dim hGlobal As Long
   Dim pGlobal As Long
   Dim nLen As Long
   Dim lRet As Long

   'First, open the clipboard to lock it before checking the format is available.
   'to avoid possible problems due to Windows multi-tasking.
   lRet = OpenClipboard(0&)
   If lRet Then
      'Check the format is available.
      If IsClipboardFormatAvailable(uFormat) Then
         'Grab text from clipboard, if available. (i.e., get the handle to the clipboard data.)
         hGlobal = GetClipboardData(uFormat)
         If hGlobal Then
            'Get the memory pointer to the clipbaord date.
            pGlobal = GlobalLock(hGlobal)
            'Now, get the length of the data.
            'COMMENT: for this purpose, lstrlenA's StringPtr argument is Long, not String.
            '                  Usually, String is used.
            nLen = lstrlenA(ByVal pGlobal)
            If nLen Then
               'Prepare buffer.
               ReDim Buffer(0 To (nLen - 1)) As Byte
               'Copy the text to the buffer.
               CopyMemory Buffer(0), ByVal pGlobal, nLen
               'Convert to the Unicode text.
               GetClipText = StrConv(Buffer, vbUnicode)
            End If
            'Release the handle to the clipboard data.
            Call GlobalUnlock(hGlobal)
         End If
      End If
   End If
   'Nofity to Widnows that we have finsihed working with the clipboard.
   Call CloseClipboard

End Function

Public Function Between( _
                        ByVal Source As String, _
                        ByVal FirstTarget As String, _
                        ByVal SecondTarget As String, _
                        Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                        Optional FirstTarget_StartPos As Long, _
                        Optional SecondTarget_EndPos As Long, _
                        Optional ByVal bReturnWholeTextIfNotFound As Boolean, _
                        Optional ByRef retError As Long) As String

'Returns the text between the two targets.

   On Error GoTo Err_Handler

   FirstTarget_StartPos = InStr(1, Source, FirstTarget, Compare)
   If FirstTarget_StartPos > 0 Then
      SecondTarget_EndPos = VBA.Instr(FirstTarget_StartPos + Len(FirstTarget), _
                            Source, SecondTarget, Compare)
      If SecondTarget_EndPos > 0 Then
         SecondTarget_EndPos = SecondTarget_EndPos + Len(SecondTarget) - 1
         Source = After(Source, FirstTarget, Compare)
         Between = Before(Source, SecondTarget, Compare)
      Else 'NOT SECONDTARGET_ENDPOS...
         If bReturnWholeTextIfNotFound Then
            Between = Source
         Else 'BRETURNWHOLETEXTIFNOTFOUND = FALSE
            Between = ""
         End If
      End If
   Else 'NOT FIRSTTARGET_STARTPOS...
      If bReturnWholeTextIfNotFound Then
         Between = Source
      Else 'BRETURNWHOLETEXTIFNOTFOUND = FALSE
         Between = ""
      End If
   End If

   retError = 0

Exit Function

Err_Handler:
   retError = -1

End Function

Public Function VBGetModuleHead(Text As String) As String

'Returns the module head text.
'To identify head text, uses the delimiter "Attribute VB_"

   Dim pos As Long, pos1 As Long

   pos = InStrRev(Text, ksAttribute_VB)
   If pos <= 0 Then
      VBGetModuleHead = Text
   Else 'NOT POS...
      pos1 = InStr(pos, Text, vbCrLf)
      If pos1 > 0 Then
         VBGetModuleHead = Mid$(Text, 1, pos1 + Len(vbCrLf))
      Else 'NOT POS1...
         VBGetModuleHead = Text
      End If
   End If

End Function

Public Function VBGetModuleHeadPositions(Text As String) As MODULEHEADPOSINFO

'Returns the positions of the elements of module head,
'such as Attributes section, Menu section, Form properties section, Controls section...

'Used to save the created menu text to the existing file.

   Dim pos As Long, pos1 As Long
   Dim strHead As String

   With VBGetModuleHeadPositions
      strHead = VBGetModuleHead(Text)

      With .HeadPos
         .Start = 1
         .End = Len(strHead)
         .Length = .End - .Start + 1
         'Debug.Assert .Length > 0
      End With '.HEADPOS

      With .AttributesPos
         .Start = InStr(1, strHead, ksAttribute_VB)
         If .Start Then
            pos = InStrRev(strHead, ksAttribute_VB)
            If pos > .Start Then
               pos1 = InStr(pos, strHead, vbCrLf)
            Else 'NOT POS...
               pos1 = InStr(.Start, strHead, vbCrLf)
            End If
            If pos1 > 0 Then
               .End = pos1 - 1
            Else 'NOT POS1...
               .End = Len(strHead)
            End If
            .Length = .End - .Start + 1
            'Debug.Assert .Length > 0
         End If
      End With '.ATTRIBUTESPOS

      pos = InStr(1, strHead, ksBegin & vbSpace & "VB.")
      If pos Then
         With .ModulePropPos
            .Start = pos
            pos = InStr(.Start + Len(ksBegin & vbSpace & "VB.") + 1, strHead, ksBegin & vbSpace)
            If pos > .Start Then
               pos1 = InStrRev(strHead, vbCrLf, pos) - 1
               If pos1 > 0 Then
                  .End = pos1
               Else 'NOT POS1...
                  .End = pos
               End If
            Else 'NOT POS...
               pos = InStr(.Start + 1, strHead, vbCrLf & ksEnd)
               'Debug.Assert pos > 0
               If pos Then
                  .End = pos - 1
               Else 'POS = FALSE
                  .End = InStr(.Start + 1, strHead, ksEnd) - 1
               End If
            End If
            .Length = .End - .Start + 1
            'Debug.Assert .Length > 0
         End With '.MODULEPROPPOS
      End If

      If .ModulePropPos.Start > 0 Then
         .ModulePropAllPos.Start = .ModulePropPos.Start
         If .AttributesPos.Start > 0 Then
            pos = InStrRev(strHead, ksEnd, .AttributesPos.Start - 1)
         Else 'NOT .ATTRIBUTESPOS.START...
            pos = InStrRev(strHead, ksEnd)
         End If
         With .ModulePropEndPos
            .Start = pos
            .Length = Len(ksEnd)
            .End = .Start + .Length - 1
            'Debug.Assert .Length > 0
         End With '.MODULEPROPENDPOS
         With .ModulePropAllPos
            If pos > 0 Then
               .End = pos + Len(ksEnd) - 1
            Else 'NOT POS...
               .End = Len(strHead)
            End If
            .Length = .End - .Start + 1
            'Debug.Assert .Length > 0
         End With '.MODULEPROPALLPOS

         pos = InStrRev(strHead, vbCrLf, .ModulePropPos.Start)
         If pos > 1 Then
            pos1 = InStr(1, strHead, ksObject)
            If pos1 < pos Then
               With .ObjectsPos
                  .Start = pos1
                  .End = pos - 1
                  .Length = .End - .Start + 1
                  'Debug.Assert .Length >= 0
               End With '.OBJECTSPOS
            End If
         End If

         pos = InStr(1, strHead, ksBegin_VB_Menu)
         If pos > 0 Then
            pos1 = InStrRev(strHead, vbCrLf, pos)
            If pos1 > 0 Then
               .MenusPos.Start = pos1 + Len(vbCrLf)
            Else 'NOT POS1...
               .MenusPos.Start = pos
            End If
            pos = InStrRev(strHead, ksEnd, .ModulePropAllPos.End - Len(ksEnd))
            With .MenusPos
               .End = pos + Len(ksEnd) - 1
               .Length = .End - .Start + 1
               'Debug.Assert .Length > 0
            End With '.MENUSPOS
         End If

         If .MenusPos.Start > 0 Then
            pos = InStr(.ModulePropPos.End + 1, strHead, ksBegin & vbSpace)
            If pos Then
               pos1 = InStrRev(strHead, vbCrLf, pos) 'Control Pos Start
               If pos1 > .ModulePropPos.End Then
                  pos1 = pos1 + Len(vbCrLf) 'Control Pos Start
               Else 'NOT POS1...
                  pos1 = pos
               End If
               pos = InStrRev(strHead, ksEnd, .MenusPos.Start)
               If pos > pos1 Then
                  With .ControlsPos
                     .Start = pos1
                     .End = pos + Len(ksEnd) - 1
                     .Length = .End - .Start + 1
                     'Debug.Assert .Length > 0
                  End With '.CONTROLSPOS
               End If
            End If
         Else 'NOT .MENUSPOS.START...
            pos = InStr(.ModulePropPos.End + 1, strHead, ksBegin & vbSpace)
            If pos Then
               pos1 = InStrRev(strHead, vbCrLf, pos)
               If pos1 > .ModulePropPos.End Then
                  .ControlsPos.Start = pos1
               Else 'NOT POS1...
                  .ControlsPos.Start = pos
               End If
               pos = InStrRev(strHead, ksEnd, .ModulePropAllPos.End - 1)
               With .ControlsPos
                  .End = pos + Len(ksEnd) - 1
                  .Length = .End - .Start + 1
                  'Debug.Assert .Length > 0
               End With '.CONTROLSPOS
            End If
         End If
      Else ' .ModulePropPos.Start <= 0'NOT .MODULEPROPPOS.START...
         pos = InStr(1, strHead, ksObject & " =")
         If pos > 1 Then
            If .AttributesPos.Start > 0 Then
               .ObjectsPos.End = .AttributesPos.Start - 1
               With .ObjectsPos
                  .Start = pos
                  .Length = .End - .Start + 1
                  'Debug.Assert .Length > 0
               End With '.OBJECTSPOS
            Else 'NOT .ATTRIBUTESPOS.START...
               With .ObjectsPos
                  .Start = pos
                  .End = Len(strHead)
                  .Length = .End - .Start + 1
                  'Debug.Assert .Length > 0
               End With '.OBJECTSPOS
            End If
         End If
      End If

      If .ObjectsPos.Start > 0 Then
         pos = InStrRev(strHead, vbCrLf, .ObjectsPos.Start)
      ElseIf .ModulePropPos.Start > 0 Then 'NOT .OBJECTSPOS.START...
         pos = InStrRev(strHead, vbCrLf, .ModulePropPos.Start)
      ElseIf .AttributesPos.Start > 0 Then 'NOT .MODULEPROPPOS.START...
         pos = InStrRev(strHead, vbCrLf, .AttributesPos.Start)
      End If
      If pos > 1 Then
         With .HeadingPos
            .Start = 1
            .End = pos - 1
            .Length = .End - .Start + 1
            'Debug.Assert .Length > 0
         End With '.HEADINGPOS
      End If
   End With 'VBGETMODULEHEADPOSITIONS

End Function

Public Function After( _
                      ByVal Source As String, _
                      ByVal Target As String, _
                      Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                      Optional Return_pos As Long, _
                      Optional bReturnWholeTextIfNotFound As Boolean, _
                      Optional ByRef retError As Long) As String

'Returns the text after the target.

   On Error GoTo Err_Handler

   Return_pos = VBA.Instr(1, Source, Target, Compare)
   If Return_pos <= 0 Then
      If bReturnWholeTextIfNotFound Then
         After = Source
      Else 'BRETURNWHOLETEXTIFNOTFOUND = FALSE
         After = vbNullString
      End If
   Else 'NOT RETURN_POS...
      After = VBA.Mid$(Source, Return_pos + Len(Target))
   End If

   retError = 0

   Exit Function

Err_Handler:
   retError = -1

End Function

Public Function Before( _
                       ByVal Source As String, _
                       ByVal Target As String, _
                       Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                       Optional Return_pos As Long, _
                       Optional bReturnWholeTextIfNotFound As Boolean, _
                       Optional ByRef retError As Long) As String

'Returns the text before the target.

   On Error GoTo Err_Handler
   Return_pos = VBA.Instr(1, Source, Target, Compare)
   If Return_pos = 0 Then
      Before = IIf(bReturnWholeTextIfNotFound, Source, "")
   Else 'NOT RETURN_POS...
      Before = VBA.Left$(Source, Return_pos - 1)
   End If
   retError = 0

   Exit Function

Err_Handler:
   retError = -1

End Function

Public Function GetShortCut(ByVal Index As eShortCut) As String

'Returns the shortcut.

   Select Case Index
   Case 0 'SC_000_NONE
      Exit Function
   Case 1 To 26
      GetShortCut = "^" + Chr$(64 + Index)
   Case 27 To 38
      GetShortCut = "{F" + Format$(Index - 26) + "}"
   Case 39 To 50
      GetShortCut = "^{F" + Format$(Index - 38) + "}"
   Case 51 To 62
      GetShortCut = "+{F" + Format$(Index - 50) + "}"
   Case 63 To 74
      GetShortCut = "+^{F" + Format$(Index - 62) + "}"
   Case 75
      GetShortCut = "^{INSERT}"
   Case 76
      GetShortCut = "+{INSERT}"
   Case 77
      GetShortCut = "{DEL}"
   Case 78
      GetShortCut = "+{DEL}"
   Case 79
      GetShortCut = "%{BKSP}"
   End Select

End Function

Public Function GetShortCutDesc(ByVal Index As eShortCut) As String

'Returns descriptive text for the shortcut.

   Dim strShortCut As String
   Dim strDesc As String

   strShortCut = GetShortCut(Index)
   If LenB(strShortCut) = 0 Then
      Exit Function
   End If
   
   If InStr(1, strShortCut, "%") Then
      strDesc = strDesc & "Alt+"
   End If
   If InStr(1, strShortCut, "^") Then
      strDesc = strDesc & "Ctrl+"
   End If
   If InStr(1, strShortCut, "+") Then
      strDesc = strDesc & "Shift+"
   End If

   strDesc = strDesc & Between(strShortCut, "{", "}", , , , True)
   strDesc = Replace(strDesc, "^", "")
   GetShortCutDesc = strDesc '& "=" & Index & vbSpace4 & vbSQ & "Shortcut = " & strShortCut

End Function

Public Function GetShortCutText(VBMenuItemText As String) As String

'Extracts the shortcut text from a menu item text.

   If InStr(1, VBMenuItemText, "Shortcut") Then
      GetShortCutText = TrimCrLfTab(After(VBMenuItemText, "="))
      If InStr(1, GetShortCutText, vbSpace) Then
         GetShortCutText = Trim$(Before(GetShortCutText, vbSpace))
      End If
   End If

End Function

Public Function GetShortCutIndex(ShortCut As String) As eShortCut

'Returns short index from the given shortcut string.

   Dim x As Long
   Dim strTemp As String

   If LenB(ShortCut) = 0 Then
      Exit Function '>---> Bottom
   End If

   For x = 1 To 79
      strTemp = GetShortCut(x)
      If ShortCut = strTemp Then
         GetShortCutIndex = x
         Exit Function '>---> Bottom
      End If
   Next x

End Function

Public Function AfterRev( _
                         ByVal Source As String, ByVal Target As String, _
                         Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                         Optional Return_pos As Long, _
                         Optional bReturnWholeTextIfNotFound As Boolean) As String

'Searches the target backward on the source.
'If there is a match, returns the after text of the first match.

   Return_pos = VBA.InStrRev(Source, Target, -1, Compare)
   If Return_pos = 0 Then
      If bReturnWholeTextIfNotFound Then
         AfterRev = Source
      Else 'BRETURNWHOLETEXTIFNOTFOUND = FALSE
         AfterRev = ""
      End If
   Else 'NOT RETURN_POS...
      AfterRev = VBA.Mid$(Source, Return_pos + Len(Target))
   End If

End Function

Public Function StripQuotes(sPath As String, Optional bTrimCrLfTab As Boolean = True) As String

'Remove quotes from the given path (or any text).

   If bTrimCrLfTab Then
      StripQuotes = PathStripQuotes(TrimCrLfTab(sPath))
   Else 'BTRIMCRLFTAB = FALSE
      StripQuotes = PathStripQuotes(sPath)
   End If

End Function

Public Function PathStripQuotes(ByVal sPath As String) As String

   Call PathUnquoteSpacesA(sPath)
   PathStripQuotes = TrimNull(sPath)
   If InStr(1, PathStripQuotes, "'", vbTextCompare) = 1 Then
      PathStripQuotes = Between(PathStripQuotes, "'", "'", vbBinaryCompare, , , True)
   End If

End Function

Public Function ReplaceAll( _
                           ByVal sText As String, _
                           ByVal Replace As String, _
                           ParamArray Finds()) As String

'Replace all Finds() with Replace string. (complete replacing with Do..Loop)
   Dim i As Long

   For i = LBound(Finds) To UBound(Finds)
      Do While VBA.Instr(1, sText, Finds(i), vbTextCompare) > 0
         sText = VBA.Replace(sText, Finds(i), Replace, , , vbTextCompare)
         If VBA.Instr(1, Replace, Finds(i), vbTextCompare) > 0 Then
            Exit Do '>---> Loop
         End If
      Loop
   Next i

   ReplaceAll = sText

End Function

'*********************************************************************************************
' Draws a line with the specified color and width.
'*********************************************************************************************
Public Sub DrawLineEx( _
                      hDC As Long, _
                      BeginX As Long, BeginY As Long, _
                      EndX As Long, EndY As Long, _
                      lPenStyle As ePenStyle, _
                      lWidth As Long, _
                      lColor As eColor)

   Dim hPen As Long
   Dim hOldPen As Long

   'Create the pen
   hPen = CreatePen(lPenStyle, lWidth, lColor)

   'Select the pen onto the device context
   hOldPen = SelectObject(hDC, hPen)

   'Move the drawing position to utBegin
   MoveTo hDC, BeginX, BeginY, ByVal 0&

   'Draw the line
   LineTo hDC, EndX, EndY

   'Restore to the old pen
   Call SelectObject(hDC, hOldPen)

   'Delete the pen
   DeleteObject hPen

End Sub

Public Function TranslateColor(ByVal dwOleColour As Long) As eColor

'translate OLE color to valid color if passed

   OleTranslateColor dwOleColour, 0, TranslateColor

End Function


Public Function eShortCutDesc( _
   ByVal Index As eShortCut) As String
   Select Case Index
   Case SC_000_NONE:       eShortCutDesc = "SC_000_NONE"
   Case SC_CTRL_A:      eShortCutDesc = "SC_CTRL_A"
   Case SC_CTRL_B:      eShortCutDesc = "SC_CTRL_B"
   Case SC_CTRL_C:      eShortCutDesc = "SC_CTRL_C"
   Case SC_CTRL_D:      eShortCutDesc = "SC_CTRL_D"
   Case SC_CTRL_E:      eShortCutDesc = "SC_CTRL_E"
   Case SC_CTRL_F:      eShortCutDesc = "SC_CTRL_F"
   Case SC_CTRL_G:      eShortCutDesc = "SC_CTRL_G"
   Case SC_CTRL_H:      eShortCutDesc = "SC_CTRL_H"
   Case SC_CTRL_I:      eShortCutDesc = "SC_CTRL_I"
   Case SC_CTRL_J:      eShortCutDesc = "SC_CTRL_J"
   Case SC_CTRL_K:      eShortCutDesc = "SC_CTRL_K"
   Case SC_CTRL_L:      eShortCutDesc = "SC_CTRL_L"
   Case SC_CTRL_M:      eShortCutDesc = "SC_CTRL_M"
   Case SC_CTRL_N:      eShortCutDesc = "SC_CTRL_N"
   Case SC_CTRL_O:      eShortCutDesc = "SC_CTRL_O"
   Case SC_CTRL_P:      eShortCutDesc = "SC_CTRL_P"
   Case SC_CTRL_Q:      eShortCutDesc = "SC_CTRL_Q"
   Case SC_CTRL_R:      eShortCutDesc = "SC_CTRL_R"
   Case SC_CTRL_S:      eShortCutDesc = "SC_CTRL_S"
   Case SC_CTRL_T:      eShortCutDesc = "SC_CTRL_T"
   Case SC_CTRL_U:      eShortCutDesc = "SC_CTRL_U"
   Case SC_CTRL_V:      eShortCutDesc = "SC_CTRL_V"
   Case SC_CTRL_W:      eShortCutDesc = "SC_CTRL_W"
   Case SC_CTRL_X:      eShortCutDesc = "SC_CTRL_X"
   Case SC_CTRL_Y:      eShortCutDesc = "SC_CTRL_Y"
   Case SC_CTRL_Z:      eShortCutDesc = "SC_CTRL_Z"
   Case SC_F1:       eShortCutDesc = "SC_F1"
   Case SC_F2:       eShortCutDesc = "SC_F2"
   Case SC_F3:       eShortCutDesc = "SC_F3"
   Case SC_F4:       eShortCutDesc = "SC_F4"
   Case SC_F5:       eShortCutDesc = "SC_F5"
   Case SC_F6:       eShortCutDesc = "SC_F6"
   Case SC_F7:       eShortCutDesc = "SC_F7"
   Case SC_F8:       eShortCutDesc = "SC_F8"
   Case SC_F9:       eShortCutDesc = "SC_F9"
   Case SC_F10:      eShortCutDesc = "SC_F10"
   Case SC_F11:      eShortCutDesc = "SC_F11"
   Case SC_F12:      eShortCutDesc = "SC_F12"
   Case SC_CTRL_F1:        eShortCutDesc = "SC_CTRL_F1"
   Case SC_CTRL_F2:        eShortCutDesc = "SC_CTRL_F2"
   Case SC_CTRL_F3:        eShortCutDesc = "SC_CTRL_F3"
   Case SC_CTRL_F4:        eShortCutDesc = "SC_CTRL_F4"
   Case SC_CTRL_F5:        eShortCutDesc = "SC_CTRL_F5"
   Case SC_CTRL_F6:        eShortCutDesc = "SC_CTRL_F6"
   Case SC_CTRL_F7:        eShortCutDesc = "SC_CTRL_F7"
   Case SC_CTRL_F8:        eShortCutDesc = "SC_CTRL_F8"
   Case SC_CTRL_F9:        eShortCutDesc = "SC_CTRL_F9"
   Case SC_CTRL_F10:       eShortCutDesc = "SC_CTRL_F10"
   Case SC_CTRL_F11:       eShortCutDesc = "SC_CTRL_F11"
   Case SC_CTRL_F12:       eShortCutDesc = "SC_CTRL_F12"
   Case SC_SHIFT_F1:       eShortCutDesc = "SC_SHIFT_F1"
   Case SC_SHIFT_F2:       eShortCutDesc = "SC_SHIFT_F2"
   Case SC_SHIFT_F3:       eShortCutDesc = "SC_SHIFT_F3"
   Case SC_SHIFT_F4:       eShortCutDesc = "SC_SHIFT_F4"
   Case SC_SHIFT_F5:       eShortCutDesc = "SC_SHIFT_F5"
   Case SC_SHIFT_F6:       eShortCutDesc = "SC_SHIFT_F6"
   Case SC_SHIFT_F7:       eShortCutDesc = "SC_SHIFT_F7"
   Case SC_SHIFT_F8:       eShortCutDesc = "SC_SHIFT_F8"
   Case SC_SHIFT_F9:       eShortCutDesc = "SC_SHIFT_F9"
   Case SC_SHIFT_F10:      eShortCutDesc = "SC_SHIFT_F10"
   Case SC_SHIFT_F11:      eShortCutDesc = "SC_SHIFT_F11"
   Case SC_SHIFT_F12:      eShortCutDesc = "SC_SHIFT_F12"
   Case SC_CTRL_SHIFT_F1:        eShortCutDesc = "SC_CTRL_SHIFT_F1"
   Case SC_CTRL_SHIFT_F2:        eShortCutDesc = "SC_CTRL_SHIFT_F2"
   Case SC_CTRL_SHIFT_F3:        eShortCutDesc = "SC_CTRL_SHIFT_F3"
   Case SC_CTRL_SHIFT_F4:        eShortCutDesc = "SC_CTRL_SHIFT_F4"
   Case SC_CTRL_SHIFT_F5:        eShortCutDesc = "SC_CTRL_SHIFT_F5"
   Case SC_CTRL_SHIFT_F6:        eShortCutDesc = "SC_CTRL_SHIFT_F6"
   Case SC_CTRL_SHIFT_F7:        eShortCutDesc = "SC_CTRL_SHIFT_F7"
   Case SC_CTRL_SHIFT_F8:        eShortCutDesc = "SC_CTRL_SHIFT_F8"
   Case SC_CTRL_SHIFT_F9:        eShortCutDesc = "SC_CTRL_SHIFT_F9"
   Case SC_CTRL_SHIFT_F10:       eShortCutDesc = "SC_CTRL_SHIFT_F10"
   Case SC_CTRL_SHIFT_F11:       eShortCutDesc = "SC_CTRL_SHIFT_F11"
   Case SC_CTRL_SHIFT_F12:       eShortCutDesc = "SC_CTRL_SHIFT_F12"
   Case SC_CTRL_INSERT:       eShortCutDesc = "SC_CTRL_INSERT"
   Case SC_SHIFT_INSERT:      eShortCutDesc = "SC_SHIFT_INSERT"
   Case SC_DEL:      eShortCutDesc = "SC_DEL"
   Case SC_SHIFT_DEL:      eShortCutDesc = "SC_SHIFT_DEL"
   Case SC_ALT_BKSP:       eShortCutDesc = "SC_ALT_BKSP"
   Case Else
   End Select
End Function

Public Function Quoto(Text As String) As String
   Quoto = vbQ & Text & vbQ
End Function

'*********************************************************************************************
' Returns the caption of the menu item specifeid by ID for the owner window. Works well with WM_COMMAND.
'*********************************************************************************************
Public Function GetMenuCaption(ByVal hMainMenu As Long, ByVal nItemID As Long) As String
    GetMenuCaption = String$(255, 0)
    Call GetMenuString(hMainMenu, nItemID, GetMenuCaption, 255, MF_BY_COMMAND)
    GetMenuCaption = TrimNull(GetMenuCaption)
End Function

Public Function CreateMainPopupMenu(Optional InfoReturnMainMenuHandle As Long) As Long
    CreateMainPopupMenu = CreatePopupMenu()
End Function

Public Function PopupMenuEx( _
               hMainMenu As Long, _
               Optional ByVal OwnerHwnd As Long, _
               Optional ByVal x As Long, Optional ByVal y As Long, _
               Optional InfoReturnItemID As Long) As ERRSUCCESS
               
   Dim ptCursor As POINTAPI
   Dim lprc As RECT
   If OwnerHwnd = 0 Then
      OwnerHwnd = GetForegroundWindow()
   End If
   If x = 0 Or y = 0 Then
      Call GetCursorPos(ptCursor)
   Else
      ptCursor.x = x
      ptCursor.y = y
   End If
   PopupMenuEx = TrackPopupMenu(hMainMenu, _
                                   TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_TOPALIGN Or _
                                   TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RIGHTBUTTON, _
                                   ptCursor.x, ptCursor.y, 0, OwnerHwnd, lprc)
   If PopupMenuEx <> 0 Then
   'Processs item
   '  RaiseEvent ItemClick(hMainMenu, lItem)
   End If
End Function

Public Function PopupMenuCx( _
               hMainMenu As Long, _
               Optional ByVal OwnerHwnd As Long, _
               Optional InfoReturnItemID As Long) As ERRSUCCESS
               
   Dim ptCursor As POINTAPI
   Dim lprc As RECT
   If OwnerHwnd = 0 Then
      OwnerHwnd = GetForegroundWindow()
   End If
   Call GetCursorPos(ptCursor)
   PopupMenuCx = TrackPopupMenu(hMainMenu, _
                                   TPM_TOPALIGN Or _
                                   TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RIGHTBUTTON, _
                                   ptCursor.x, ptCursor.y, 0, OwnerHwnd, lprc)
   If PopupMenuCx <> 0 Then
   'Processs item
   '  RaiseEvent ItemClick(hMainMenu, lItem)
   End If
End Function

Public Function CreateSubMenu( _
               ByVal hParentMenu As Long, _
               ByVal Caption As String, _
               Optional ByVal ItemID As eSubMenu = IDM_FILE, _
               Optional ByVal ItemType As eMenuItemInfoType = MFT_STRING, _
               Optional ByVal ItemState As eMenuItemInfoState = MFS_ENABLED, _
               Optional ByVal ItemData As Long, _
               Optional InfoReturnSubMenuHandle As Long) As Long

   Dim lpMII As MENUITEMINFO
   Dim abCaption() As Byte

   With lpMII
      .cbSize = Len(lpMII) 'The size of this structure.
      'Which elements of the structure to use.
      .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU Or MIIM_DATA
      .fType = ItemType 'Ttype of item: a string.
      .fState = ItemState 'Enable/disable/check state
      .dwItemData = ItemData 'Set The ItemData
      .wID = ItemID 'Assign this item an item identifier.
      abCaption = StrConv(Caption & vbNullChar, vbFromUnicode)
      .dwTypeDataPtr = VarPtr(abCaption(0))
      .cch = UBound(abCaption) + 1
      .hSubMenu = CreatePopupMenu()
                           'We would set submenu to the handle of an existing popup
                           'to bind them together
      CreateSubMenu = .hSubMenu
   End With
   
   'Insert by item ID
    Call InsertMenuItem(hParentMenu, ItemID, bFALSE, lpMII)
End Function

Public Function AddSeparator( _
               ByVal hMainOrSubMenu As Long, _
               Optional ByVal ItemID As Long = 35000, _
               Optional InfoReturnGeneralSuccessValue As Long) As ERRSUCCESS
               
   AddSeparator = AddMenuItem(hMainOrSubMenu, "-", ItemID, SC_000_NONE, _
                           MFT_SEPARATOR, MFS_ENABLED, 0, InfoReturnGeneralSuccessValue)

End Function

Public Function AddMenuItem( _
               ByVal hMainOrSubMenu As Long, _
               ByVal Caption As String, _
               Optional ByVal ItemID As Long = 28729, _
               Optional ByVal ItemShortcut As eShortCut = SC_000_NONE, _
               Optional ByVal ItemType As eMenuItemInfoType = MFT_STRING, _
               Optional ByVal ItemState As eMenuItemInfoState = MFS_ENABLED, _
               Optional ByVal ItemData As Long, _
               Optional InfoReturnGeneralSuccessValue As Long) As ERRSUCCESS
               
   Dim lpMII As MENUITEMINFO
   Dim abCaption() As Byte
   Dim strShortCutDesc As String

   With lpMII
      .cbSize = Len(lpMII) 'The size of this structure.
      'Which elements of the structure to use.
      .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_DATA
      .fType = ItemType 'Ttype of item: a string.
      .fState = ItemState 'Enable/disable/check state
      .dwItemData = ItemData 'Set The ItemData
      .wID = ItemID 'Assign this item an item identifier.
      strShortCutDesc = GetShortCutDesc(ItemShortcut)
      If Len(strShortCutDesc) Then
         Caption = Caption & vbTab & strShortCutDesc
      End If
      abCaption = StrConv(Caption & vbNullChar, vbFromUnicode)
      .dwTypeDataPtr = VarPtr(abCaption(0))
      .cch = UBound(abCaption) + 1
      .hSubMenu = 0 'We would set submenu to the handle of an existing popup
                               'to bind them together
   End With
   
   AddMenuItem = InsertMenuItem(hMainOrSubMenu, ItemID, bFALSE, lpMII)
End Function


Public Function JoinCollection(ByVal colItems As Collection, _
   Optional Delimiter As String = vbCrLf, _
   Optional Prefix As String, _
   Optional Suffix As String) As String
   
   If ObjPtr(colItems) = 0 Then
      Exit Function
   End If
   
   Dim i As Long
   Dim asItems() As String
   
   With colItems
      If .Count = 0 Then
         Exit Function
      End If
      ReDim asItems(1 To .Count)
      For i = 1 To .Count
         asItems(i) = Prefix & .Item(i) & Suffix
      Next i
      JoinCollection = Join(asItems, Delimiter)
   End With
End Function


'Return -1 if we're running in the IDE or 0 if were running compiled.
Public Function InIDE() As Boolean
  Static Value As Long
  
  If Value = 0 Then
    Value = 1
    Debug.Assert True Or InIDE() 'This line won't exist in the compiled app
    InIDE = Value - 1
  End If

  Value = 0
End Function
