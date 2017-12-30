VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If False Then
Private AutoCompleteOptionNone, AutoCompleteOptionAppend, AutoCompleteOptionSuggest, AutoCompleteOptionSuggestAppend, AutoCompleteOptionSearch, AutoCompleteOptionFilterPrefixes, AutoCompleteOptionUseTab, AutoCompleteOptionUpDownKeyDropsList
Private AutoCompleteSourceNone, AutoCompleteSourceHistory, AutoCompleteSourceFileSystem, AutoCompleteSourceMRU, AutoCompleteSourceAll, AutoCompleteSourceCustomSource
Private AutoCompleteFileSystemOptionNone, AutoCompleteFileSystemOptionCurrentDir, AutoCompleteFileSystemOptionMyComputer, AutoCompleteFileSystemOptionDesktop, AutoCompleteFileSystemOptionFavorites, AutoCompleteFileSystemOptionFileSysOnly, AutoCompleteFileSystemOptionFileSysDirs
#End If
Private Const ACO_NONE As Long = &H0
Private Const ACO_AUTOSUGGEST As Long = &H1
Private Const ACO_AUTOAPPEND As Long = &H2
Private Const ACO_SEARCH As Long = &H4
Private Const ACO_FILTERPREFIXES As Long = &H8
Private Const ACO_USETAB As Long = &H10
Private Const ACO_UPDOWNKEYDROPSLIST As Long = &H20
Public Enum AutoCompleteOptionConstants
AutoCompleteOptionNone = ACO_NONE
AutoCompleteOptionAppend = ACO_AUTOAPPEND
AutoCompleteOptionSuggest = ACO_AUTOSUGGEST
AutoCompleteOptionSuggestAppend = (ACO_AUTOSUGGEST Or ACO_AUTOAPPEND)
AutoCompleteOptionSearch = ACO_SEARCH
AutoCompleteOptionFilterPrefixes = ACO_FILTERPREFIXES
AutoCompleteOptionUseTab = ACO_USETAB
AutoCompleteOptionUpDownKeyDropsList = ACO_UPDOWNKEYDROPSLIST
End Enum
Private Const ACLO_NONE As Long = 0
Private Const ACLO_CURRENTDIR As Long = 1
Private Const ACLO_MYCOMPUTER As Long = 2
Private Const ACLO_DESKTOP As Long = 4
Private Const ACLO_FAVORITES As Long = 8
Private Const ACLO_FILESYSONLY As Long = 16
Private Const ACLO_FILESYSDIRS As Long = 32
Public Enum AutoCompleteFileSystemOptionConstants
AutoCompleteFileSystemOptionNone = ACLO_NONE
AutoCompleteFileSystemOptionCurrentDir = ACLO_CURRENTDIR
AutoCompleteFileSystemOptionMyComputer = ACLO_MYCOMPUTER
AutoCompleteFileSystemOptionDesktop = ACLO_DESKTOP
AutoCompleteFileSystemOptionFavorites = ACLO_FAVORITES
AutoCompleteFileSystemOptionFileSysOnly = ACLO_FILESYSONLY
AutoCompleteFileSystemOptionFileSysDirs = ACLO_FILESYSDIRS
End Enum
Public Enum AutoCompleteSourceConstants
AutoCompleteSourceNone = 0
AutoCompleteSourceHistory = 1
AutoCompleteSourceFileSystem = 2
AutoCompleteSourceMRU = 3
AutoCompleteSourceAll = 4
AutoCompleteSourceCustomSource = 5
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Const S_OK As Long = &H0, S_FALSE As Long = &H1
Private Const E_POINTER As Long = &H80004003
Private Const E_NOTIMPL As Long = &H80004001
Private Const ACDD_VISIBLE As Long = 1
Implements AutoCompleteGuids.IEnumString
Implements AutoCompleteGuids.IEnumStringVB
Private IEnumStringElements() As String
Private IEnumStringCurrentElement As Long
Private IEnumStringCount As Long
Private IEnumStringSubclassed As Boolean
Private IAutoComplete As AutoCompleteGuids.IAutoComplete2
Private IAutoCompleteDropDown As AutoCompleteGuids.IAutoCompleteDropDown
Private IAutoCompleteIACList2 As AutoCompleteGuids.IACList2
Private PropOptions As AutoCompleteOptionConstants
Private PropDisabled As Boolean

Private Sub Class_Initialize()
Set IAutoComplete = CreateIAutoComplete2()
End Sub

Private Sub Class_Terminate()
If IEnumStringSubclassed = True Then Call RemoveVTableSubclassIEnumString(Me)
End Sub

Private Sub IEnumString_Next(ByVal cElt As Long, ByVal rgElt As Long, ByVal pcEltFetched As Long)
End Sub

Private Sub IEnumString_Skip(ByVal cElt As Long)
End Sub

Private Sub IEnumString_Reset()
End Sub

Private Sub IEnumString_Clone(ByRef ppEnum As AutoCompleteGuids.IEnumString)
End Sub

Private Sub IEnumStringVB_Next(ByRef RetVal As Long, ByVal cElt As Long, ByVal rgElt As Long, ByVal pcEltFetched As Long)
If rgElt = 0 Then RetVal = E_POINTER: Exit Sub
Dim i As Long, cEltFetched As Long, lpString As Long
cEltFetched = 0
For i = 0 To cElt - 1
    If IEnumStringCurrentElement = IEnumStringCount Then Exit For
    lpString = AutoCompleteGuids.[_CoTaskMemAlloc](LenB(IEnumStringElements(IEnumStringCurrentElement) & vbNullChar))
    If lpString = 0 Then RetVal = S_FALSE: Exit Sub
    AutoCompleteGuids.[_CopyMemory] ByVal lpString, ByVal StrPtr(IEnumStringElements(IEnumStringCurrentElement)), LenB(IEnumStringElements(IEnumStringCurrentElement) & vbNullChar)
    AutoCompleteGuids.[_CopyMemory] ByVal UnsignedAdd(rgElt, i * 4), lpString, 4
    IEnumStringCurrentElement = IEnumStringCurrentElement + 1
    cEltFetched = cEltFetched + 1
Next i
If pcEltFetched <> 0 Then AutoCompleteGuids.[_CopyMemory] ByVal pcEltFetched, cEltFetched, 4
If i = cElt Then
    RetVal = S_OK
Else
    RetVal = S_FALSE
End If
End Sub

Private Sub IEnumStringVB_Skip(ByRef RetVal As Long, ByVal cElt As Long)
If IEnumStringCurrentElement + cElt <= IEnumStringCount Then
    IEnumStringCurrentElement = IEnumStringCurrentElement + cElt
    RetVal = S_OK
Else
    IEnumStringCurrentElement = IEnumStringCount
    RetVal = S_FALSE
End If
End Sub

Private Sub IEnumStringVB_Reset(ByRef RetVal As Long)
IEnumStringCurrentElement = 0
RetVal = S_OK
End Sub

Private Sub IEnumStringVB_Clone(ByRef RetVal As Long, ByVal ppEnum As Long)
RetVal = E_NOTIMPL
End Sub

Public Sub Init(ByVal hWndEdit As Long, ByVal Source As AutoCompleteSourceConstants)
Attribute Init.VB_Description = "Initializes the autocomplete object."
If IAutoComplete Is Nothing Then Exit Sub
Dim IACLHistory As IUnknown
Dim IACListISF As IUnknown
Dim IACLMRU As IUnknown
Select Case Source
    Case AutoCompleteSourceHistory
        Set IACLHistory = CreateIACLHistory()
        IAutoComplete.Init hWndEdit, IACLHistory, 0, 0
    Case AutoCompleteSourceFileSystem
        If IAutoCompleteIACList2 Is Nothing Then
            Set IACListISF = CreateIACListISF()
            Set IAutoCompleteIACList2 = IACListISF
        End If
        IAutoComplete.Init hWndEdit, IAutoCompleteIACList2, 0, 0
    Case AutoCompleteSourceMRU
        Set IACLMRU = CreateIACLMRU()
        IAutoComplete.Init hWndEdit, IACLMRU, 0, 0
    Case AutoCompleteSourceAll
        Set IACLHistory = CreateIACLHistory()
        If IAutoCompleteIACList2 Is Nothing Then
            Set IACListISF = CreateIACListISF()
            Set IAutoCompleteIACList2 = IACListISF
        End If
        Set IACLMRU = CreateIACLMRU()
        Dim IObjMgr As AutoCompleteGuids.IObjMgr
        Set IObjMgr = CreateIObjMgr()
        IObjMgr.Append IACLHistory
        IObjMgr.Append IAutoCompleteIACList2
        IObjMgr.Append IACLMRU
        IAutoComplete.Init hWndEdit, IObjMgr, 0, 0
    Case AutoCompleteSourceCustomSource
        Call SetVTableSubclassIEnumString(Me)
        IAutoComplete.Init hWndEdit, Me, 0, 0
        IEnumStringSubclassed = True
    Case Else
        Err.Raise 380
End Select
End Sub

Public Property Get Options() As AutoCompleteOptionConstants
Attribute Options.VB_Description = "Returns/sets one or more values that determines how the automatic completion works."
If Not IAutoComplete Is Nothing Then
    If PropDisabled = False Then
        IAutoComplete.GetOptions Options
    Else
        Options = PropOptions
    End If
End If
End Property

Public Property Let Options(ByVal Value As AutoCompleteOptionConstants)
If Not IAutoComplete Is Nothing Then
    PropOptions = Value
    If PropDisabled = False Then IAutoComplete.SetOptions PropOptions
End If
End Property

Public Property Get FileSystemOptions() As AutoCompleteFileSystemOptionConstants
Attribute FileSystemOptions.VB_Description = "Returns/sets one or more values that determines which objects are enumerated when the AutoCompleteSource property is set to FileSystem."
If IAutoCompleteIACList2 Is Nothing Then
    Dim IACListISF As IUnknown
    Set IACListISF = CreateIACListISF()
    Set IAutoCompleteIACList2 = IACListISF
End If
If Not IAutoCompleteIACList2 Is Nothing Then IAutoCompleteIACList2.GetOptions FileSystemOptions
End Property

Public Property Let FileSystemOptions(ByVal Value As AutoCompleteFileSystemOptionConstants)
If IAutoCompleteIACList2 Is Nothing Then
    Dim IACListISF As IUnknown
    Set IACListISF = CreateIACListISF()
    Set IAutoCompleteIACList2 = IACListISF
End If
If Not IAutoCompleteIACList2 Is Nothing Then IAutoCompleteIACList2.SetOptions Value
End Property

Public Property Get CustomSource() As Variant
Attribute CustomSource.VB_Description = "Returns/sets a custom string array to use when the AutoCompleteSource property is set to CustomSource."
CustomSource = IEnumStringElements()
End Property

Public Property Let CustomSource(ByVal StringArray As Variant)
If IsArray(StringArray) Then
    Dim Ptr As Long
    CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(StringArray), 8), 4
    If Ptr <> 0 Then
        Dim RetVal As Long
        CopyMemory ByVal VarPtr(RetVal), Ptr, 4
        If RetVal <> 0 Then
            Dim DimensionCount As Integer
            CopyMemory DimensionCount, ByVal Ptr, 2
            If DimensionCount = 1 Then
                Erase IEnumStringElements()
                IEnumStringCurrentElement = 0
                IEnumStringCount = 0
                Dim i As Long
                For i = LBound(StringArray) To UBound(StringArray)
                    If VarType(StringArray(i)) = vbString Then
                        If Not StringArray(i) = vbNullString Then
                            ReDim Preserve IEnumStringElements(0 To IEnumStringCount) As String
                            IEnumStringElements(IEnumStringCount) = StringArray(i)
                            IEnumStringCount = IEnumStringCount + 1
                        End If
                    End If
                Next i
            Else
                Err.Raise Number:=5, Description:="Array must be single dimensioned"
            End If
        Else
            Err.Raise Number:=91, Description:="Array is not allocated"
        End If
    Else
        Err.Raise 5
    End If
ElseIf IsEmpty(StringArray) Then
    Erase IEnumStringElements()
    IEnumStringCurrentElement = 0
    IEnumStringCount = 0
Else
    Err.Raise 380
End If
End Property

Public Function DroppedDown() As Boolean
Attribute DroppedDown.VB_Description = "Returns a value that determines whether the drop-down list is dropped down or not."
If IAutoCompleteDropDown Is Nothing Then Set IAutoCompleteDropDown = IAutoComplete
If Not IAutoCompleteDropDown Is Nothing Then
    Dim Flags As Long
    IAutoCompleteDropDown.GetDropDownStatus Flags, vbNullString
    DroppedDown = CBool((Flags And ACDD_VISIBLE) = ACDD_VISIBLE)
End If
End Function

Public Sub Enable()
Attribute Enable.VB_Description = "Enables the auto completion."
If Not IAutoComplete Is Nothing Then
    IAutoComplete.SetOptions PropOptions
    IAutoComplete.Enable 1
    PropDisabled = False
End If
End Sub

Public Sub Disable()
Attribute Disable.VB_Description = "Disables the auto completion."
If Not IAutoComplete Is Nothing Then
    IAutoComplete.SetOptions AutoCompleteOptionNone
    IAutoComplete.Enable 0
    PropDisabled = True
End If
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces the autocomplete object to refresh its drop-down list of suggestions when it is visible."
If IAutoCompleteDropDown Is Nothing Then Set IAutoCompleteDropDown = IAutoComplete
If Not IAutoCompleteDropDown Is Nothing Then IAutoCompleteDropDown.ResetEnumerator
End Sub
