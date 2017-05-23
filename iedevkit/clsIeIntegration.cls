VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIeIntegration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CoCreateGuid Lib "ole32.dll" (pGUID As guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrCLSID As Long, ByVal cbMax As Long) As Long

Private Type guid
    GUID1 As Long
    GUID2 As Long
    GUID3 As Long
    GUID4 As Long
End Type

Enum MenuShowWhen
   swDefault = 1
   swImage = 2
   swControl = 4
   swTable = 8
   swTextSelect = 16
   swAnchor = 32
   swUnknown = 64
End Enum

Dim clsReg As clsRegistry2

'[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MenuExt\<menu name as folder>]
'"HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\" '& guid

Const baseMenuPath = "\Software\Microsoft\Internet Explorer\MenuExt\"
Const baseToolbarPth = "\Software\Microsoft\Internet Explorer\Extensions\"
 
Private Sub Class_Initialize()
    Set clsReg = New clsRegistry2
End Sub

Private Sub Class_Terminate()
    Set clsReg = Nothing
End Sub

Sub AddToRightClickMenu(menuText As String, _
                             sFilePath As String, _
                             Optional showWhen As MenuShowWhen = swDefault, _
                             Optional showWindow As Boolean = False)

        clsReg.Hive = HKEY_CURRENT_USER
        
        If Not clsReg.keyExists(baseMenuPath) Then
            clsReg.CreateKey baseMenuPath
        End If
        
        If Not clsReg.keyExists(baseMenuPath & menuText) Then
            clsReg.CreateKey baseMenuPath & menuText
        End If

        clsReg.SetValue baseMenuPath & menuText, "", sFilePath
        clsReg.SetValue baseMenuPath & menuText, "Contexts", showWhen, REG_DWORD
         
        If showWindow Then
            clsReg.SetValue baseMenuPath & menuText, "Flags", 1, REG_DWORD
        End If

End Sub

Function RemoveFromRightClickMenu(menuText As String) As Boolean
         clsReg.Hive = HKEY_CURRENT_USER
         RemoveFromRightClickMenu = clsReg.DeleteKey(baseMenuPath & menuText)
End Function

Function ShowInstalledMenuExtensions() As String()
    clsReg.Hive = HKEY_CURRENT_USER
    ShowInstalledMenuExtensions = clsReg.EnumKeys(baseMenuPath)
End Function

Function IntegrateWithIEToolbar(buttonText As String, _
                                hotIconPath As String, _
                                defaultIconPath As String, _
                                exeOrScriptPath As String, _
                                Optional isVisible As Boolean = True _
                                ) As String
    Dim myGUid As String
    Dim myBase As String
    Dim isExe As Boolean
    
    myGUid = CreateGUID()
    myBase = baseToolbarPth & myGUid
                                
    clsReg.Hive = HKEY_LOCAL_MACHINE
                              
    If Not clsReg.keyExists(baseToolbarPth) Then
            clsReg.CreateKey baseToolbarPth
    End If
                                
    If Not clsReg.keyExists(myBase) Then
            clsReg.CreateKey myBase
    End If
        
    ModifyToolBarSetting myGUid, buttonText, hotIconPath, defaultIconPath, exeOrScriptPath, isVisible
    
    IntegrateWithIEToolbar = myGUid

End Function


Function RemoveIEToolBar(sGuid As String) As Boolean
    If Len(sGuid) = 0 Then Exit Function
    clsReg.Hive = HKEY_LOCAL_MACHINE
    RemoveIEToolBar = clsReg.DeleteKey(baseToolbarPth & sGuid)
End Function


Function ModifyToolBarSetting(sGuid As String, _
                              Optional buttonText As String, _
                              Optional hotIconPath As String, _
                              Optional defaultIconPath As String, _
                              Optional exeOrScriptPath As String, _
                              Optional isVisible As Boolean = True)
                                
    Dim myBase As String
    Dim isExe As Boolean
    
    myBase = baseToolbarPth & sGuid
    
    clsReg.Hive = HKEY_LOCAL_MACHINE
    clsReg.SetValue myBase, "Default Visible", CStr(IIf(isVisible, "Yes", "No"))
    
    If Len(buttonText) > 0 Then _
        clsReg.SetValue myBase, "ButtonText", buttonText
    
    If Len(hotIconPath) > 0 Then _
        clsReg.SetValue myBase, "HotIcon", hotIconPath
        
        clsReg.SetValue myBase, "Flags", 1, REG_DWORD
    
    If Len(defaultIconPath) > 0 Then _
        clsReg.SetValue myBase, "Icon", defaultIconPath
    
    clsReg.SetValue myBase, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
        
    If Len(exeOrScriptPath) > 0 Then
        isExe = True
        If GetExtension(exeOrScriptPath) Like "*ht*" Then isExe = False
        clsReg.SetValue myBase, IIf(isExe, "Exec", "Script"), exeOrScriptPath
    End If

End Function





Private Function CreateGUID() As String

    Dim udtGUID As guid
    Dim strGUID As String
    Dim bytGUID()  As Byte
    Const lLen = 40

    bytGUID = String(lLen, 0)
    CoCreateGuid udtGUID
    StringFromGUID2 udtGUID, VarPtr(bytGUID(0)), lLen
    strGUID = bytGUID
    CreateGUID = Left(strGUID, InStr(strGUID, Chr(0)) - 1)
    
End Function

Private Function GetExtension(path) As String
    Dim tmp() As String, ub As String
    On Error Resume Next
    
    If Len(path) = 0 Then Exit Function
    
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    
    If InStr(1, ub, ".") > 0 Then
       GetExtension = Mid(ub, InStrRev(ub, "."), Len(ub))
    Else
       GetExtension = ""
    End If
    
End Function

