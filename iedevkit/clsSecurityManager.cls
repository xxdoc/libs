VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurityManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oZoneManager As IInternetZoneManager

Private Const upDisallow = 3
Private Const upAllow = 0
Private Const URLZONEREG_DEFAULT = 0

Enum ZoneTemplates
   ztLow = &H10000
   ztMedium = &H11000
   ztHigh = &H2000
End Enum

Enum urlZone
    zLocalMachine = 0
    zIntranet = 1
    zTrusted = 2
    zInternet = 3
    zUntrusted = 4
End Enum

Enum urlActions
    RunActiveX = 4608
    OverrideSafeForScripting = 4611
    OverrideSafeObjects = 4609
    OverrideDataSafety = 4610
    AllowCookies = 6658
    AllowSessionCookies = 6659
    DownloadSignedActx = 4097
    DownloadUnSignedActx = 4100
    AllowJava = 5637
    AllowSubmitForms = 5633
    AllowUserdataPersistance = 5638
    AllowScripting = 5120
    AllowScriptingActx = 5125
End Enum
    
    
Private Sub Class_Initialize()
    CoInternetCreateZoneManager Nothing, oZoneManager, 0&
End Sub

Sub ShowInternetOptions()
    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", vbNormalFocus
End Sub

Sub SetPolicyToTemlpate(template As ZoneTemplates, Optional Zone As urlZone = zInternet)
    oZoneManager.CopyTemplatePoliciesToZone template, Zone
End Sub

Function IsPolicyActive(p As urlActions, Optional Zone As urlZone = zInternet) As Boolean
    Dim lResult As Long
    oZoneManager.GetZoneActionPolicy Zone, p, lResult, 4&, URLZONEREG_DEFAULT
    IsPolicyActive = IIf(lResult = upDisallow, False, True)
End Function

Sub SetPolicy(p As urlActions, active As Boolean, Optional Zone As urlZone = zInternet)
    Dim lResult As Long
    lResult = IIf(active, upAllow, upDisallow)
    oZoneManager.SetZoneActionPolicy Zone, p, lResult, 4&, URLZONEREG_DEFAULT
End Sub

