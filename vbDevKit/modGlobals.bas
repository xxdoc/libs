Attribute VB_Name = "modGlobals"
Option Explicit

Global isRegistered As Boolean
Global isInitalized As Boolean

Sub TellThemAllAboutIt()
    
    isRegistered = True
    isInitalized = True
    Exit Sub
    
    On Error Resume Next
    Dim X, Y
    Const crackamsg = "Yes I know I know! I am Jimmy CrackCorn you see."
    
    If Not isInitalized Then
        X = GetSetting("iedevkit", "settings", "la", "")
        
        If Len(X) = 0 Then
            Y = 30
        Else
            Y = DateDiff("n", Now(), X)
        End If
        
        If Y > 15 Or Y < -15 Then
            SaveSetting "iedevkit", "settings", "la", Now()
            frmAbout.Show 1
            If Err.Number > 0 Then frmAbout.Show
        Else
            isInitalized = True
        End If
        
        If Err.Number > 0 Then frmAbout.Show
    End If
    
    
    
End Sub
