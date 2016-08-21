VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'requires vbdevkit

Public isInstalled As Boolean
Public homeDir As String
Public fullVersion As String
Public curVersion As String
Public inPath As Boolean
Public isJDK As Boolean
Public isJRE As Boolean

Dim cmd As New CCmdOutput

Function CompileSource(jfile, Optional classPath = "", Optional ByRef exitCode As Long) As String
    
    If Not FileExists(CStr(jfile)) Then
        CompileSource = "Source file not found: " & jfile
        Exit Function
    End If
    
    If Not isInstalled Then
        CompileSource = "No Java JDK installation found"
        Exit Function
    End If
    
    If Not isJDK Then
        CompileSource = "You must install jdk"
        Exit Function
    End If
    
    Dim c As String, failed As Boolean, exe As String
    
    exe = homeDir & "javac.exe"
    If Len(classPath) > 0 Then c = c & " -cp """ & classPath & """"
    'c = c & "javac.exe """ & jfile & """"
    c = c & " "" & jfile & """""
    
    If Not cmd.GetCommandOutput(exe, c, CompileSource, exitCode) Then
        CompileSource = "Failed to compile"
        Exit Function
    End If
    
End Function

Function RunClassFile(classFile, Optional args = "", Optional ByRef sOutput As String, Optional ByRef exitCode As Long, Optional hideWindow As Boolean = True, Optional wait As Boolean = True) As Boolean
    
    If Not FileExists(CStr(classFile)) Then
        sOutput = "Class file not found: " & classFile
        Exit Function
    End If
    
    If Not isInstalled Then
        sOutput = "No Java installation found"
        Exit Function
    End If
    
    Dim c As String, exe As String
    
    exe = homeDir & "java.exe"
    c = c & " -cp """ & GetParentFolder(classFile) & """ """ & GetBaseName(CStr(classFile)) & """"
    If Len(args) > 0 Then c = c & " " & args
    
    If wait Then
        If Not cmd.GetCommandOutput(exe, c, sOutput, exitCode, hideWindow) Then
            sOutput = "Failed to launch java"
            Exit Function
        End If
        RunClassFile = True
    Else
        If Not cmd.LaunchProcess(exe, c, False, hideWindow, exitCode) Then
            sOutput = "Failed to launch java"
            Exit Function
        End If
        RunClassFile = True
        Exit Function
    End If
    
    sOutput = Replace(sOutput, vbCrLf, Chr(5))
    sOutput = Replace(sOutput, vbLf, vbCrLf)
    sOutput = Replace(sOutput, Chr(5), vbCrLf)
    
End Function

Private Sub Class_Initialize()

    Dim reg As Object 'New clsRegistry2
    Dim base As String, failed As Boolean, home As String, exe As String
    
    Const jdk = "Java Development Kit"
    Const jre = "Java Runtime Environment"
    
    On Error Resume Next
    Set reg = CreateObject("vbdevkit.clsRegistry2")
    
    If reg Is Nothing Then
        MsgBox "CJava requires vbdevkit installed", vbInformation
        Exit Sub
    End If
    
    base = "\SOFTWARE\JavaSoft\"
    reg.hive = HKEY_LOCAL_MACHINE
    
    If reg.keyExists(base & jdk) Then
        base = base & jdk
        isJDK = True
    ElseIf reg.keyExists(base & jre) Then
        base = base & jre
        isJRE = True
    Else
        Exit Sub
    End If
    
    curVersion = reg.ReadValue(base, "CurrentVersion")
    If Len(curVersion) = 0 Then Exit Sub
    
    home = reg.ReadValue(base & "\" & curVersion, "JavaHome")
    If Len(home) = 0 Then Exit Sub
    
    If Not FolderExists(home & "\bin") Then Exit Sub
    
    homeDir = home & "\bin\"
    If Not FileExists(homeDir & "\java.exe") Then Exit Sub
    
    isInstalled = True
    If cmd.GetCommandOutput(, "java -version", fullVersion, , True, True) Then
        inPath = True
    Else
        exe = home & "java.exe"
        If Not cmd.GetCommandOutput(exe, " -version", fullVersion, , , True) Then Exit Sub
    End If
    
    a = InStr(fullVersion, """")
    If a > 0 Then
        a = a + 1
        b = InStr(a + 1, fullVersion, """")
        If b > a Then
            fullVersion = Mid(fullVersion, a, b - a)
        End If
    End If

End Sub

Function GetJarOutput(jar As String, args As String, Optional ByRef sOutput As String, Optional ByRef exitCode As Long) As Boolean
    
    Dim c As String, exe As String
    Dim failed As Boolean

    If Not isInstalled Then
        sOutput = "No Java Installation detected"
        Exit Function
    End If
        
    exe = homeDir & "java.exe"
    c = " -jar """ & jar & """"
    If Len(args) > 0 Then c = c & " """ & args & """"
    
    If Not cmd.GetCommandOutput(exe, c, sOutput, exitCode) Then
        sOutput = "Failed to launch Jar" & vbCrLf & sOutput
    Else
        GetJarOutput = True
        sOutput = Replace(sOutput, vbCrLf, Chr(5))
        sOutput = Replace(sOutput, vbLf, vbCrLf)
        sOutput = Replace(sOutput, Chr(5), vbCrLf)
    End If
    
End Function


