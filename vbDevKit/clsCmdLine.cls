VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCmdLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cmdLine As String
Private cmdArgs As Collection
Private argCount As Integer

Property Get CommandLine() As String
    CommandLine = cmdLine
End Property

Property Let CommandLine(cmd As String)
    cmdLine = cmd
    ParseIt cmdLine
End Property

Private Sub Class_Initialize()
    ' If Not isRegistered And Not isInitalized Then TellThemAllAboutIt
    cmdLine = Command
    ParseIt cmdLine
End Sub

Function IsArgPresent(argName As String) As Boolean
    Dim i As Integer
    For i = 0 To argCount - 1
        If LCase(argName) = LCase(cmdArgs(i & "a")) Then
            IsArgPresent = True
            Exit Function
        End If
    Next
End Function

Function GetArgumentsToArray() As String()
    Dim i As Integer, ret() As String
    For i = 0 To argCount - 1
        push ret(), cmdArgs(i & "a")
    Next
    GetArgumentsToArray = ret()
End Function

Private Sub ParseIt(cmd)
    
    Dim inquotes As Boolean, inminus As Boolean, isword As Boolean
    Dim tmp As String, letter As String, nextlet As String
    Dim i As Integer
    
    Set cmdArgs = New Collection
    
    tmp = ""
    cmd = Replace(cmd, """", "'")
    
    For i = 1 To Len(cmd)
        letter = Mid(cmd, i, 1)
        nextlet = Mid(cmd, i + 1, 1)
        
        Select Case letter
          Case "-": inminus = True: isword = False
          Case " ":
                      inminus = False
                      If isword Then
                        isword = False
                        AddArgument tmp
                        tmp = ""
                      End If
          Case "'":
                     isword = False
                     If inquotes = True Then
                      inquotes = False
                      AddArgument tmp
                      tmp = ""
                     Else
                       inquotes = True
                     End If
        End Select
        
        If inminus And Not inquotes And letter <> "-" Then
           AddArgument letter
           'args() = push(args, letter)
        ElseIf inquotes And letter <> "'" Then
           tmp = tmp & letter
        ElseIf Not inminus And Not inquotes And letter <> "'" Then
           isword = True
           tmp = tmp & letter
           If i = Len(cmd) Then AddArgument tmp
        End If
    Next
End Sub

Private Sub AddArgument(arg As String)
    On Error Resume Next
    cmdArgs.Add arg, argCount & "a"
    argCount = argCount + 1
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
