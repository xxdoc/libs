VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public data As String
Public targets As New Collection
Public node As oGraph.oPicture


Function setName(x)
    name = Replace(x, ":", Empty)
End Function

Function AddTargetFromLabel(ByVal x)
    x = Trim(x)
    a = InStrRev(x, " ")
    x = Mid(x, a + 1)
    targets.Add x
End Function

Function GetTargets() As String
    For Each x In targets
        GetTargets = GetTargets & x & ", "
    Next
End Function

Function DisplayData() As String
    
    Dim tmp() As String
    Dim hasLabel As Boolean
    Dim longestLine As Long
    
    tmp = Split(data, vbCrLf)
    
    If LastChar(tmp(0)) = ":" Then hasLabel = True
        
    
    For i = 0 To UBound(tmp)
        If i = 0 And hasLabel = False Then tmp(i) = "  " & tmp(i)
        If i > 0 Then tmp(i) = "  " & tmp(i)
        If Len(tmp(i)) > longestLine Then longestLine = Len(tmp(i))
    Next
    
    buf = Space(longestLine)
    For i = 0 To UBound(tmp)
         tmp(i) = Left(tmp(i) & buf, longestLine)
    Next
    
    
    
    DisplayData = Join(tmp, vbCrLf)
    DisplayData = Replace(DisplayData, vbCrLf & vbCrLf, vbCrLf)
    
End Function

Function LastChar(ByVal x) As String
    x = Trim(x)
    LastChar = Mid(x, Len(x))
End Function

