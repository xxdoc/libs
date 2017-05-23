VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFlowGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public blocks As New Collection

 Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1

Function ParseAsm(asm As String)

    Dim tmp() As String
    Dim data As String
    Dim block As New CBlock
    
    tmp = Split(asm, vbCrLf)
    
    block.name = "start"
    
    For Each x In tmp
            
        If Len(x) = 0 Then
            DoEvents
        ElseIf LastChar(x) = ":" Then  'its a label start of a new block
            
            block.data = data
            blocks.Add block, block.name
            block.targets.Add "next"
            'List1.AddItem "Found label, completing old block: " & block.name & " targets: " & block.GetTargets()
            
            Set block = New CBlock
            data = x & vbCrLf
            block.setName x
            'List1.AddItem "Starting new block for label: " & block.name
            
        ElseIf FirstChar(x) = "j" Then 'its the end of a block
            block.data = data & x & vbCrLf
            
            block.AddTargetFromLabel x
            If VBA.Left(x, 3) <> "jmp" Then block.targets.Add "next"
            
            'List1.AddItem "Completing block " & block.name & " targets: " & block.GetTargets()
            
            blocks.Add block, block.name
            
            data = Empty
            Set block = New CBlock
            block.name = "block_" & blocks.Count   'default name
        
        'elseif lastElemetn then close out last block?
        
        Else
            data = data & x & vbCrLf
        End If
    
    Next
    
End Function

Private Function LastChar(ByVal x) As String
    x = Trim(x)
    LastChar = Mid(x, Len(x))
End Function

Private Function FirstChar(x) As String
    FirstChar = Mid(x, 1, 1)
End Function

Sub GenerateNodeImage(block As CBlock, p As PictureBox)

    Dim msg As String
    
    msg = block.DisplayData
    
    p.Picture = LoadPicture()
    
    padding = 3
    
    p.ScaleMode = vbCharacters
    
    p.AutoRedraw = True
    
    p.FontName = "Courier"
    p.FontSize = 11
    
    w = p.TextWidth(msg)
    h = p.TextHeight(msg)
    
    p.ScaleMode = vbPixels

    p.Width = w + padding
    p.Height = h
    p.CurrentX = 10
    p.CurrentY = padding
    
    p.Print msg

'    Dim rct As RECT
'
'    With rct
'        .Left = 0
'        .Right = p.Width
'        .Top = 0
'        .Bottom = p.Height
'    End With
'
'    DrawText p.hDC, msg, -1, rct, 0 'DT_CENTER
 

    
End Sub
