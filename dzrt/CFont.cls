VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Name
Public Bold
Public Italic
Public size
Public Underline
Public Charset
Public Strikethrough
Public color

Public Property Get selected() As Boolean
    selected = (Len(Name) > 0)
End Property

Sub Clear()
    Name = Empty
    Bold = Empty
    Italic = Empty
    size = Empty
    Underline = Empty
    Charset = Empty
    Strikethrough = Empty
    color = Empty
End Sub

Sub LoadFromObj(obj As Object)
    
    On Error Resume Next
    
    Bold = obj.FontBold
    If Err.Number <> 0 Then
        Bold = obj.Font.Bold
        Err.Clear
    End If
    
    Italic = obj.FontItalic
    If Err.Number <> 0 Then
        Italic = obj.Font.Italic
        Err.Clear
    End If
    
    Underline = obj.FontUnderline
    If Err.Number <> 0 Then
        Underline = obj.Font.Underline
        Err.Clear
    End If

    Strikethrough = obj.FontStrikethru
    If Err.Number <> 0 Then
        Strikethrough = obj.Font.Strikethru
        Err.Clear
    End If

    color = obj.ForeColor
    If Err.Number <> 0 Then
        color = obj.Font.ForeColor
        Err.Clear
    End If
    
    Name = obj.fontname
    If Err.Number <> 0 Then
        Name = obj.Font.Name
        Err.Clear
    End If
    
    size = obj.FontSize
    If Err.Number <> 0 Then
        size = obj.Font.size
        Err.Clear
    End If
    
End Sub

Sub SaveToObj(obj As Object)

    On Error Resume Next
    
    obj.FontBold = Bold
    obj.Font.Bold = Bold
    
    obj.FontItalic = Italic
    obj.Font.Italic = Italic
    
    obj.FontUnderline = Underline
    obj.Font.Underline = Underline
    
    obj.FontStrikethru = Strikethrough
    obj.Font.Strikethru = Strikethrough
    
    obj.ForeColor = color
    obj.Font.ForeColor = color
    
    obj.fontname = Name
    obj.Font.Name = Name
    
    obj.FontSize = size
    obj.Font.size = size
    
End Sub
