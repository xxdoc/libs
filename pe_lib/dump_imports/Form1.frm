VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save Back"
      Height          =   330
      Left            =   6255
      TabIndex        =   4
      Top             =   630
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Strip lordpe"
      Height          =   330
      Left            =   4455
      TabIndex        =   3
      Top             =   630
      Width           =   1725
   End
   Begin VB.TextBox Text2 
      Height          =   4560
      Left            =   135
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   765
      Width           =   7440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dump Imports"
      Height          =   420
      Left            =   5985
      TabIndex        =   1
      Top             =   180
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "drag and drop"
      Top             =   270
      Width           =   5640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pe As New CPEEditor
Dim loadedFile As String

Private Sub Command1_Click()
    
    If Not pe.LoadFile(Text1) Then
        MsgBox "Failed to load " & pe.errMessage
        Exit Sub
    End If
    
    Dim c As Collection
    Dim m As CImport
    Dim tmp() As String
    
    Set c = pe.Imports.Modules
    For Each m In c
        For Each f In m.functions
            push tmp, f
        Next
    Next
    
    Text2 = Join(tmp, vbCrLf)
    
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Sub Command2_Click()
    X = Split(Text2, vbCrLf)
    For i = 0 To UBound(X)
        X(i) = Trim(X(i))
        a = InStrRev(X(i), " ")
        If a > 0 Then
            X(i) = Mid(X(i), a)
        End If
        X(i) = Trim(Replace(X(i), """", Empty))
    Next
    Text2 = Join(X, vbCrLf)
End Sub

Private Sub Command3_Click()
    If pe.FileExists(loadedFile) Then
        WriteFile loadedFile, Text2
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1 = Data.Files(1)
End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    loadedFile = Data.Files(1)
    Text2 = ReadFile(Data.Files(1))
End Sub




Function ReadFile(filename)
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function



Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub


